from __future__ import annotations

import json
import pickle
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime

import faiss
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer

from .config import AppConfig, get_default_config
from .parsing import ChunkRecord, discover_source_files, load_source_documents, split_into_chunks
from .retrieval import (
    classify_chunk_group,
    embed_texts,
    extract_priority_legal_terms,
    extract_query_terms,
)


@dataclass
class StoreStats:
    documents: int
    chunks: int
    sources: list[str]


@dataclass
class HistoryEntry:
    id: int
    session_id: str
    turn_id: int
    memory_group_id: int
    memory_keywords: list[str]
    question: str
    answer: str
    thinking: str
    question_segments: list[dict[str, object]]
    answer_segments: list[dict[str, object]]
    citations: list[dict[str, object] | str]
    llm_used: bool
    llm_error: str
    created_at: str


class LegalRAGStore:
    def __init__(self, config: AppConfig | None = None):
        self.config = config or get_default_config()
        self.config.runtime_dir.mkdir(parents=True, exist_ok=True)
        self._initialize_database()

    def rebuild(self) -> StoreStats:
        source_files = discover_source_files(
            self.config.source_roots,
            self.config.excluded_dir_names,
            self.config.supported_extensions,
        )
        documents = []
        skipped_sources = []
        for path in source_files:
            try:
                documents.extend(load_source_documents(path))
            except Exception as exc:
                skipped_sources.append({"path": str(path), "error": str(exc)})

        chunks: list[ChunkRecord] = []
        for document in documents:
            chunks.extend(
                split_into_chunks(
                    document,
                    chunk_size=self.config.chunk_size,
                    overlap=self.config.chunk_overlap,
                )
            )

        self._replace_database(documents, chunks)
        self._build_indexes(chunks)

        manifest = {
            "source_roots": [str(path) for path in self.config.source_roots],
            "document_count": len(documents),
            "chunk_count": len(chunks),
            "sources": [document.source_name for document in documents],
            "skipped_sources": skipped_sources,
        }
        self.config.manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return StoreStats(
            documents=len(documents),
            chunks=len(chunks),
            sources=[document.source_name for document in documents],
        )

    def get_stats(self) -> StoreStats:
        with self._connect() as conn:
            doc_count = conn.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
            chunk_count = conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
            rows = conn.execute(
                "SELECT source_name FROM documents ORDER BY source_name"
            ).fetchall()
        return StoreStats(
            documents=doc_count,
            chunks=chunk_count,
            sources=[row[0] for row in rows],
        )

    def fetch_chunks(self) -> list[dict]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    c.chunk_id,
                    c.chunk_index,
                    c.content,
                    c.metadata_json,
                    d.source_name,
                    d.source_path,
                    d.title,
                    d.file_type
                FROM chunks c
                JOIN documents d ON d.id = c.document_id
                ORDER BY d.source_name, c.chunk_index
                """
            ).fetchall()
        return [
            {
                "chunk_id": row[0],
                "chunk_index": row[1],
                "text": row[2],
                "metadata": json.loads(row[3] or "{}"),
                "source_name": row[4],
                "source_path": row[5],
                "title": row[6],
                "file_type": row[7],
            }
            for row in rows
        ]

    def _initialize_database(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                PRAGMA foreign_keys = ON;

                CREATE TABLE IF NOT EXISTS documents (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_name TEXT NOT NULL UNIQUE,
                    source_path TEXT NOT NULL,
                    title TEXT NOT NULL,
                    file_type TEXT NOT NULL,
                    checksum TEXT NOT NULL,
                    char_count INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS chunks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    chunk_id TEXT NOT NULL UNIQUE,
                    document_id INTEGER NOT NULL,
                    chunk_index INTEGER NOT NULL,
                    content TEXT NOT NULL,
                    char_count INTEGER NOT NULL,
                    metadata_json TEXT NOT NULL,
                    FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS chat_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT NOT NULL DEFAULT 'legacy',
                    turn_id INTEGER NOT NULL DEFAULT 0,
                    memory_group_id INTEGER NOT NULL DEFAULT 0,
                    memory_keywords_json TEXT NOT NULL DEFAULT '[]',
                    question TEXT NOT NULL,
                    answer TEXT NOT NULL,
                    thinking_text TEXT NOT NULL DEFAULT '',
                    question_segments_json TEXT NOT NULL DEFAULT '[]',
                    answer_segments_json TEXT NOT NULL DEFAULT '[]',
                    citations_json TEXT NOT NULL,
                    llm_used INTEGER NOT NULL DEFAULT 0,
                    llm_error TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS chat_memory_groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT NOT NULL,
                    group_label TEXT NOT NULL,
                    keywords_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_chat_memory_groups_session
                ON chat_memory_groups(session_id);
                """
            )
            self._ensure_chat_history_columns(conn)

    def save_history_entry(
        self,
        session_id: str,
        question: str,
        answer: str,
        thinking: str,
        citations: list[dict[str, object] | str],
        llm_used: bool,
        llm_error: str,
    ) -> int:
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        question_segments = self._build_text_segments(question)
        answer_segments = self._build_text_segments(answer)
        memory_keywords = self._extract_memory_keywords(question, answer)
        with self._connect() as conn:
            turn_id = self._next_turn_id(conn, session_id)
            memory_group_id, memory_keywords = self._resolve_memory_group(
                conn=conn,
                session_id=session_id,
                entry_keywords=memory_keywords,
                created_at=created_at,
            )
            cursor = conn.execute(
                """
                INSERT INTO chat_history (
                    session_id,
                    turn_id,
                    memory_group_id,
                    memory_keywords_json,
                    question,
                    answer,
                    thinking_text,
                    question_segments_json,
                    answer_segments_json,
                    citations_json,
                    llm_used,
                    llm_error,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    turn_id,
                    memory_group_id,
                    json.dumps(memory_keywords, ensure_ascii=False),
                    question,
                    answer,
                    thinking,
                    json.dumps(question_segments, ensure_ascii=False),
                    json.dumps(answer_segments, ensure_ascii=False),
                    json.dumps(citations, ensure_ascii=False),
                    1 if llm_used else 0,
                    llm_error,
                    created_at,
                ),
            )
            return int(cursor.lastrowid)

    def list_history_entries(
        self,
        limit: int = 20,
        keyword: str = "",
    ) -> list[HistoryEntry]:
        keyword = " ".join(keyword.split()).strip()
        where_clause = ""
        params: list[object] = []
        if keyword:
            where_clause = "WHERE question LIKE ? OR answer LIKE ?"
            like_keyword = f"%{keyword}%"
            params.extend([like_keyword, like_keyword])
        params.append(limit)
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT
                    id,
                    session_id,
                    turn_id,
                    memory_group_id,
                    memory_keywords_json,
                    question,
                    answer,
                    thinking_text,
                    question_segments_json,
                    answer_segments_json,
                    citations_json,
                    llm_used,
                    llm_error,
                    created_at
                FROM chat_history
                {where_clause}
                ORDER BY id DESC
                LIMIT ?
                """,
                tuple(params),
            ).fetchall()
        return [
            HistoryEntry(
                id=int(row[0]),
                session_id=row[1] or "legacy",
                turn_id=int(row[2] or 0),
                memory_group_id=int(row[3] or 0),
                memory_keywords=json.loads(row[4] or "[]"),
                question=row[5],
                answer=row[6],
                thinking=row[7] or "",
                question_segments=json.loads(row[8] or "[]"),
                answer_segments=json.loads(row[9] or "[]"),
                citations=json.loads(row[10] or "[]"),
                llm_used=bool(row[11]),
                llm_error=row[12] or "",
                created_at=row[13],
            )
            for row in rows
        ]

    def get_history_entry(self, entry_id: int) -> HistoryEntry | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    id,
                    session_id,
                    turn_id,
                    memory_group_id,
                    memory_keywords_json,
                    question,
                    answer,
                    thinking_text,
                    question_segments_json,
                    answer_segments_json,
                    citations_json,
                    llm_used,
                    llm_error,
                    created_at
                FROM chat_history
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
        if row is None:
            return None
        return HistoryEntry(
            id=int(row[0]),
            session_id=row[1] or "legacy",
            turn_id=int(row[2] or 0),
            memory_group_id=int(row[3] or 0),
            memory_keywords=json.loads(row[4] or "[]"),
            question=row[5],
            answer=row[6],
            thinking=row[7] or "",
            question_segments=json.loads(row[8] or "[]"),
            answer_segments=json.loads(row[9] or "[]"),
            citations=json.loads(row[10] or "[]"),
            llm_used=bool(row[11]),
            llm_error=row[12] or "",
            created_at=row[13],
        )

    def list_session_entries(self, session_id: str) -> list[HistoryEntry]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    id,
                    session_id,
                    turn_id,
                    memory_group_id,
                    memory_keywords_json,
                    question,
                    answer,
                    thinking_text,
                    question_segments_json,
                    answer_segments_json,
                    citations_json,
                    llm_used,
                    llm_error,
                    created_at
                FROM chat_history
                WHERE session_id = ?
                ORDER BY id ASC
                """,
                (session_id,),
            ).fetchall()
        return [
            HistoryEntry(
                id=int(row[0]),
                session_id=row[1] or "legacy",
                turn_id=int(row[2] or 0),
                memory_group_id=int(row[3] or 0),
                memory_keywords=json.loads(row[4] or "[]"),
                question=row[5],
                answer=row[6],
                thinking=row[7] or "",
                question_segments=json.loads(row[8] or "[]"),
                answer_segments=json.loads(row[9] or "[]"),
                citations=json.loads(row[10] or "[]"),
                llm_used=bool(row[11]),
                llm_error=row[12] or "",
                created_at=row[13],
            )
            for row in rows
        ]

    def fetch_memory_entries(
        self,
        session_id: str,
        limit: int = 30,
    ) -> list[dict[str, object]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    id,
                    question,
                    answer,
                    created_at
                FROM chat_history
                WHERE session_id = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (session_id, limit),
            ).fetchall()
        rows = list(reversed(rows))
        return [
            {
                "id": int(row[0]),
                "session_id": session_id,
                "question": row[1],
                "answer": row[2],
                "created_at": row[3],
                "text": f"问题：{row[1]}\n回答：{row[2]}",
            }
            for row in rows
        ]

    def delete_history_entry(self, entry_id: int) -> bool:
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM chat_history WHERE id = ?",
                (entry_id,),
            )
            LegalRAGStore._purge_orphan_memory_groups(conn)
            return cursor.rowcount > 0

    def clear_history_entries(self, keyword: str = "") -> int:
        keyword = " ".join(keyword.split()).strip()
        with self._connect() as conn:
            if keyword:
                like_keyword = f"%{keyword}%"
                cursor = conn.execute(
                    """
                    DELETE FROM chat_history
                    WHERE question LIKE ? OR answer LIKE ?
                    """,
                    (like_keyword, like_keyword),
                )
            else:
                cursor = conn.execute("DELETE FROM chat_history")
            LegalRAGStore._purge_orphan_memory_groups(conn)
            return cursor.rowcount

    @staticmethod
    def _ensure_chat_history_columns(conn: sqlite3.Connection) -> None:
        LegalRAGStore._ensure_memory_group_table(conn)
        columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(chat_history)").fetchall()
        }
        if "session_id" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN session_id TEXT NOT NULL DEFAULT 'legacy'"
            )
        if "thinking_text" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN thinking_text TEXT NOT NULL DEFAULT ''"
            )
        if "turn_id" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN turn_id INTEGER NOT NULL DEFAULT 0"
            )
        if "memory_group_id" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN memory_group_id INTEGER NOT NULL DEFAULT 0"
            )
        if "memory_keywords_json" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN memory_keywords_json TEXT NOT NULL DEFAULT '[]'"
            )
        if "question_segments_json" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN question_segments_json TEXT NOT NULL DEFAULT '[]'"
            )
        if "answer_segments_json" not in columns:
            conn.execute(
                "ALTER TABLE chat_history ADD COLUMN answer_segments_json TEXT NOT NULL DEFAULT '[]'"
            )
        conn.execute(
            """
            UPDATE chat_history
            SET session_id = 'legacy-' || id
            WHERE session_id = 'legacy'
            """
        )
        LegalRAGStore._backfill_turn_ids(conn)
        LegalRAGStore._backfill_segments(conn)
        LegalRAGStore._backfill_memory_groups(conn)
        LegalRAGStore._purge_orphan_memory_groups(conn)

    @staticmethod
    def _ensure_memory_group_table(conn: sqlite3.Connection) -> None:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS chat_memory_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                group_label TEXT NOT NULL,
                keywords_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_chat_memory_groups_session
            ON chat_memory_groups(session_id);
            """
        )

    @staticmethod
    def _next_turn_id(conn: sqlite3.Connection, session_id: str) -> int:
        row = conn.execute(
            "SELECT COALESCE(MAX(turn_id), 0) FROM chat_history WHERE session_id = ?",
            (session_id,),
        ).fetchone()
        return int(row[0]) + 1

    @staticmethod
    def _build_text_segments(text: str) -> list[dict[str, object]]:
        compact = " ".join(text.split()).strip()
        if not compact:
            return []
        parts = [segment.strip() for segment in re.split(r"(?<=[。！？!?；;\n])", compact) if segment.strip()]
        return [
            {
                "seq_id": index,
                "text": segment,
            }
            for index, segment in enumerate(parts, start=1)
        ]

    @staticmethod
    def _extract_memory_keywords(question: str, answer: str, limit: int = 16) -> list[str]:
        content = f"{question}\n{answer}".strip()
        if not content:
            raise ValueError("History entry content is empty; cannot extract memory keywords.")
        tokens: list[str] = []
        tokens.extend(extract_priority_legal_terms(content))
        query_terms = sorted(
            extract_query_terms(content),
            key=len,
            reverse=True,
        )
        tokens.extend(term for term in query_terms if len(term) >= 2)
        deduped: list[str] = []
        seen: set[str] = set()
        for token in tokens:
            normalized = token.strip().lower()
            if len(normalized) < 2:
                continue
            if normalized in seen:
                continue
            seen.add(normalized)
            deduped.append(token.strip())
            if len(deduped) >= limit:
                break
        if not deduped:
            raise ValueError("Failed to extract memory keywords from history entry.")
        return deduped

    @staticmethod
    def _resolve_memory_group(
        conn: sqlite3.Connection,
        session_id: str,
        entry_keywords: list[str],
        created_at: str,
    ) -> tuple[int, list[str]]:
        keyword_set = {token.strip().lower() for token in entry_keywords if token.strip()}
        if not keyword_set:
            raise ValueError("Entry keywords are empty; cannot resolve memory group.")
        rows = conn.execute(
            """
            SELECT id, keywords_json, created_at
            FROM chat_memory_groups
            WHERE session_id = ?
            ORDER BY id ASC
            """,
            (session_id,),
        ).fetchall()
        best_group_id = 0
        best_group_keywords: list[str] = []
        best_score = 0.0
        for row in rows:
            group_id = int(row[0])
            group_keywords = json.loads(row[1] or "[]")
            group_set = {token.strip().lower() for token in group_keywords if str(token).strip()}
            if not group_set:
                continue
            overlap = len(group_set & keyword_set)
            if overlap == 0:
                continue
            union = len(group_set | keyword_set)
            score = overlap / max(union, 1)
            score += min(overlap, 6) * 0.08
            if score > best_score:
                best_score = score
                best_group_id = group_id
                best_group_keywords = [str(token) for token in group_keywords if str(token).strip()]
        if best_group_id and best_score >= 0.24:
            merged_keywords = LegalRAGStore._merge_memory_keywords(best_group_keywords, entry_keywords)
            conn.execute(
                """
                UPDATE chat_memory_groups
                SET keywords_json = ?,
                    group_label = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    json.dumps(merged_keywords, ensure_ascii=False),
                    LegalRAGStore._build_memory_group_label(merged_keywords),
                    created_at,
                    best_group_id,
                ),
            )
            return best_group_id, merged_keywords
        group_keywords = LegalRAGStore._merge_memory_keywords([], entry_keywords)
        cursor = conn.execute(
            """
            INSERT INTO chat_memory_groups (
                session_id,
                group_label,
                keywords_json,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                session_id,
                LegalRAGStore._build_memory_group_label(group_keywords),
                json.dumps(group_keywords, ensure_ascii=False),
                created_at,
                created_at,
            ),
        )
        return int(cursor.lastrowid), group_keywords

    @staticmethod
    def _merge_memory_keywords(base_keywords: list[str], new_keywords: list[str], limit: int = 18) -> list[str]:
        merged: list[str] = []
        seen: set[str] = set()
        for token in list(base_keywords) + list(new_keywords):
            normalized = token.strip().lower()
            if len(normalized) < 2:
                continue
            if normalized in seen:
                continue
            seen.add(normalized)
            merged.append(token.strip())
            if len(merged) >= limit:
                break
        if not merged:
            raise ValueError("Merged memory keywords are empty.")
        return merged

    @staticmethod
    def _build_memory_group_label(keywords: list[str]) -> str:
        if not keywords:
            raise ValueError("Cannot build memory group label from empty keywords.")
        return " / ".join(keywords[:3])

    @staticmethod
    def _backfill_turn_ids(conn: sqlite3.Connection) -> None:
        sessions = conn.execute("SELECT DISTINCT session_id FROM chat_history").fetchall()
        for (session_id,) in sessions:
            rows = conn.execute(
                "SELECT id FROM chat_history WHERE session_id = ? ORDER BY id ASC",
                (session_id,),
            ).fetchall()
            for index, (row_id,) in enumerate(rows, start=1):
                conn.execute(
                    "UPDATE chat_history SET turn_id = ? WHERE id = ?",
                    (index, row_id),
                )

    @staticmethod
    def _backfill_segments(conn: sqlite3.Connection) -> None:
        rows = conn.execute(
            """
            SELECT id, question, answer
            FROM chat_history
            WHERE question_segments_json = '[]' OR answer_segments_json = '[]'
            """
        ).fetchall()
        for row_id, question, answer in rows:
            question_segments = LegalRAGStore._build_text_segments(str(question))
            answer_segments = LegalRAGStore._build_text_segments(str(answer))
            conn.execute(
                """
                UPDATE chat_history
                SET question_segments_json = ?,
                    answer_segments_json = ?
                WHERE id = ?
                """,
                (
                    json.dumps(question_segments, ensure_ascii=False),
                    json.dumps(answer_segments, ensure_ascii=False),
                    row_id,
                ),
            )

    @staticmethod
    def _backfill_memory_groups(conn: sqlite3.Connection) -> None:
        rows = conn.execute(
            """
            SELECT
                id,
                session_id,
                question,
                answer,
                created_at
            FROM chat_history
            WHERE memory_group_id = 0 OR memory_keywords_json = '[]'
            ORDER BY session_id ASC, id ASC
            """
        ).fetchall()
        for row_id, session_id, question, answer, created_at in rows:
            keywords = LegalRAGStore._extract_memory_keywords(str(question), str(answer))
            group_id, resolved_keywords = LegalRAGStore._resolve_memory_group(
                conn=conn,
                session_id=str(session_id),
                entry_keywords=keywords,
                created_at=str(created_at),
            )
            conn.execute(
                """
                UPDATE chat_history
                SET memory_group_id = ?,
                    memory_keywords_json = ?
                WHERE id = ?
                """,
                (
                    group_id,
                    json.dumps(resolved_keywords, ensure_ascii=False),
                    row_id,
                ),
            )

    @staticmethod
    def _purge_orphan_memory_groups(conn: sqlite3.Connection) -> None:
        conn.execute(
            """
            DELETE FROM chat_memory_groups
            WHERE id NOT IN (
                SELECT DISTINCT memory_group_id
                FROM chat_history
                WHERE memory_group_id > 0
            )
            """
        )

    def _replace_database(self, documents, chunks: list[ChunkRecord]) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM chunks")
            conn.execute("DELETE FROM documents")

            document_ids: dict[str, int] = {}
            for document in documents:
                cursor = conn.execute(
                    """
                    INSERT INTO documents (
                        source_name,
                        source_path,
                        title,
                        file_type,
                        checksum,
                        char_count
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        document.source_name,
                        str(document.source_path),
                        document.title,
                        document.file_type,
                        document.checksum,
                        len(document.text),
                    ),
                )
                document_ids[document.source_name] = cursor.lastrowid

            for chunk in chunks:
                metadata = {
                    "source_id": chunk.source_id,
                    "source_name": chunk.source_name,
                    "source_path": chunk.source_path,
                    "title": chunk.title,
                    "chunk_index": chunk.chunk_index,
                }
                metadata.update(chunk.metadata)
                metadata.update(
                    classify_chunk_group(
                        source_name=chunk.source_name,
                        title=chunk.title,
                        text=chunk.text,
                        metadata=metadata,
                    )
                )
                conn.execute(
                    """
                    INSERT INTO chunks (
                        chunk_id,
                        document_id,
                        chunk_index,
                        content,
                        char_count,
                        metadata_json
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        chunk.chunk_id,
                        document_ids[chunk.source_name],
                        chunk.chunk_index,
                        chunk.text,
                        len(chunk.text),
                        json.dumps(metadata, ensure_ascii=False),
                    ),
                )

    def _build_indexes(self, chunks: list[ChunkRecord]) -> None:
        texts = [chunk.text for chunk in chunks]
        if not texts:
            raise ValueError("No chunks were generated from the source documents.")

        embeddings = embed_texts(texts, self.config)
        embeddings = np.asarray(embeddings, dtype="float32")
        faiss.normalize_L2(embeddings)
        index = faiss.IndexFlatIP(embeddings.shape[1])
        index.add(embeddings)
        faiss.write_index(index, str(self.config.faiss_path))

        vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
        matrix = vectorizer.fit_transform(texts)
        with self.config.tfidf_path.open("wb") as handle:
            pickle.dump(
                {
                    "vectorizer": vectorizer,
                    "matrix": matrix,
                    "chunk_ids": [chunk.chunk_id for chunk in chunks],
                },
                handle,
            )

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.config.sqlite_path)
