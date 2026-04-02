from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import faiss
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer

from .config import AppConfig, get_default_config
from .retrieval import (
    embed_texts,
    extract_priority_legal_terms,
    extract_query_terms,
    load_reranker_model,
)


@dataclass
class MemoryHit:
    entry_id: int
    session_id: str
    score: float
    relevance: float
    question: str
    answer: str
    created_at: str
    text: str
    metadata: dict[str, Any]


class SessionMemoryRetriever:
    def __init__(
        self,
        rows: list[dict[str, object]],
        config: AppConfig | None = None,
    ):
        self.config = config or get_default_config()
        self.rows = rows
        self.row_map = {int(row["id"]): row for row in rows}
        self.row_ids = [int(row["id"]) for row in rows]
        self.reranker = load_reranker_model(self.config)
        self.vectorizer: TfidfVectorizer | None = None
        self.tfidf_matrix = None
        self.index = None

        texts = [str(row["text"]) for row in rows]
        if not texts:
            return

        embeddings = np.asarray(embed_texts(texts, self.config), dtype="float32")
        faiss.normalize_L2(embeddings)
        self.index = faiss.IndexFlatIP(embeddings.shape[1])
        self.index.add(embeddings)

        self.vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
        self.tfidf_matrix = self.vectorizer.fit_transform(texts)

    def retrieve(self, query: str, min_relevance: float = 0.62) -> list[MemoryHit]:
        if not self.rows or self.index is None or self.vectorizer is None or self.tfidf_matrix is None:
            return []

        dense_rank = self._dense_candidates(query, min(8, len(self.row_ids)))
        sparse_rank = self._sparse_candidates(query, min(8, len(self.row_ids)))
        query_terms = extract_query_terms(query)
        preferred_terms = extract_priority_legal_terms(query)

        scored: list[MemoryHit] = []
        for row_id in set(dense_rank) | set(sparse_rank):
            row = self.row_map[row_id]
            text = str(row["text"])
            coverage = sum(1 for term in query_terms if term in text)
            preferred_hits = sum(1 for term in preferred_terms if term in text)
            score = dense_rank.get(row_id, 0.0) * 0.5
            score += sparse_rank.get(row_id, 0.0) * 0.25
            score += min(coverage, 8) * 0.025
            score += min(preferred_hits, 4) * 0.06
            if _looks_follow_up_query(query):
                score += _follow_up_bonus(query, row)
            scored.append(
                MemoryHit(
                    entry_id=row_id,
                    session_id=str(row.get("session_id", "")),
                    score=score,
                    relevance=0.0,
                    question=str(row.get("question", "")),
                    answer=str(row.get("answer", "")),
                    created_at=str(row.get("created_at", "")),
                    text=text,
                    metadata={
                        "memory_type": "chat_turn",
                        "memory_group_id": int(row.get("memory_group_id", 0) or 0),
                        "memory_keywords": row.get("memory_keywords", []),
                    },
                )
            )

        scored.sort(key=lambda item: item.score, reverse=True)
        self._rerank(query, scored)
        scored.sort(key=lambda item: item.score, reverse=True)
        if not scored:
            return []

        max_score = max(item.score for item in scored)
        min_score = min(item.score for item in scored)
        spread = max(max_score - min_score, 1e-6)
        if spread <= 1e-6:
            for item in scored:
                item.relevance = 1.0
        else:
            for item in scored:
                item.relevance = (item.score - min_score) / spread

        return [item for item in scored if item.relevance >= min_relevance]

    def _dense_candidates(self, query: str, top_n: int) -> dict[int, float]:
        embedding = np.asarray(embed_texts([query], self.config), dtype="float32")
        faiss.normalize_L2(embedding)
        scores, indices = self.index.search(embedding, top_n)
        ranked: dict[int, float] = {}
        max_score = max((float(value) for value in scores[0] if value > 0), default=0.0)
        for raw_score, idx in zip(scores[0], indices[0], strict=False):
            if idx < 0:
                continue
            normalized = float(raw_score) / max_score if max_score > 0 else 0.0
            ranked[self.row_ids[idx]] = max(0.0, normalized)
        return ranked

    def _sparse_candidates(self, query: str, top_n: int) -> dict[int, float]:
        query_vector = self.vectorizer.transform([query])
        scores = (self.tfidf_matrix @ query_vector.T).toarray().ravel()
        indices = np.argsort(scores)[::-1][:top_n]
        ranked: dict[int, float] = {}
        max_score = max((float(scores[idx]) for idx in indices if scores[idx] > 0), default=0.0)
        for idx in indices:
            if scores[idx] <= 0:
                continue
            normalized = float(scores[idx]) / max_score if max_score > 0 else 0.0
            ranked[self.row_ids[idx]] = max(0.0, normalized)
        return ranked

    def _rerank(self, query: str, candidates: list[MemoryHit]) -> None:
        if not candidates or self.reranker is None:
            return
        head = candidates[: min(6, len(candidates))]
        pairs = [(query, item.text[:1000]) for item in head]
        try:
            rerank_scores = self.reranker.predict(pairs)
            rerank_max = max((float(score) for score in rerank_scores), default=0.0)
            rerank_min = min((float(score) for score in rerank_scores), default=0.0)
            spread = max(rerank_max - rerank_min, 1e-6)
            for item, rerank_score in zip(head, rerank_scores, strict=False):
                normalized = (float(rerank_score) - rerank_min) / spread
                item.score += normalized * 0.35
        except Exception:
            return


def _looks_follow_up_query(query: str) -> bool:
    query = query.strip()
    if len(query) <= 20:
        return True
    return bool(
        re.search(
            r"(这个|这个情况|那这个|那这种|上一个|上面|前面|刚才|继续|展开|详细说说|为什么|那如果|这种情况下|该行为)",
            query,
        )
    )


def _follow_up_bonus(query: str, row: dict[str, object]) -> float:
    score = 0.0
    text = f"{row.get('question', '')}\n{row.get('answer', '')}"
    if re.search(r"(这个|那这个|上面|前面|继续|为什么)", query):
        if any(term in text for term in ("是否", "构成", "属于", "结论", "回答")):
            score += 0.08
    if re.search(r"(展开|详细|依据|法条)", query):
        if any(term in text for term in ("法条", "依据", "刑法", "民法典", "处罚")):
            score += 0.08
    return score
