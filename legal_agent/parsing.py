from __future__ import annotations

import csv
import hashlib
import json
import re
import sqlite3
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET

import pdfplumber


DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass
class SourceDocument:
    source_id: str
    source_path: Path
    source_name: str
    title: str
    file_type: str
    text: str
    checksum: str
    metadata: dict[str, Any]


@dataclass
class ChunkRecord:
    source_id: str
    chunk_id: str
    source_name: str
    source_path: str
    title: str
    chunk_index: int
    text: str
    metadata: dict[str, Any]


def discover_source_files(
    source_roots: Iterable[Path],
    excluded_dir_names: Iterable[str],
    supported_extensions: Iterable[str],
) -> list[Path]:
    excluded = set(excluded_dir_names)
    supported = {suffix.lower() for suffix in supported_extensions}
    candidates: list[tuple[int, Path]] = []

    for root_index, root in enumerate(source_roots):
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.lower() not in supported:
                continue
            if any(part in excluded for part in path.parts):
                continue
            candidates.append((root_index, path))

    grouped: dict[str, list[tuple[int, Path]]] = {}
    for root_index, path in candidates:
        grouped.setdefault(_dedupe_key(path), []).append((root_index, path))

    ext_priority = {
        ".pdf": 0,
        ".docx": 1,
        ".jsonl": 2,
        ".csv": 3,
        ".db": 4,
        ".sqlite": 4,
        ".sqlite3": 4,
    }
    preferred: list[Path] = []
    for items in grouped.values():
        _, chosen = sorted(
            items,
            key=lambda item: (
                ext_priority.get(item[1].suffix.lower(), 99),
                item[0],
                _has_duplicate_suffix(item[1].stem),
                item[1].name,
            ),
        )[0]
        preferred.append(chosen)
    return sorted(preferred, key=lambda item: item.name.lower())


def load_source_documents(path: Path) -> list[SourceDocument]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        document = build_pdf_source_document(path)
        return [document] if document.text else []
    if suffix == ".docx":
        text = extract_docx_text(path)
        document = _make_source_document(
            source_path=path,
            source_name=path.name,
            title=path.stem,
            file_type=suffix.lstrip("."),
            text=text,
        )
        return [document] if document.text else []
    if suffix == ".jsonl":
        return load_jsonl_documents(path)
    if suffix == ".csv":
        return load_csv_documents(path)
    if suffix in {".db", ".sqlite", ".sqlite3"}:
        return load_sqlite_documents(path)
    raise ValueError(f"Unsupported file type: {path}")


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx_text(path)
    if suffix == ".pdf":
        return extract_pdf_text(path)
    raise ValueError(f"Unsupported file type: {path}")


def load_jsonl_documents(path: Path) -> list[SourceDocument]:
    documents = []
    with path.open("r", encoding="utf-8") as handle:
        for index, line in enumerate(handle, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                payload = {"text": line}
            text = _flatten_to_text(payload)
            document = _make_source_document(
                source_path=path,
                source_name=f"{path.name}::line{index}",
                title=f"{path.stem} line {index}",
                file_type="jsonl",
                text=text,
            )
            if document.text:
                documents.append(document)
    return documents


def load_csv_documents(path: Path) -> list[SourceDocument]:
    documents = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for index, row in enumerate(reader, start=1):
            text = _flatten_to_text(row)
            document = _make_source_document(
                source_path=path,
                source_name=f"{path.name}::row{index}",
                title=f"{path.stem} row {index}",
                file_type="csv",
                text=text,
            )
            if document.text:
                documents.append(document)
    return documents


def load_sqlite_documents(path: Path) -> list[SourceDocument]:
    documents = []
    conn = sqlite3.connect(path)
    try:
        table_rows = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()
        for (table_name,) in table_rows:
            cursor = conn.execute(f'SELECT * FROM "{table_name}"')
            columns = [desc[0] for desc in cursor.description or []]
            for index, row in enumerate(cursor.fetchall(), start=1):
                payload = dict(zip(columns, row))
                text = _flatten_to_text(payload)
                document = _make_source_document(
                    source_path=path,
                    source_name=f"{path.name}::{table_name}#{index}",
                    title=f"{path.stem}/{table_name}/{index}",
                    file_type=path.suffix.lower().lstrip("."),
                    text=text,
                )
                if document.text:
                    documents.append(document)
    finally:
        conn.close()
    return documents


def extract_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    paragraphs = []
    for node in root.findall(".//w:body/w:p", DOCX_NS):
        text = "".join(part.text or "" for part in node.findall(".//w:t", DOCX_NS))
        text = _normalize_text(text)
        if text:
            paragraphs.append(text)
    return "\n\n".join(paragraphs).strip()


def extract_pdf_text(path: Path) -> str:
    return build_pdf_source_document(path).text


def build_pdf_source_document(path: Path) -> SourceDocument:
    pages = []
    page_spans: list[dict[str, int]] = []
    cursor = 0
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = _normalize_text(page.extract_text() or "")
            if text:
                pages.append(text)
                start = cursor
                end = start + len(text)
                page_spans.append(
                    {
                        "page": page_number,
                        "start": start,
                        "end": end,
                    }
                )
                cursor = end + 2

    return _make_source_document(
        source_path=path,
        source_name=path.name,
        title=path.stem,
        file_type="pdf",
        text="\n\n".join(pages).strip(),
        metadata={
            "page_spans": page_spans,
            "page_count": len(page_spans),
        },
    )


def split_into_chunks(document: SourceDocument, chunk_size: int, overlap: int) -> list[ChunkRecord]:
    legal_chunks = _split_legal_article_chunks(document)
    if legal_chunks:
        return legal_chunks

    paragraphs = _split_paragraphs_with_offsets(document.text)
    if not paragraphs:
        return []

    chunks: list[ChunkRecord] = []
    chunk_index = 0
    buffer_items: list[tuple[str, int, int]] = []
    buffer_len = 0
    buffer_preview_start: int | None = None

    for paragraph, paragraph_start, paragraph_end in paragraphs:
        if len(paragraph) > chunk_size:
            if buffer_items:
                chunks.append(
                    _make_chunk_from_items(
                        document,
                        chunk_index,
                        buffer_items,
                        buffer_preview_start,
                    )
                )
                chunk_index += 1
                buffer_items = []
                buffer_len = 0
                buffer_preview_start = None

            for piece, piece_start, piece_end in _split_long_paragraph(
                paragraph,
                paragraph_start,
                chunk_size,
                overlap,
            ):
                chunks.append(
                    ChunkRecord(
                        source_id=document.source_id,
                        chunk_id=f"{document.source_id}_c{chunk_index}",
                        source_name=document.source_name,
                        source_path=str(document.source_path),
                        title=document.title,
                        chunk_index=chunk_index,
                        text=piece,
                        metadata=_build_chunk_metadata(
                            document,
                            piece_start,
                            piece_end,
                            preview_start_offset=piece_start,
                        ),
                    )
                )
                chunk_index += 1
            continue

        paragraph_len = len(paragraph) if not buffer_items else len(paragraph) + 2
        if buffer_items and buffer_len + paragraph_len > chunk_size:
            chunks.append(
                _make_chunk_from_items(
                    document,
                    chunk_index,
                    buffer_items,
                    buffer_preview_start,
                )
            )
            chunk_index += 1
            buffer_items = _select_overlap_items(buffer_items, overlap)
            buffer_len = _joined_items_length(buffer_items)
            buffer_preview_start = paragraph_start

        buffer_items.append((paragraph, paragraph_start, paragraph_end))
        buffer_len = _joined_items_length(buffer_items)
        if buffer_preview_start is None:
            buffer_preview_start = paragraph_start

    if buffer_items:
        chunks.append(
            _make_chunk_from_items(
                document,
                chunk_index,
                buffer_items,
                buffer_preview_start,
            )
        )
    return chunks


def _split_long_paragraph(
    text: str,
    start_offset: int,
    chunk_size: int,
    overlap: int,
) -> Iterable[tuple[str, int, int]]:
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        piece = text[start:end].strip()
        piece_start = start_offset + start
        piece_end = piece_start + len(piece)
        yield piece, piece_start, piece_end
        if end >= len(text):
            break
        start = max(0, end - overlap)


def _make_chunk_from_items(
    document: SourceDocument,
    chunk_index: int,
    items: list[tuple[str, int, int]],
    preview_start_offset: int | None,
    extra_metadata: dict[str, Any] | None = None,
) -> ChunkRecord:
    text = "\n\n".join(part for part, _, _ in items).strip()
    start_offset = items[0][1]
    end_offset = items[-1][2]
    return ChunkRecord(
        source_id=document.source_id,
        chunk_id=f"{document.source_id}_c{chunk_index}",
        source_name=document.source_name,
        source_path=str(document.source_path),
        title=document.title,
        chunk_index=chunk_index,
        text=text,
        metadata=_build_chunk_metadata(
            document,
            start_offset,
            end_offset,
            preview_start_offset=preview_start_offset,
            extra_metadata=extra_metadata,
        ),
    )


def _select_overlap_items(
    items: list[tuple[str, int, int]],
    overlap: int,
) -> list[tuple[str, int, int]]:
    if overlap <= 0 or not items:
        return []

    selected: list[tuple[str, int, int]] = []
    total = 0
    for item in reversed(items):
        text = item[0]
        addition = len(text) if not selected else len(text) + 2
        selected.insert(0, item)
        total += addition
        if total >= overlap:
            break
    return selected


def _joined_items_length(items: list[tuple[str, int, int]]) -> int:
    if not items:
        return 0
    return sum(len(text) for text, _, _ in items) + 2 * (len(items) - 1)


def _normalize_text(text: str) -> str:
    text = text.replace("\u3000", "  ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _make_source_document(
    source_path: Path,
    source_name: str,
    title: str,
    file_type: str,
    text: str,
    metadata: dict[str, Any] | None = None,
) -> SourceDocument:
    normalized_text = _normalize_text(text)
    checksum = hashlib.sha256(normalized_text.encode("utf-8")).hexdigest()
    source_key = f"{source_path.resolve()}::{source_name}::{title}::{file_type}"
    source_id = hashlib.sha1(source_key.encode("utf-8")).hexdigest()[:16]
    return SourceDocument(
        source_id=source_id,
        source_path=source_path.resolve(),
        source_name=source_name,
        title=title,
        file_type=file_type,
        text=normalized_text,
        checksum=checksum,
        metadata=metadata or {},
    )


def _split_paragraphs_with_offsets(text: str) -> list[tuple[str, int, int]]:
    paragraphs: list[tuple[str, int, int]] = []
    search_start = 0
    for part in re.split(r"\n{2,}", text):
        paragraph = part.strip()
        if not paragraph:
            continue
        offset = text.find(paragraph, search_start)
        if offset < 0:
            offset = search_start
        end = offset + len(paragraph)
        paragraphs.append((paragraph, offset, end))
        search_start = end
    return paragraphs


def _build_chunk_metadata(
    document: SourceDocument,
    start_offset: int,
    end_offset: int,
    preview_start_offset: int | None = None,
    extra_metadata: dict[str, Any] | None = None,
) -> dict[str, Any]:
    metadata = dict(document.metadata)
    if extra_metadata:
        metadata.update(extra_metadata)
    metadata["char_start"] = max(0, start_offset)
    metadata["char_end"] = max(metadata["char_start"], end_offset)
    preview_start = metadata["char_start"] if preview_start_offset is None else max(
        metadata["char_start"],
        preview_start_offset,
    )
    metadata["preview_char_start"] = preview_start
    metadata["preview_offset_in_chunk"] = max(0, preview_start - metadata["char_start"])
    page_spans = document.metadata.get("page_spans", [])
    if not page_spans:
        return metadata

    overlapping_pages = [
        span["page"]
        for span in page_spans
        if not (metadata["char_end"] < span["start"] or metadata["char_start"] > span["end"])
    ]
    if not overlapping_pages:
        nearest_page = page_spans[-1]["page"] if page_spans else 1
        overlapping_pages = [nearest_page]

    metadata["page_numbers"] = overlapping_pages
    metadata["page_start"] = overlapping_pages[0]
    metadata["page_end"] = overlapping_pages[-1]
    return metadata


def _split_legal_article_chunks(document: SourceDocument) -> list[ChunkRecord]:
    article_pattern = re.compile(
        r"(?<![\u4e00-\u9fffA-Za-z0-9])"
        r"(第[一二三四五六七八九十百千万零〇\d]+条(?:之[一二三四五六七八九十百千万零〇\d]+)?)"
        r"(?=[\s　])"
    )
    matches = list(article_pattern.finditer(document.text))
    if len(matches) < 3:
        return []

    chunks: list[ChunkRecord] = []
    boundaries = [match.start() for match in matches] + [len(document.text)]
    for chunk_index, match in enumerate(matches):
        start_offset = boundaries[chunk_index]
        end_offset = boundaries[chunk_index + 1]
        text = document.text[start_offset:end_offset].strip()
        if not text:
            continue
        article_anchor = match.group(1)
        article_heading = re.split(r"[\n。；;]", text, maxsplit=1)[0].strip()
        chunks.append(
            ChunkRecord(
                source_id=document.source_id,
                chunk_id=f"{document.source_id}_c{chunk_index}",
                source_name=document.source_name,
                source_path=str(document.source_path),
                title=document.title,
                chunk_index=chunk_index,
                text=text,
                metadata=_build_chunk_metadata(
                    document,
                    start_offset,
                    end_offset,
                    preview_start_offset=start_offset,
                    extra_metadata={
                        "law_chunk_type": "article",
                        "article_anchor": article_anchor,
                        "article_heading": article_heading[:120],
                    },
                ),
            )
        )
    return chunks


def _dedupe_key(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".pdf", ".docx"}:
        normalized = re.sub(r" \(\d+\)$", "", path.stem)
        return normalized
    return str(path.resolve())


def _has_duplicate_suffix(stem: str) -> int:
    return 1 if re.search(r" \(\d+\)$", stem) else 0


def _flatten_to_text(payload: Any, prefix: str = "") -> str:
    lines: list[str] = []
    if isinstance(payload, dict):
        for key, value in payload.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            child_text = _flatten_to_text(value, child_prefix)
            if child_text:
                lines.append(child_text)
    elif isinstance(payload, list):
        for index, value in enumerate(payload):
            child_prefix = f"{prefix}[{index}]" if prefix else f"[{index}]"
            child_text = _flatten_to_text(value, child_prefix)
            if child_text:
                lines.append(child_text)
    elif payload is None:
        return ""
    else:
        value = str(payload).strip()
        if value:
            return f"{prefix}: {value}" if prefix else value
    return "\n".join(lines).strip()
