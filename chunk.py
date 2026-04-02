# -*- coding: utf-8 -*-
import json
import re

IN_JSONL = r"D:\PythonFile\JCAI\RAG\agent_guide_ocr_pages.jsonl"
OUT_CHUNKS = r"D:\PythonFile\JCAI\RAG\agent_guide_chunks.jsonl"

CHUNK_CHAR = 1200       # 简化：用字符数近似 token
OVERLAP_CHAR = 200

def split_with_overlap(text: str, chunk_size: int, overlap: int):
    text = text.strip()
    if not text:
        return []
    chunks = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        chunk = text[start:end]
        chunks.append(chunk)
        if end == len(text):
            break
        start = max(0, end - overlap)
    return chunks

def main():
    all_pages = []
    with open(IN_JSONL, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            obj = json.loads(line)
            all_pages.append(obj)
    chunks_out = []
    for p in all_pages:
        page_no = p["page"]
        text = re.sub(r"\n{3,}", "\n\n", p["text"]).strip()
        pieces = split_with_overlap(text, CHUNK_CHAR, OVERLAP_CHAR)
        for idx, piece in enumerate(pieces):
            chunks_out.append({
                "id": f"agent_guide_p{page_no}_c{idx}",
                "text": piece,
                "metadata": {
                    "pdf": "agent_guide.pdf",
                    "page_start": page_no,
                    "page_end": page_no,
                    "chunk_index": idx
                }
            })

    with open(OUT_CHUNKS, "w", encoding="utf-8") as f:
        for c in chunks_out:
            f.write(json.dumps(c, ensure_ascii=False) + "\n")

    print(f"切分完成：chunks={len(chunks_out)} 输出：{OUT_CHUNKS}")

if __name__ == "__main__":
    main()