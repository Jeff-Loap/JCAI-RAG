# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import tempfile
from dataclasses import replace
from pathlib import Path

import faiss
import numpy as np

from legal_agent import LegalRAGStore, get_default_config
from legal_agent.config import EMBEDDING_REPO_CANDIDATES
from legal_agent.retrieval import LocalHybridRetriever, embed_texts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="评测本地法律检索 embedding 模型效果。")
    parser.add_argument(
        "--benchmark",
        default=str(Path(__file__).resolve().parent / "eval" / "legal_retrieval_benchmark.json"),
        help="评测样本 JSON 文件路径。",
    )
    parser.add_argument(
        "--models",
        nargs="*",
        default=list(EMBEDDING_REPO_CANDIDATES),
        help="待评测的 embedding 模型 repo id 或本地目录路径。",
    )
    parser.add_argument(
        "--top-k",
        type=int,
        default=5,
        help="每个问题评估的 top-k 命中范围。",
    )
    parser.add_argument(
        "--details",
        action="store_true",
        help="输出每道题的详细命中情况。",
    )
    return parser.parse_args()


def load_benchmark(path: Path) -> list[dict]:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_model_dir(model_spec: str) -> tuple[str, Path] | None:
    path = Path(model_spec)
    if path.exists():
        return path.name, path.resolve()

    hub_dir = Path.home() / ".cache" / "huggingface" / "hub"
    snapshot_root = hub_dir / f"models--{model_spec.replace('/', '--')}" / "snapshots"
    snapshots = sorted(snapshot_root.glob("*"))
    if not snapshots:
        return None
    return model_spec, snapshots[-1].resolve()


def build_temp_retriever(store: LegalRAGStore, model_name: str, model_dir: Path) -> LocalHybridRetriever:
    base_config = store.config
    chunk_rows = store.fetch_chunks()
    texts = [row["text"] for row in chunk_rows]
    embeddings = np.asarray(embed_texts(texts, replace(base_config, embedding_model_name=model_name, embedding_model_dir=model_dir)), dtype="float32")
    embeddings = embeddings.astype("float32", copy=False)
    faiss.normalize_L2(embeddings)

    with tempfile.NamedTemporaryFile(suffix=".faiss", delete=False) as handle:
        faiss_path = Path(handle.name)

    index = faiss.IndexFlatIP(embeddings.shape[1])
    index.add(embeddings)
    faiss.write_index(index, str(faiss_path))

    retriever_config = replace(
        base_config,
        embedding_model_name=model_name,
        embedding_model_dir=model_dir,
        faiss_path=faiss_path,
    )
    return LocalHybridRetriever(chunk_rows, retriever_config)


def chunk_matches(chunk: dict, expected: dict) -> bool:
    if chunk["source_name"] != expected["source_name"]:
        return False
    article_anchor = str(chunk.get("article_anchor", "") or "")
    return article_anchor == expected["article_anchor"]


def evaluate_model(
    retriever: LocalHybridRetriever,
    benchmark: list[dict],
    top_k: int,
) -> tuple[dict[str, float], list[dict[str, object]]]:
    hits_at_1 = 0
    hits_at_k = 0
    reciprocal_rank = 0.0
    details: list[dict[str, object]] = []

    for item in benchmark:
        results = [
            {
                "source_name": chunk.source_name,
                "article_anchor": str(chunk.metadata.get("article_anchor", "") or ""),
                "score": round(chunk.score, 4),
            }
            for chunk in retriever.retrieve(item["question"], top_k=top_k)
        ]

        matched_rank = None
        for idx, chunk in enumerate(results, start=1):
            if any(chunk_matches(chunk, expected) for expected in item["expected"]):
                matched_rank = idx
                break

        if matched_rank == 1:
            hits_at_1 += 1
        if matched_rank is not None:
            hits_at_k += 1
            reciprocal_rank += 1.0 / matched_rank

        details.append(
            {
                "id": item["id"],
                "question": item["question"],
                "matched_rank": matched_rank,
                "results": results,
            }
        )

    total = max(len(benchmark), 1)
    metrics = {
        "hit@1": hits_at_1 / total,
        f"hit@{top_k}": hits_at_k / total,
        "mrr": reciprocal_rank / total,
    }
    return metrics, details


def format_metrics(metrics: dict[str, float]) -> str:
    return "  ".join(f"{name}={value:.3f}" for name, value in metrics.items())


def main() -> None:
    args = parse_args()
    benchmark = load_benchmark(Path(args.benchmark))
    store = LegalRAGStore(get_default_config())

    print(f"评测样本数：{len(benchmark)}")
    print(f"知识库 chunks：{store.get_stats().chunks}")
    print("")

    for model_spec in args.models:
        resolved = resolve_model_dir(model_spec)
        if resolved is None:
            print(f"[跳过] 未找到本地模型缓存：{model_spec}")
            continue

        model_name, model_dir = resolved
        print(f"[评测] {model_name}")
        retriever = build_temp_retriever(store, model_name, model_dir)
        metrics, details = evaluate_model(retriever, benchmark, args.top_k)
        print(format_metrics(metrics))

        if args.details:
            for item in details:
                print(f"- {item['id']} | matched_rank={item['matched_rank']}")
                for rank, result in enumerate(item["results"], start=1):
                    print(
                        f"  {rank}. {result['source_name']} {result['article_anchor']} score={result['score']}"
                    )
        Path(retriever.config.faiss_path).unlink(missing_ok=True)
        print("")


if __name__ == "__main__":
    main()
