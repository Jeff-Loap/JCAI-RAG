from __future__ import annotations

import argparse
import os

from huggingface_hub.errors import LocalEntryNotFoundError
from huggingface_hub import snapshot_download

from legal_agent.config import (
    EMBEDDING_REPO_CANDIDATES,
    RERANKER_REPO_CANDIDATES,
    get_default_config,
)


def download_repo(repo_id: str) -> str:
    try:
        return snapshot_download(repo_id=repo_id)
    except LocalEntryNotFoundError:
        original_endpoint = os.environ.get("HF_ENDPOINT", "").strip()
        if original_endpoint:
            raise
        os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"
        print(f"Direct download failed, retrying via mirror for {repo_id}")
        return snapshot_download(repo_id=repo_id)
    finally:
        os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="下载本地法律 RAG 所需模型。")
    parser.add_argument(
        "--embedding-repo",
        default=EMBEDDING_REPO_CANDIDATES[0],
        help="指定要下载的 embedding 模型 repo id。",
    )
    parser.add_argument(
        "--reranker-repo",
        default=RERANKER_REPO_CANDIDATES[0],
        help="指定要下载的 reranker 模型 repo id。",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    embedding_repo = args.embedding_repo
    reranker_repo = args.reranker_repo

    print(f"Downloading embedding model: {embedding_repo}")
    embedding_path = download_repo(embedding_repo)
    print(f"Embedding cached at: {embedding_path}")

    print(f"Downloading reranker model: {reranker_repo}")
    reranker_path = download_repo(reranker_repo)
    print(f"Reranker cached at: {reranker_path}")

    config = get_default_config()
    print(f"Resolved embedding model: {config.embedding_model_name} -> {config.embedding_model_dir}")
    print(f"Resolved reranker model: {config.reranker_model_name} -> {config.reranker_model_dir}")


if __name__ == "__main__":
    main()
