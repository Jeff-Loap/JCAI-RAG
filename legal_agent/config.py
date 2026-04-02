from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


EMBEDDING_REPO_CANDIDATES = (
    "BAAI/bge-base-zh-v1.5",
    "sentence-transformers/all-MiniLM-L6-v2",
    "BAAI/bge-small-zh-v1.5",
    "moka-ai/m3e-base",
)

RERANKER_REPO_CANDIDATES = (
    "BAAI/bge-reranker-base",
    "BAAI/bge-reranker-v2-m3",
    "maidalun1020/bce-reranker-base_v1",
)


def _resolve_cached_model_dir() -> tuple[str, Path]:
    env_value = os.getenv("RAG_EMBED_MODEL_DIR")
    if env_value:
        path = Path(env_value).expanduser().resolve()
        return _guess_repo_name_from_path(path), path

    resolved = _resolve_hf_snapshot_dir(EMBEDDING_REPO_CANDIDATES)
    if resolved is not None:
        return resolved

    hub_dir = Path.home() / ".cache" / "huggingface" / "hub"
    return EMBEDDING_REPO_CANDIDATES[0], hub_dir / "missing-embedding-model"


def _resolve_cached_reranker_dir() -> tuple[str, Path | None]:
    env_value = os.getenv("RAG_RERANK_MODEL_DIR")
    if env_value:
        path = Path(env_value).expanduser().resolve()
        return _guess_repo_name_from_path(path), path if path.exists() else None

    resolved = _resolve_hf_snapshot_dir(RERANKER_REPO_CANDIDATES)
    if resolved is not None:
        return resolved
    return RERANKER_REPO_CANDIDATES[0], None


def _resolve_hf_snapshot_dir(repo_ids: tuple[str, ...]) -> tuple[str, Path] | None:
    hub_dir = Path.home() / ".cache" / "huggingface" / "hub"
    for repo_id in repo_ids:
        snapshot_root = hub_dir / f"models--{repo_id.replace('/', '--')}" / "snapshots"
        snapshots = sorted(snapshot_root.glob("*"))
        if snapshots:
            return repo_id, snapshots[-1].resolve()
    return None


def _guess_repo_name_from_path(path: Path) -> str:
    parts = list(path.parts)
    if "hub" in parts:
        hub_index = parts.index("hub")
        if hub_index + 1 < len(parts):
            model_dir = parts[hub_index + 1]
            if model_dir.startswith("models--"):
                return model_dir.removeprefix("models--").replace("--", "/")
    return path.name


@dataclass(frozen=True)
class AppConfig:
    project_root: Path
    source_dir: Path
    source_roots: tuple[Path, ...]
    runtime_dir: Path
    sqlite_path: Path
    faiss_path: Path
    tfidf_path: Path
    manifest_path: Path
    embedding_model_name: str
    embedding_model_dir: Path
    reranker_model_name: str
    reranker_model_dir: Path | None
    excluded_dir_names: tuple[str, ...] = (
        "legal_agent_runtime",
        "legal_agent",
        "__pycache__",
        ".streamlit",
    )
    supported_extensions: tuple[str, ...] = (
        ".pdf",
        ".docx",
        ".jsonl",
        ".csv",
        ".db",
        ".sqlite",
        ".sqlite3",
    )
    chunk_size: int = 900
    chunk_overlap: int = 150
    dense_candidate_k: int = 24
    sparse_candidate_k: int = 24
    rerank_candidate_k: int = 12
    final_top_k: int = 6
    memory_relevance_threshold: float = 0.62


@dataclass(frozen=True)
class LLMSettings:
    base_url: str = ""
    api_key: str = ""
    model: str = ""
    temperature: float = 0.1
    max_tokens: int = 700
    retrieval_mode: str = "llm_retrieval"
    answer_profile: str = "quality"

    @property
    def enabled(self) -> bool:
        return bool(self.base_url and self.api_key and self.model)


def get_default_config() -> AppConfig:
    rag_dir = Path(__file__).resolve().parent.parent
    runtime_dir = rag_dir / "legal_agent_runtime"
    embedding_model_name, embedding_model_dir = _resolve_cached_model_dir()
    reranker_model_name, reranker_model_dir = _resolve_cached_reranker_dir()
    candidate_roots = [
        rag_dir / "pdf_data",
        rag_dir / "raw_data",
        rag_dir / "external_data",
        rag_dir / "external_db",
        rag_dir / "db_data",
    ]
    source_roots = tuple(path.resolve() for path in candidate_roots if path.exists())
    if not source_roots:
        source_roots = (rag_dir.resolve(),)
    return AppConfig(
        project_root=rag_dir.parent,
        source_dir=rag_dir,
        source_roots=source_roots,
        runtime_dir=runtime_dir,
        sqlite_path=runtime_dir / "rag_external.db",
        faiss_path=runtime_dir / "legal_chunks.faiss",
        tfidf_path=runtime_dir / "legal_chunks_tfidf.pkl",
        manifest_path=runtime_dir / "manifest.json",
        embedding_model_name=embedding_model_name,
        embedding_model_dir=embedding_model_dir,
        reranker_model_name=reranker_model_name,
        reranker_model_dir=reranker_model_dir,
    )


def load_llm_settings_from_env() -> LLMSettings:
    return LLMSettings(
        base_url=os.getenv("RAG_LLM_BASE_URL", "").strip(),
        api_key=os.getenv("RAG_LLM_API_KEY", "").strip(),
        model=os.getenv("RAG_LLM_MODEL", "").strip(),
        temperature=float(os.getenv("RAG_LLM_TEMPERATURE", "0.1")),
        max_tokens=int(os.getenv("RAG_LLM_MAX_TOKENS", "700")),
        retrieval_mode=os.getenv("RAG_RETRIEVAL_MODE", "llm_retrieval").strip() or "llm_retrieval",
        answer_profile=os.getenv("RAG_ANSWER_PROFILE", "quality").strip() or "quality",
    )
