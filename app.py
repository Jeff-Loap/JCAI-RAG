# -*- coding: utf-8 -*-
from __future__ import annotations

import configparser
from pathlib import Path
from uuid import uuid4

import fitz
import streamlit as st

from legal_agent import LegalRAGAgent, LegalRAGStore, get_default_config
from legal_agent.config import LLMSettings


st.set_page_config(page_title="法律 RAG Agent", page_icon="⚖️", layout="wide")
APP_DIR = Path(__file__).resolve().parent
CONFIG_INI_PATH = APP_DIR / "config.ini"


@st.cache_resource(show_spinner=False)
def get_store() -> LegalRAGStore:
    return LegalRAGStore(get_default_config())


@st.cache_resource(show_spinner=False)
def get_agent() -> LegalRAGAgent:
    config = get_default_config()
    return LegalRAGAgent(store=get_store(), config=config)


def rebuild_knowledge_base() -> str:
    store = get_store()
    stats = store.rebuild()
    get_agent().refresh()
    return f"重建完成：{stats.documents} 个文档，{stats.chunks} 个 chunks。"


def load_llm_settings_from_ini() -> LLMSettings:
    parser = configparser.ConfigParser()
    if CONFIG_INI_PATH.exists():
        parser.read(CONFIG_INI_PATH, encoding="utf-8")

    return LLMSettings(
        base_url=parser.get("llm", "base_url", fallback="").strip(),
        api_key=parser.get("llm", "api_key", fallback="").strip(),
        model=parser.get("llm", "model", fallback="").strip(),
        temperature=parser.getfloat("llm", "temperature", fallback=0.1),
        max_tokens=parser.getint("llm", "max_tokens", fallback=700),
        retrieval_mode=parser.get("llm", "retrieval_mode", fallback="llm_retrieval").strip() or "llm_retrieval",
        answer_profile=parser.get("llm", "answer_profile", fallback="quality").strip() or "quality",
    )


@st.cache_data(show_spinner=False)
def render_pdf_pages_as_images(
    source_path: str,
    page_start: int,
    page_end: int,
) -> list[tuple[int, bytes]]:
    images: list[tuple[int, bytes]] = []
    document = fitz.open(source_path)
    try:
        total_pages = len(document)
        start = max(1, min(page_start, total_pages))
        end = max(start, min(page_end, total_pages))
        for page_number in range(start, end + 1):
            page = document.load_page(page_number - 1)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            images.append((page_number, pixmap.tobytes("png")))
    finally:
        document.close()
    return images


def render_sidebar() -> LLMSettings:
    config = get_default_config()
    store = get_store()
    stats = store.get_stats()
    ini_settings = load_llm_settings_from_ini()

    with st.sidebar:
        st.title("法律知识库")
        st.caption("支持大模型检索与混合检索两种本地法律问答模式。")
        st.write("数据目录：")
        for path in config.source_roots:
            st.write(f"- `{path}`")
        st.write(f"SQLite：`{config.sqlite_path}`")
        st.write(f"向量索引：`{config.faiss_path}`")
        st.write(f"Embedding：`{config.embedding_model_name}`")
        st.write(
            "Reranker："
            f"`{config.reranker_model_name if config.reranker_model_dir else f'{config.reranker_model_name}（未下载，当前使用规则重排）'}`"
        )
        st.write(f"当前文档数：`{stats.documents}`")
        st.write(f"当前 chunks：`{stats.chunks}`")

        if st.button("重建索引", width="stretch"):
            with st.spinner("正在重建本地知识库..."):
                message = rebuild_knowledge_base()
            st.success(message)

        with st.expander("已入库文件", expanded=False):
            for source in stats.sources:
                st.write(f"- {source}")

        st.subheader("历史记录")
        history_keyword = st.text_input(
            "搜索历史",
            value=st.session_state.get("history_keyword", ""),
            placeholder="按问题或答案关键词筛选",
        )
        st.session_state.history_keyword = history_keyword
        history_entries = store.list_history_entries(limit=50, keyword=history_keyword)

        if st.button("新建会话", width="stretch"):
            st.session_state.messages = []
            st.session_state.history_entry_id = None
            st.session_state.loaded_history_entry_id = None
            st.session_state.chat_session_id = uuid4().hex
            st.rerun()

        clear_label = "清空筛选结果" if history_keyword else "清空全部历史"
        confirm_clear = st.checkbox("确认执行清空", key="confirm_clear_history")
        if st.button(clear_label, width="stretch"):
            if not confirm_clear:
                st.warning("先勾选“确认执行清空”。")
            else:
                deleted = store.clear_history_entries(keyword=history_keyword)
                st.session_state.messages = []
                st.session_state.history_entry_id = None
                st.session_state.loaded_history_entry_id = None
                st.session_state.chat_session_id = uuid4().hex
                st.session_state.confirm_clear_history = False
                st.success(f"已删除 {deleted} 条历史记录。")
                st.rerun()

        if not history_entries:
            st.caption("还没有历史问答。")
        else:
            for entry in history_entries:
                label = f"{entry.created_at} | {truncate_text(entry.question, 22)}"
                left, right = st.columns([4, 1])
                with left:
                    if st.button(label, key=f"history_{entry.id}", width="stretch"):
                        st.session_state.history_entry_id = entry.id
                        st.rerun()
                with right:
                    if st.button("删", key=f"delete_history_{entry.id}", width="stretch"):
                        deleted = store.delete_history_entry(entry.id)
                        if deleted:
                            if st.session_state.get("loaded_history_entry_id") == entry.id:
                                st.session_state.messages = []
                                st.session_state.loaded_history_entry_id = None
                                st.session_state.chat_session_id = uuid4().hex
                            st.success(f"已删除历史 #{entry.id}")
                        else:
                            st.warning("该历史记录不存在或已删除。")
                        st.rerun()

        st.subheader("LLM 配置（可选）")
        st.caption(f"默认读取：`{CONFIG_INI_PATH}`")
        mode_options = {
            "大模型检索": "llm_retrieval",
            "混合检索（大模型+小模型）": "hybrid",
        }
        current_mode_label = next(
            (
                label
                for label, value in mode_options.items()
                if value == ini_settings.retrieval_mode
            ),
            "大模型检索",
        )
        retrieval_mode_label = st.selectbox(
            "问答模式",
            options=list(mode_options.keys()),
            index=list(mode_options.keys()).index(current_mode_label),
            help="“大模型检索”会先从本地文件做宽召回候选，再由大模型筛证据并回答；“混合检索”使用当前小模型检索 + 大模型整理。",
        )
        profile_options = {
            "快速模式": "fast",
            "精确模式": "quality",
        }
        current_profile_label = next(
            (
                label
                for label, value in profile_options.items()
                if value == ini_settings.answer_profile
            ),
            "精确模式",
        )
        answer_profile_label = st.selectbox(
            "回答档位",
            options=list(profile_options.keys()),
            index=list(profile_options.keys()).index(current_profile_label),
            help="快速模式尽量减少前置 LLM 步骤，优先响应速度；精确模式保留完整的法律改写、筛选、自检与引用校验链路。",
        )
        base_url = st.text_input(
            "Base URL",
            value=ini_settings.base_url,
            placeholder="https://your-openai-compatible-endpoint/v1",
        )
        api_key = st.text_input("API Key", value=ini_settings.api_key, type="password")
        model = st.text_input(
            "Model",
            value=ini_settings.model,
            placeholder="qwen / gpt / local-chat-model",
        )

    return LLMSettings(
        base_url=base_url,
        api_key=api_key,
        model=model,
        temperature=ini_settings.temperature,
        max_tokens=ini_settings.max_tokens,
        retrieval_mode=mode_options[retrieval_mode_label],
        answer_profile=profile_options[answer_profile_label],
    )


def main():
    llm_settings = render_sidebar()
    st.title("法律 RAG Agent")
    st.write("支持本地法律文件入库、混合检索和带引用的问答。")

    if "messages" not in st.session_state:
        st.session_state.messages = []
    if "history_entry_id" not in st.session_state:
        st.session_state.history_entry_id = None
    if "loaded_history_entry_id" not in st.session_state:
        st.session_state.loaded_history_entry_id = None
    if "chat_session_id" not in st.session_state:
        st.session_state.chat_session_id = uuid4().hex

    if st.session_state.history_entry_id is not None:
        entry = get_store().get_history_entry(st.session_state.history_entry_id)
        if entry is not None:
            session_entries = get_store().list_session_entries(entry.session_id)
            messages = []
            for item in session_entries:
                messages.append({"role": "user", "content": item.question})
                messages.append(
                    {
                        "role": "assistant",
                        "content": item.answer,
                        "turn_id": item.turn_id,
                        "thinking": item.thinking,
                        "question_segments": item.question_segments,
                        "answer_segments": item.answer_segments,
                        "citations": item.citations,
                        "llm_error": item.llm_error,
                        "conversation_scope": "legal",
                    }
                )
            st.session_state.messages = messages
            st.session_state.loaded_history_entry_id = entry.id
            st.session_state.chat_session_id = entry.session_id or uuid4().hex
        st.session_state.history_entry_id = None

    for message_index, message in enumerate(st.session_state.messages):
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
            if message.get("turn_id"):
                render_turn_segments(
                    turn_id=int(message["turn_id"]),
                    question_segments=message.get("question_segments", []),
                    answer_segments=message.get("answer_segments", []),
                )
            if message.get("thinking"):
                render_thinking_trace(message["thinking"], scope_key=f"thinking_{message_index}")
            if message.get("llm_error"):
                st.warning(f"LLM 已自动降级：{message['llm_error']}")
            render_retrieval_mode_notice(message)
            if message.get("memory_hits"):
                render_memory_hits(message["memory_hits"])
            if message.get("citations"):
                render_citations(message["citations"], scope_key=f"message_{message_index}")

    question = st.chat_input("请输入你的法律问题")
    if not question:
        return

    st.session_state.messages.append({"role": "user", "content": question})
    with st.chat_message("user"):
        st.markdown(question)

    final_result = None
    with st.chat_message("assistant"):
        thinking_placeholder = st.empty()
        response_placeholder = st.empty()
        thinking_buffer = ""
        answer_buffer = ""
        spinner_text = (
            "正在让大模型从本地候选证据中筛选并生成答案..."
            if llm_settings.retrieval_mode == "llm_retrieval"
            else "正在本地检索并组织答案..."
        )
        if llm_settings.answer_profile == "fast":
            spinner_text = "正在快速生成答案..."
        with st.spinner(spinner_text):
            for event in get_agent().stream_ask(
                question,
                session_id=st.session_state.chat_session_id,
                llm_settings=llm_settings,
            ):
                if event["type"] == "thinking_token":
                    thinking_buffer += event["content"]
                    with thinking_placeholder.container():
                        render_thinking_trace(thinking_buffer)
                elif event["type"] == "token":
                    answer_buffer += event["content"]
                    response_placeholder.markdown(answer_buffer)
                elif event["type"] == "done":
                    final_result = event["result"]

        if final_result is None:
            final_result = {
                "answer": answer_buffer or "生成过程中未返回结果。",
                "citations": [],
                "llm_used": False,
                "llm_error": "未收到完整结果。",
        }

        thinking_placeholder.empty()
        if final_result.get("thinking"):
            render_thinking_trace(final_result["thinking"], scope_key="current_thinking")
        response_placeholder.markdown(final_result["answer"])
        if final_result.get("llm_error"):
            st.warning(f"LLM 已自动降级：{final_result['llm_error']}")
        render_retrieval_mode_notice(final_result)
        if final_result.get("memory_hits"):
            render_memory_hits(final_result["memory_hits"])
        if final_result["citations"]:
            render_citations(final_result["citations"], scope_key="current_answer")

    saved_id = get_store().save_history_entry(
        session_id=st.session_state.chat_session_id,
        question=question,
        answer=final_result["answer"],
        thinking=final_result.get("thinking", ""),
        citations=final_result["citations"],
        llm_used=final_result.get("llm_used", False),
        llm_error=final_result.get("llm_error", ""),
    )
    saved_entry = get_store().get_history_entry(saved_id)
    if saved_entry is None:
        raise RuntimeError(f"Failed to load saved history entry: id={saved_id}")

    st.session_state.messages.append(
        {
            "role": "assistant",
            "content": final_result["answer"],
            "turn_id": saved_entry.turn_id,
            "thinking": final_result.get("thinking", ""),
            "question_segments": saved_entry.question_segments,
            "answer_segments": saved_entry.answer_segments,
            "citations": final_result["citations"],
            "llm_used": final_result.get("llm_used", False),
            "llm_error": final_result.get("llm_error", ""),
            "memory_hits": final_result.get("memory_hits", []),
            "retrieved_chunks": final_result.get("retrieved_chunks", []),
            "retrieval_mode": final_result.get("retrieval_mode", llm_settings.retrieval_mode),
            "conversation_scope": final_result.get("conversation_scope", "legal"),
            "scope_reason": final_result.get("scope_reason", ""),
        }
    )


def truncate_text(text: str, limit: int) -> str:
    clean = " ".join(text.split())
    if len(clean) <= limit:
        return clean
    return clean[: max(0, limit - 1)] + "…"


def render_retrieval_mode_notice(message: dict) -> None:
    if message.get("conversation_scope") == "general":
        st.info("当前问题已识别为通用对话，已跳过法律检索。")
        return

    if message.get("retrieval_mode") != "llm_retrieval":
        return

    llm_used = bool(message.get("llm_used"))
    citations = message.get("citations") or []
    retrieved_chunks = message.get("retrieved_chunks") or []

    if llm_used and citations:
        st.info("当前回答为“大模型检索”模式，由大模型从本地候选证据中筛选并引用。")
        return

    if retrieved_chunks:
        st.info("当前回答为“大模型检索”模式，但本轮未形成最终引用。")
        return

    st.warning("当前回答为“大模型检索”模式，但本轮未从本地知识库召回可用候选证据。")


def render_thinking_trace(thinking: str, scope_key: str = "thinking_current") -> None:
    clean = " ".join(str(thinking).split())
    if not clean:
        return
    with st.expander("思考过程", expanded=False):
        st.markdown(clean)


def render_turn_segments(
    turn_id: int,
    question_segments: list[dict],
    answer_segments: list[dict],
) -> None:
    with st.expander(f"时序切分 · Turn {turn_id}", expanded=False):
        st.markdown("**Q 段落序号**")
        for segment in question_segments:
            st.caption(f"Q{turn_id}.{segment['seq_id']}  {segment['text']}")
        st.markdown("**A 段落序号**")
        for segment in answer_segments:
            st.caption(f"A{turn_id}.{segment['seq_id']}  {segment['text']}")


def normalize_citation(citation: dict | str, scope_key: str, index: int) -> dict:
    if isinstance(citation, dict):
        normalized = dict(citation)
    else:
        normalized = {
            "label": str(citation),
            "source_name": str(citation),
            "source_path": "",
            "title": str(citation),
            "chunk_id": f"{scope_key}_{index}",
            "page_start": None,
            "page_end": None,
            "page_numbers": [],
            "file_type": "",
            "snippet": "",
        }
    normalized["preview_key"] = f"{scope_key}_{index}_{normalized.get('chunk_id', index)}"
    return normalized


def render_citations(citations: list[dict | str], scope_key: str) -> None:
    with st.expander("查看引用", expanded=False):
        for index, raw_citation in enumerate(citations, start=1):
            citation = normalize_citation(raw_citation, scope_key, index)
            label = citation.get("label", f"引用 {index}")
            page_start = citation.get("page_start")
            path_text = citation.get("source_path") or "未记录文件路径"
            button_label = f"预览引用 {index}"
            if page_start:
                button_label += f" · 第 {page_start} 页"

            st.markdown(f"**{label}**")
            left, right = st.columns([1, 4])
            with left:
                if citation.get("file_type") == "pdf" and citation.get("source_path"):
                    if st.button(button_label, key=f"preview_btn_{citation['preview_key']}", width="stretch"):
                        render_pdf_preview(citation)
                else:
                    st.caption("暂不支持预览")
            with right:
                st.caption(path_text)
                if citation.get("snippet"):
                    st.write(citation["snippet"])


def render_memory_hits(memory_hits: list[dict]) -> None:
    with st.expander("查看上下文记忆", expanded=False):
        for index, hit in enumerate(memory_hits, start=1):
            score = float(hit.get("score", 0.0) or 0.0)
            relevance = float(hit.get("relevance", 0.0) or 0.0)
            st.markdown(f"**记忆 {index}** · relevance={relevance:.3f} · score={score:.3f}")
            st.caption(hit.get("created_at", ""))
            st.write(f"问题：{hit.get('question', '')}")
            answer = " ".join(str(hit.get("answer", "")).split())
            if answer:
                st.write(f"回答：{answer[:280]}")


def render_pdf_preview(citation: dict) -> None:
    source_path = Path(str(citation.get("source_path", "")))
    if not source_path.exists():
        st.error(f"找不到 PDF 文件：{source_path}")
        return
    if source_path.suffix.lower() != ".pdf":
        st.info("当前引用不是 PDF 文件，暂不支持页内预览。")
        return

    page_start = int(citation.get("page_start") or 1)
    page_end = int(citation.get("page_end") or page_start)
    try:
        page_images = render_pdf_pages_as_images(
            str(source_path.resolve()),
            page_start=page_start,
            page_end=page_end,
        )
    except Exception as exc:
        st.error(f"加载 PDF 预览失败：{exc}")
        return

    title = citation.get("title", source_path.stem)
    if page_start == page_end:
        st.info(f"《{title}》命中第 {page_start} 页。")
    else:
        st.info(f"《{title}》命中第 {page_start}-{page_end} 页。")
    st.caption(str(source_path.resolve()))

    for page_number, image_bytes in page_images:
        st.image(
            image_bytes,
            caption=f"{source_path.name} · 第 {page_number} 页",
            width="stretch",
        )

    with source_path.open("rb") as handle:
        st.download_button(
            label="下载原始 PDF",
            data=handle.read(),
            file_name=source_path.name,
            mime="application/pdf",
            key=f"download_{citation['preview_key']}",
            width="stretch",
        )


if __name__ == "__main__":
    main()
