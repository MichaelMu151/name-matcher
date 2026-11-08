"""Small but complete local Streamlit review workflow."""

from __future__ import annotations

import json
import tempfile
import uuid
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook


def _read_debug(path: str | Path) -> pd.DataFrame:
    return pd.read_json(path, lines=True)


def _candidate_options(value: Any) -> list[str]:
    if isinstance(value, list):
        result: list[str] = []
        for item in value:
            if isinstance(item, dict) and item.get("entity_id"):
                result.append(str(item["entity_id"]))
            elif item:
                result.append(str(item))
        return result
    if isinstance(value, str):
        try:
            return _candidate_options(json.loads(value))
        except (TypeError, json.JSONDecodeError):
            return []
    return []


def _resolve_path(text: str, filename: str | None = None) -> Path:
    """Resolve UI paths without assuming Streamlit's current working directory."""
    raw = Path(text).expanduser()
    if raw.is_absolute():
        return raw
    repo_root = Path(__file__).resolve().parents[2]
    candidates = [Path.cwd() / raw, repo_root / raw, repo_root.parent / raw]
    if filename:
        candidates.extend([repo_root / filename, repo_root.parent / filename])
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _apply_human_choices(source_path: Path, system_path: Path, log_path: Path, codebook: Any, choices: dict[str, tuple[str, str]], target_fields: tuple[str, ...], field_mapping: dict[str, str], session_id: str, operator: str, debug_records: pd.DataFrame | None = None) -> tuple[Path, Path]:
    """Apply only entity-id selections and return revised content/log files."""

    try:
        from .pipeline import Pipeline, _clean
    except ImportError:  # Streamlit executes this file as a script.
        from src.engine.pipeline import Pipeline, _clean

    source = pd.read_excel(source_path, dtype=object) if source_path.suffix.lower() == ".xlsx" else pd.read_csv(source_path, dtype=object)
    # Keep the original workbook structure. Rebuilding this file through
    # pandas would rename duplicate blank headers to Unnamed:N and rewrite
    # empty cells, which defeats the non-target preservation guarantee.
    reviewed_book = load_workbook(system_path)
    reviewed_sheet = reviewed_book[reviewed_book.sheetnames[0]]
    reviewed_headers = {reviewed_sheet.cell(1, column).value: column for column in range(1, reviewed_sheet.max_column + 1) if reviewed_sheet.cell(1, column).value}
    for target in target_fields:
        if target not in reviewed_headers:
            reviewed_headers[target] = reviewed_sheet.max_column + 1
            reviewed_sheet.cell(1, reviewed_headers[target]).value = target
    logs = pd.read_excel(log_path, dtype=object)
    if "record_id" in source.columns:
        source_ids = source["record_id"].astype(str).tolist()
    elif debug_records is not None and len(debug_records) == len(source) and "record_id" in debug_records:
        # Pipeline-derived IDs are stable for the current source file and are
        # also present in the modification log, even when the original file
        # had no record_id column.
        source_ids = debug_records["record_id"].astype(str).tolist()
    else:
        raise ValueError("无法安全关联人工复核行：缺少 record_id 且 debug 行数不一致")
    for row_index, record_id in enumerate(source_ids):
        if record_id not in choices:
            continue
        action, entity_id = choices[record_id]
        selected = codebook.row_for_entity(entity_id) if entity_id else None
        if action == "pick" and selected is None:
            raise ValueError(f"人工选择的 entity_id 不存在或不唯一: {entity_id}")
        log_mask = logs["record_id"].astype(str) == record_id
        for target in target_fields:
            before = _clean(source.iloc[row_index].get(target))
            old_value = _clean(reviewed_sheet.cell(row_index + 2, reviewed_headers[target]).value)
            if action == "none":
                final = None
                change_type = "cleared"
            elif selected is not None:
                final = getattr(selected, field_mapping[target])
                change_type = "human_update" if before != final else "no_change"
            else:
                final = old_value
                change_type = "no_change"
            reviewed_sheet.cell(row_index + 2, reviewed_headers[target]).value = final
            if log_mask.any():
                positions = logs.index[log_mask]
                column_mask = logs.loc[positions, "column_name"].astype(str) == target
                for position in positions[column_mask]:
                    logs.at[position, "final_value"] = final
                    logs.at[position, "change_type"] = change_type
                    logs.at[position, "change_source"] = "human"
                    logs.at[position, "entity_id"] = entity_id or ""
                    logs.at[position, "operator"] = operator
                    logs.at[position, "ui_session_id"] = session_id
                    logs.at[position, "operated_at"] = pd.Timestamp.utcnow().isoformat()
                    logs.at[position, "reason"] = f"Streamlit action={action}"
    output_dir = Path(tempfile.mkdtemp(prefix="department_review_"))
    reviewed_out = output_dir / "reviewed_content.xlsx"
    log_out = output_dir / "modification_log.xlsx"
    reviewed_book.save(reviewed_out)
    logs.to_excel(log_out, index=False)
    Pipeline._validate_preserved_content(source_path, reviewed_out, target_fields)
    return reviewed_out, log_out


def run_app() -> None:
    try:
        import streamlit as st
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("review-ui 需要安装 streamlit") from exc
    try:
        from .codebook import Codebook
        from .llm_judge import DeepSeekJudge, LLMJudge
        from .pipeline import Pipeline
    except ImportError:  # Streamlit executes this file as a script.
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
        from src.engine.codebook import Codebook
        from src.engine.llm_judge import DeepSeekJudge, LLMJudge
        from src.engine.pipeline import Pipeline

    st.set_page_config(page_title="部门名称编码工作台", page_icon="🧭", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""
    <style>
    [data-testid="stSidebar"] {background: linear-gradient(180deg,#f4f7fb 0%,#eef3f8 100%);}
    .hero {padding: 1.4rem 1.6rem; border-radius: 18px; background: linear-gradient(120deg,#17324d,#286b8f); color:white; margin-bottom:1rem;}
    .hero h1 {margin:0; font-size:2rem;} .hero p {margin:.35rem 0 0; opacity:.86;}
    .hint {padding: .85rem 1rem; border-left: 4px solid #2c7a9b; background:#f2f7fa; border-radius:8px;}
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="hero"><h1>🧭 部门名称编码工作台</h1><p>标准化匹配、候选裁决与人工复核</p></div>', unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("### 运行配置")
        st.caption("文件只在本机处理；DeepSeek API Key 仅保存在本次会话内存中。")
        repo_root = Path(__file__).resolve().parents[2]
        default_codebook = repo_root / "codebook.xlsx"
        default_codebook_text = str(default_codebook)
        codebook_path = _resolve_path(st.text_input("Codebook 路径", default_codebook_text), "codebook.xlsx")
        config_path = _resolve_path(st.text_input("Config 路径", str(repo_root / "config.yaml")), "config.yaml")
        model_text = st.text_input("模型路径（可留空）", "models/synthetic_ranker.joblib")
        operator = st.text_input("审核人", "")
        llm_provider = st.selectbox("二次裁决", ["不使用外部 LLM", "DeepSeek API"], index=0)
        uploaded = st.file_uploader("上传输入 Excel/CSV", type=["xlsx", "xls", "csv"])
    deepseek_key = ""
    deepseek_model = "deepseek-chat"
    deepseek_base_url = "https://api.deepseek.com"
    deepseek_timeout = 30
    llm_budget = 100
    if llm_provider == "DeepSeek API":
        deepseek_key = st.text_input("DeepSeek API Key", type="password", placeholder="sk-…")
        deepseek_model = st.text_input("DeepSeek 模型", value="deepseek-chat")
        deepseek_base_url = st.text_input("API Base URL", value="https://api.deepseek.com")
        deepseek_timeout = st.number_input("请求超时（秒）", min_value=5, max_value=120, value=30, step=5)
        llm_budget = st.number_input("本次最多调用次数", min_value=1, max_value=10000, value=100, step=10)
    if uploaded is None:
        st.markdown('<div class="hint"><b>使用流程</b><br>① 检查 codebook 路径　② 上传数据　③ 选择是否启用 DeepSeek　④ 点击运行　⑤ 查看候选并人工复核</div>', unsafe_allow_html=True)
        return
    temp_dir = Path(tempfile.gettempdir()) / "department_matcher_streamlit"
    temp_dir.mkdir(parents=True, exist_ok=True)
    input_path = temp_dir / uploaded.name
    input_path.write_bytes(uploaded.getbuffer())
    run_clicked = st.button("▶ 开始运行", type="primary", use_container_width=True)
    if run_clicked:
        if not codebook_path.exists():
            st.error(f"找不到 codebook: {codebook_path}")
            return
        if llm_provider == "DeepSeek API":
            if not deepseek_key.strip():
                st.error("请输入 DeepSeek API Key，或选择“不使用外部 LLM”。")
                return
            try:
                judge = LLMJudge(
                    model=deepseek_model,
                    judge_fn=DeepSeekJudge(deepseek_key, model=deepseek_model, base_url=deepseek_base_url, timeout=float(deepseek_timeout)),
                    cache_path=temp_dir / "deepseek_llm_cache.jsonl",
                    budget=int(llm_budget),
                )
            except ValueError as exc:
                st.error(str(exc))
                return
        else:
            judge = LLMJudge(model="fake", cache_path=temp_dir / "fake_llm_cache.jsonl", budget=0)
        pipeline = Pipeline(codebook_path, config_path=config_path, model_path=model_text or None, llm_judge=judge)
        progress_bar = st.progress(0, text="准备运行…")
        progress_text = st.empty()
        def on_progress(stage: str, done: int, total: int) -> None:
            ratio = min(1.0, done / total) if total else 0.0
            progress_bar.progress(ratio, text=f"{stage}：{done}/{total}")
            progress_text.caption(f"当前阶段：{stage}")
        try:
            result = pipeline.run(input_path, temp_dir / "runs", progress_callback=on_progress)
        except Exception as exc:
            st.error("运行失败，程序已停止。请查看下面的错误信息。")
            st.exception(exc)
            return
        st.session_state["matcher_result"] = result
        st.session_state["matcher_input"] = str(input_path)
    elif "matcher_result" not in st.session_state:
        st.info("点击“运行匹配”生成系统建议")
        return
    result = st.session_state["matcher_result"]
    pipeline = Pipeline(codebook_path, config_path=config_path, model_path=model_text or None)

    debug_path = Path(result["output_dir"]) / "debug" / "queries.jsonl"
    debug = _read_debug(debug_path)
    tab_help, tab_result, tab_review, tab_export = st.tabs(["说明", "运行概览", "人工复核", "导出文件"])
    with tab_help:
        st.markdown("""### 这套程序如何判断
        系统先标准化原始名称，再从 codebook 召回候选，用规则和排序模型筛选。高置信结果自动写入；存在冲突的记录进入候选裁决或人工复核。DeepSeek 只能在候选白名单中选择，无法确定时会保留人工复核状态。

        ### 如何使用
        上传文件需要包含 `w_name_o`，并建议提供 `cityname`、时间和类型字段。人工复核时只选择 `entity_id`，不要直接手写标准部门字段。导出前系统会检查非目标列是否被修改。
        """)
    with tab_result:
        st.success(f"已完成：{result['input_rows']} 行，去重查询 {result['unique_queries']} 条")
        st.metric("输入行数", f"{result['input_rows']:,}")
        st.metric("去重查询", f"{result['unique_queries']:,}")
    with tab_review:
        page_size = st.number_input("每页行数", min_value=10, max_value=200, value=50, step=10, key="page_size")
    with tab_export:
        st.info("系统输出文件已生成后，可在下方下载。人工修改只会写入目标字段。")
    page = st.number_input("页码", min_value=1, max_value=max(1, (len(debug) + int(page_size) - 1) // int(page_size)), value=1, step=1)
    visible = debug.iloc[(int(page) - 1) * int(page_size) : int(page) * int(page_size)]
    st.dataframe(visible[[column for column in ["record_id", "raw_name", "cityname", "decision", "entity_id", "method", "p_calibrated", "margin", "reason"] if column in visible.columns]], use_container_width=True)

    choices: dict[str, tuple[str, str]] = {}
    st.subheader("人工复核（只能选择 entity_id，不允许手写正式字段）")
    for _, row in visible.iterrows():
        record_id = str(row.get("record_id", ""))
        options = _candidate_options(row.get("candidate_ids")) or _candidate_options(row.get("top3_candidates"))
        action = st.selectbox(f"{record_id} | {row.get('raw_name', '')}", ["confirm", "pick", "none", "hard"], key=f"action_{record_id}")
        selected = ""
        if action in {"pick", "hard"}:
            selected = st.selectbox("entity_id", [""] + options, key=f"entity_{record_id}")
        choices[record_id] = ("pick" if action in {"pick", "hard"} else action, selected)
    if st.button("导出人工复核结果"):
        session_id = uuid.uuid4().hex
        reviewed, log = _apply_human_choices(Path(st.session_state["matcher_input"]), Path(result["reviewed_content"]), Path(result["modification_log"]), pipeline.codebook, choices, pipeline.target_fields, pipeline.field_mapping, session_id, operator, debug_records=debug)
        st.session_state["human_exports"] = (str(reviewed), str(log))
    if "human_exports" in st.session_state:
        reviewed, log = st.session_state["human_exports"]
        st.download_button("下载 reviewed_content.xlsx", Path(reviewed).read_bytes(), file_name="reviewed_content.xlsx")
        st.download_button("下载 modification_log.xlsx", Path(log).read_bytes(), file_name="modification_log.xlsx")


if __name__ == "__main__":
    run_app()
