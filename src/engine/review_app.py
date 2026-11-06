"""Small but complete local Streamlit review workflow."""

from __future__ import annotations

import json
import tempfile
import uuid
from pathlib import Path
from typing import Any

import pandas as pd


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


def _apply_human_choices(source_path: Path, system_path: Path, log_path: Path, codebook: Any, choices: dict[str, tuple[str, str]], target_fields: tuple[str, ...], field_mapping: dict[str, str], session_id: str, operator: str, debug_records: pd.DataFrame | None = None) -> tuple[Path, Path]:
    """Apply only entity-id selections and return revised content/log files."""

    from .pipeline import Pipeline, _clean

    source = pd.read_excel(source_path, dtype=object) if source_path.suffix.lower() == ".xlsx" else pd.read_csv(source_path, dtype=object)
    reviewed = pd.read_excel(system_path, dtype=object)
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
            old_value = _clean(reviewed.iloc[row_index].get(target))
            if action == "none":
                final = None
                change_type = "cleared"
            elif selected is not None:
                final = getattr(selected, field_mapping[target])
                change_type = "human_update" if before != final else "no_change"
            else:
                final = old_value
                change_type = "no_change"
            reviewed.at[row_index, target] = final
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
    reviewed.to_excel(reviewed_out, index=False)
    logs.to_excel(log_out, index=False)
    Pipeline._validate_preserved_content(source_path, reviewed_out, target_fields)
    return reviewed_out, log_out


def run_app() -> None:
    try:
        import streamlit as st
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("review-ui 需要安装 streamlit") from exc
    from .codebook import Codebook
    from .pipeline import Pipeline

    st.set_page_config(page_title="Department matcher review", layout="wide")
    st.title("政府部门名称编码复核")
    codebook_path = Path(st.text_input("Codebook 路径", "codebook.xlsx"))
    config_path = Path(st.text_input("Config 路径", "config.yaml"))
    model_text = st.text_input("模型路径（可留空，使用冷启动规则）", "models/synthetic_ranker.joblib")
    operator = st.text_input("审核人", "")
    uploaded = st.file_uploader("上传输入 Excel/CSV", type=["xlsx", "xls", "csv"])
    if uploaded is None:
        st.info("上传文件后运行匹配")
        return
    temp_dir = Path(tempfile.gettempdir()) / "department_matcher_streamlit"
    temp_dir.mkdir(parents=True, exist_ok=True)
    input_path = temp_dir / uploaded.name
    input_path.write_bytes(uploaded.getbuffer())
    if st.button("运行匹配", type="primary"):
        if not codebook_path.exists():
            st.error(f"找不到 codebook: {codebook_path}")
            return
        pipeline = Pipeline(codebook_path, config_path=config_path, model_path=model_text or None)
        result = pipeline.run(input_path, temp_dir / "runs")
        st.session_state["matcher_result"] = result
        st.session_state["matcher_input"] = str(input_path)
    elif "matcher_result" not in st.session_state:
        st.info("点击“运行匹配”生成系统建议")
        return
    result = st.session_state["matcher_result"]
    pipeline = Pipeline(codebook_path, config_path=config_path, model_path=model_text or None)

    debug_path = Path(result["output_dir"]) / "debug" / "queries.jsonl"
    debug = _read_debug(debug_path)
    st.success(f"已完成：{result['input_rows']} 行，去重查询 {result['unique_queries']} 条")
    page_size = st.number_input("每页行数", min_value=10, max_value=200, value=50, step=10)
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
