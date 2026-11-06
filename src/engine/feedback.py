"""Human-review export and conservative feedback ingestion."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd


def _read(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    return pd.read_excel(path, dtype=object) if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"} else pd.read_csv(path, dtype=object)


def _json_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    return json.dumps(value, ensure_ascii=False, default=str)


def export_review(input_path: str | Path, debug_path: str | Path, output_path: str | Path) -> Path:
    """Create a review workbook without changing source columns.

    The debug file is keyed by ``record_id`` whenever available. Positional
    fallback is rejected when lengths differ, preventing silent row drift.
    """

    source = _read(input_path)
    debug = pd.read_json(debug_path, lines=True)
    if "record_id" in source.columns and "record_id" in debug.columns:
        review = source.merge(
            debug[[column for column in ["record_id", "decision", "entity_id", "reason", "flags", "top3_candidates", "candidate_ids"] if column in debug.columns]],
            on="record_id",
            how="left",
            sort=False,
            suffixes=("", "_debug"),
            validate="one_to_one",
        )
    else:
        if len(source) != len(debug):
            raise ValueError("review-export 无 record_id 且输入/debug 行数不一致，拒绝按位置合并")
        review = source.copy()
        for column in ["decision", "entity_id", "reason", "flags", "top3_candidates", "candidate_ids"]:
            review[column] = debug[column].tolist() if column in debug else ""
    rename = {"decision": "system_decision", "entity_id": "system_entity_id", "reason": "llm_reason", "flags": "system_flags", "top3_candidates": "top3_candidates", "candidate_ids": "candidate_entity_ids"}
    review = review.rename(columns={key: value for key, value in rename.items() if key in review.columns})
    for column in ["system_decision", "system_entity_id", "llm_reason", "system_flags", "top3_candidates", "candidate_entity_ids"]:
        if column not in review.columns:
            review[column] = ""
    review["top3_candidates"] = review["top3_candidates"].map(_json_text)
    review["candidate_entity_ids"] = review["candidate_entity_ids"].map(_json_text)
    review["human_decision"] = ""
    review["human_entity_id"] = ""
    review["human_note"] = ""
    review["reviewer"] = ""
    review["reviewed_at"] = ""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    review.to_excel(output_path, index=False)
    return output_path


def ingest_feedback(
    feedback_path: str | Path,
    aliases_path: str | Path = "data/aliases_learned.csv",
    operator: str = "",
    second_reviewer: str = "",
    batch_id: str | None = None,
    labels_path: str | Path | None = None,
    cases_path: str | Path | None = None,
) -> Path:
    """Ingest review decisions; aliases remain pending until a second review.

    ``confirm``/``pick`` rows enter a pending alias/label pool. A row becomes
    active only when ``second_review_status=approved`` and a second reviewer
    is recorded. This prevents a single mistaken click from changing future
    exact-match behavior.
    """

    frame = _read(feedback_path)
    required = {"w_name_o", "human_decision", "human_entity_id"}
    missing = required - set(frame.columns)
    if missing:
        raise ValueError(f"feedback 缺少列: {sorted(missing)}")
    accepted = frame[frame["human_decision"].astype(str).isin({"confirm", "pick"})].copy()
    if "system_entity_id" in accepted.columns:
        confirm_mask = accepted["human_decision"].astype(str).eq("confirm")
        blank_human = accepted["human_entity_id"].fillna("").astype(str).str.strip().eq("")
        accepted.loc[confirm_mask & blank_human, "human_entity_id"] = accepted.loc[confirm_mask, "system_entity_id"]
    accepted = accepted[accepted["human_entity_id"].fillna("").astype(str).str.strip() != ""]
    batch_id = batch_id or datetime.now(timezone.utc).strftime("batch_%Y%m%d_%H%M%S")
    now = datetime.now(timezone.utc).isoformat()
    new_rows = pd.DataFrame(
        {
            "batch_id": batch_id,
            "raw_name": accepted["w_name_o"].astype(str),
            "entity_id": accepted["human_entity_id"].astype(str),
            "operator": operator,
            "created_at": now,
            "second_reviewer": second_reviewer,
            "second_review_status": "approved" if second_reviewer else "pending_second_review",
            "status": "active" if second_reviewer else "pending_second_review",
        }
    )
    output_path = Path(aliases_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        old = pd.read_csv(output_path, dtype=object)
        new_rows = pd.concat([old, new_rows], ignore_index=True)
        new_rows = new_rows.drop_duplicates(subset=["batch_id", "raw_name", "entity_id"], keep="last")
    new_rows.to_csv(output_path, index=False)

    if labels_path:
        label_path = Path(labels_path)
        label_path.parent.mkdir(parents=True, exist_ok=True)
        labels = frame.copy()
        labels["batch_id"] = batch_id
        labels["operator"] = operator
        labels["created_at"] = now
        if label_path.exists():
            labels = pd.concat([pd.read_csv(label_path, dtype=object), labels], ignore_index=True)
        labels.to_csv(label_path, index=False)
    if cases_path:
        case_path = Path(cases_path)
        case_path.parent.mkdir(parents=True, exist_ok=True)
        cases = accepted.copy()
        cases["batch_id"] = batch_id
        cases["operator"] = operator
        cases["created_at"] = now
        if case_path.exists():
            cases = pd.concat([pd.read_csv(case_path, dtype=object), cases], ignore_index=True)
        cases.to_csv(case_path, index=False)
    return output_path
