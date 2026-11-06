"""Evaluation and export-safety metrics for matcher runs."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd


def _read_table(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=object)
    return pd.read_excel(path, dtype=object)


def _candidate_ids(value: Any) -> list[str]:
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
            return _candidate_ids(json.loads(value))
        except (TypeError, json.JSONDecodeError):
            return []
    return []


def _ece(probabilities: list[float], labels: list[int], bins: int = 10) -> float:
    if not probabilities:
        return 0.0
    total = len(probabilities)
    error = 0.0
    for index in range(bins):
        low, high = index / bins, (index + 1) / bins
        members = [position for position, value in enumerate(probabilities) if low <= value < high or (index == bins - 1 and value == high)]
        if not members:
            continue
        confidence = sum(probabilities[position] for position in members) / len(members)
        accuracy = sum(labels[position] for position in members) / len(members)
        error += len(members) / total * abs(confidence - accuracy)
    return float(error)


def evaluate_predictions(
    debug_path: str | Path,
    gold_path: str | Path | None = None,
    output_path: str | Path = "reports/eval.md",
    final: bool = False,
    input_path: str | Path | None = None,
    reviewed_path: str | Path | None = None,
    log_path: str | Path | None = None,
) -> dict[str, Any]:
    predictions = pd.read_json(debug_path, lines=True)
    result: dict[str, Any] = {
        "n_predictions": int(len(predictions)),
        "auto_rate": float((predictions.get("decision", pd.Series(dtype=str)) == "AUTO").mean()) if len(predictions) else 0.0,
        "llm_rate": float((predictions.get("method", pd.Series(dtype=str)) == "llm").mean()) if len(predictions) else 0.0,
        "human_rate": float((predictions.get("decision", pd.Series(dtype=str)) == "HUMAN").mean()) if len(predictions) else 0.0,
    }
    errors: list[dict[str, Any]] = []
    if gold_path:
        gold_path = Path(gold_path)
        if "test" in gold_path.stem.lower() and not final:
            raise ValueError("冻结 test 集只能通过 eval --final 读取")
        gold = _read_table(gold_path)
        if "record_id" not in gold or "record_id" not in predictions:
            raise ValueError("gold 和 debug 都必须包含 record_id")
        gold_cols = [column for column in ["entity_id", "w_name", "wrank_t", "s1_sector"] if column in gold]
        if "entity_id" not in gold_cols:
            raise ValueError("gold 必须包含 entity_id")
        merged = predictions.merge(gold[["record_id", *gold_cols]], on="record_id", suffixes=("", "_gold"))
        if len(merged):
            top1 = merged["entity_id"].fillna("").astype(str)
            gold_ids = merged["entity_id_gold"].fillna("").astype(str)
            result["top1_accuracy"] = float((top1 == gold_ids).mean())
            recalls: dict[str, float] = {}
            reciprocal: list[float] = []
            labels: list[int] = []
            probs: list[float] = []
            for _, row in merged.iterrows():
                candidates = _candidate_ids(row.get("candidate_ids")) or _candidate_ids(row.get("top3_candidates"))
                # Debug stores top-3; report Recall@10 when a full candidate
                # list is supplied by a future run, otherwise expose the
                # available top-3 honestly.
                for k in (1, 3, 5, 10):
                    recalls.setdefault(str(k), []).append(1.0 if str(row["entity_id_gold"]) in candidates[:k] else 0.0)
                if str(row["entity_id_gold"]) in candidates:
                    reciprocal.append(1.0 / (candidates.index(str(row["entity_id_gold"])) + 1))
                else:
                    reciprocal.append(0.0)
                labels.append(int(str(row["entity_id"]) == str(row["entity_id_gold"])))
                try:
                    probs.append(float(row.get("p_calibrated", 0.0)))
                except (TypeError, ValueError):
                    probs.append(0.0)
                if str(row["entity_id"]) != str(row["entity_id_gold"]):
                    errors.append({"record_id": row["record_id"], "predicted": row["entity_id"], "gold": row["entity_id_gold"], "decision": row.get("decision", ""), "method": row.get("method", "")})
            for key, values in recalls.items():
                result[f"recall_at_{key}"] = float(sum(values) / len(values))
            result["mrr"] = float(sum(reciprocal) / len(reciprocal))
            result["ece"] = _ece(probs, labels)
            gold_nil = gold_ids.isin(["", "NONE", "NIL"])
            pred_nil = top1.isin(["", "NONE", "MULTIPLE", "NOT_A_DEPARTMENT"])
            result["nil_precision"] = float(((pred_nil & gold_nil).sum() / max(1, pred_nil.sum())))
            result["nil_recall"] = float(((pred_nil & gold_nil).sum() / max(1, gold_nil.sum())))
            for field in ["w_name", "wrank_t", "s1_sector"]:
                if field in merged and f"{field}_gold" in merged:
                    result[f"{field}_accuracy"] = float((merged[field].fillna("").astype(str) == merged[f"{field}_gold"].fillna("").astype(str)).mean())

    if input_path and reviewed_path:
        source, reviewed = _read_table(input_path), _read_table(reviewed_path)
        targets = {"w_name", "wrank_t", "s1_sector"}
        non_targets = [column for column in source.columns if column not in targets]
        unchanged = all(source[column].fillna("<NA>").astype(str).equals(reviewed[column].fillna("<NA>").astype(str)) for column in non_targets if column in reviewed)
        result["non_target_zero_change"] = bool(unchanged)
    if log_path:
        log = _read_table(log_path)
        required = {"record_id", "column_name", "before_value", "final_value", "entity_id", "change_type"}
        result["log_required_columns"] = bool(required.issubset(log.columns))
        result["log_rows"] = int(len(log))

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    lines = ["# Evaluation", "", *[f"- {key}: {value}" for key, value in result.items()]]
    output.write_text("\n".join(lines) + "\n", encoding="utf-8")
    if errors:
        pd.DataFrame(errors).to_csv(output.parent / "errors.csv", index=False)
    return result
