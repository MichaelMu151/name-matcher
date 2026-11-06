"""End-to-end batch matching and safe two-file export."""

from __future__ import annotations

import hashlib
import json
import shutil
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from .candidates import generate_candidates
from .codebook import Codebook, CodebookRow, normalize_name
from .config import load_config
from .features import build_feature_rows
from .index import NameIndex
from .llm_judge import JudgeError, LLMJudge
from .normalize import normalize_name as normalize_runtime
from .policy import Decision, decide
from .ranker import Ranker


TARGET_FIELDS = ("w_name", "wrank_t", "s1_sector")


def _clean(value: Any) -> Any:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def _version(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()[:12]


class Pipeline:
    def __init__(self, codebook_path: str | Path, config_path: str | Path = "config.yaml", model_path: str | Path | None = None, llm_judge: LLMJudge | None = None):
        config_file = Path(config_path)
        self.config_dir = config_file.parent.resolve()
        self.config = load_config(config_path)
        configured_targets = self.config.get("audited_fields", TARGET_FIELDS)
        if not isinstance(configured_targets, (list, tuple)) or not configured_targets:
            raise ValueError("config.audited_fields 必须是非空列表")
        self.target_fields = tuple(str(field) for field in configured_targets)
        self.field_mapping = dict(self.config.get("field_mapping", {}))
        missing_mapping = [field for field in self.target_fields if field not in self.field_mapping]
        if missing_mapping:
            raise ValueError(f"field_mapping 缺少 audited_fields: {missing_mapping}")
        self.codebook_path = Path(codebook_path)
        reports_dir = Path(self.config["paths"]["reports_dir"])
        if not reports_dir.is_absolute():
            reports_dir = self.config_dir / reports_dir
        self.resource_dir = Path(self.config["paths"]["resource_dir"])
        if not self.resource_dir.is_absolute():
            self.resource_dir = self.config_dir / self.resource_dir
        self.codebook = Codebook(codebook_path, reports_dir=reports_dir)
        self.index = NameIndex(self.codebook)
        self.ranker = Ranker.load(model_path) if model_path and Path(model_path).exists() else Ranker()
        self.llm_judge = llm_judge or LLMJudge(model=self.config["models"].get("llm_model", "fake"), budget=int(self.config["models"].get("llm_budget", 100)))

    @staticmethod
    def _record_id(row: pd.Series, occurrence: int) -> str:
        payload = json.dumps({str(k): _clean(v) for k, v in row.items()}, ensure_ascii=False, sort_keys=True, default=str)
        return hashlib.sha256(f"{payload}|{occurrence}".encode("utf-8")).hexdigest()[:16]

    def _query_record(self, row: pd.Series, record_id: str) -> dict[str, Any]:
        input_mapping = self.config.get("input_mapping", {})
        dtype_field = str(input_mapping.get("dtype", "dtype"))
        dtype_fallback = str(input_mapping.get("dtype_fallback", "dtype_t"))
        dtype = _clean(row.get(dtype_field)) or _clean(row.get(dtype_fallback)) or _clean(row.get("dtype20"))
        context_parts = []
        context_fields = input_mapping.get("context_fields", ["p_name", "post1", "context"])
        for field in context_fields:
            value = _clean(row.get(field))
            if value:
                context_parts.append(f"{field}={value}")
        return {
            "record_id": record_id,
            "year": _clean(row.get(input_mapping.get("year", "year"))),
            "styr": _clean(row.get("styr")),
            "stmo": _clean(row.get("stmo")),
            "fiyr": _clean(row.get("fiyr")),
            "fimo": _clean(row.get("fimo")),
            "cityname": str(_clean(row.get("cityname")) or "").strip(),
            "distname": str(_clean(row.get("distname")) or "").strip(),
            "dtype": dtype,
            "w_name_o": str(_clean(row.get(input_mapping.get("w_name_o", "w_name_o"))) or ""),
            "context": "；".join(context_parts),
        }

    def _process_query(self, query: dict[str, Any]) -> tuple[Decision, list[dict[str, Any]], list[str]]:
        normalized = normalize_runtime(query["w_name_o"], self.resource_dir / "admin_divisions.txt", self.resource_dir / "suffixes.txt")
        candidates, flags = generate_candidates(
            query,
            self.index,
            normalized,
            top_k=int(self.config["matching"]["candidate_k"]),
            distname_hard_filter=bool(self.config["matching"].get("distname_hard_filter", False)),
            month_tolerance=int(self.config["matching"].get("month_tolerance", 0)),
            year_tolerance=int(self.config["matching"].get("year_tolerance", 0)),
        )
        feature_rows = build_feature_rows(normalized, candidates)
        probabilities = self.ranker.predict_proba(feature_rows).tolist()
        decision = decide(candidates, probabilities, flags, tau_auto=float(self.config["policy"]["tau_auto"]), margin_auto=float(self.config["policy"]["margin_auto"]), tau_nil=float(self.config["policy"]["tau_nil"]))
        if decision.decision == "LLM":
            candidate_payload = [
                {
                    "entity_id": candidate.entity_id,
                    "w_name": candidate.row.w_name,
                    "s1_sector": candidate.row.s1_sector,
                    "wrank": candidate.row.wrank,
                    "year_valid_from": candidate.row.year_valid_from,
                    "month_valid_from": candidate.row.month_valid_from,
                    "year_valid_to": candidate.row.year_valid_to,
                    "month_valid_to": candidate.row.month_valid_to,
                    "notes": candidate.row.notes[:120],
                    "matched_alias": candidate.matched_variant,
                    "flags": candidate.flags,
                }
                for candidate in candidates
            ]
            try:
                verdict = self.llm_judge.judge(query, candidate_payload)
                if verdict.uncertain or verdict.decision in {"NONE", "MULTIPLE", "NOT_A_DEPARTMENT"}:
                    reason = verdict.reason or "LLM 未能安全选择唯一候选"
                    decision = Decision("HUMAN", None, "llm", decision.p_calibrated, decision.margin, reason, decision.flags)
                elif any(candidate.entity_id == verdict.decision for candidate in candidates):
                    decision = Decision("AUTO", verdict.decision, "llm", decision.p_calibrated, decision.margin, verdict.reason, decision.flags)
                else:
                    decision = Decision("HUMAN", None, "llm", decision.p_calibrated, decision.margin, "LLM 返回非法候选", decision.flags)
            except JudgeError as exc:
                decision = Decision("HUMAN", None, "llm", decision.p_calibrated, decision.margin, str(exc), decision.flags)
        debug = []
        for index, candidate in enumerate(candidates):
            row = dict(feature_rows[index]) if index < len(feature_rows) else {}
            row.update({"entity_id": candidate.entity_id, "dept_key": candidate.dept_key, "w_name": candidate.row.w_name, "wrank": candidate.row.wrank, "s1_sector": candidate.row.s1_sector, "method": candidate.retrieval_path, "probability": probabilities[index] if index < len(probabilities) else 0.0, "flags": candidate.flags})
            debug.append(row)
        return decision, debug, flags

    def run(self, input_path: str | Path, output_dir: str | Path = "outputs", run_id: str | None = None) -> dict[str, Any]:
        input_path = Path(input_path)
        run_id = run_id or datetime.now(timezone.utc).strftime("run_%Y%m%d_%H%M%S")
        output_dir = Path(output_dir) / run_id
        output_dir.mkdir(parents=True, exist_ok=True)
        if input_path.suffix.lower() == ".csv":
            frame = pd.read_csv(input_path, dtype=object)
        else:
            frame = pd.read_excel(input_path, dtype=object)
        occurrence: Counter[str] = Counter()
        input_mapping = self.config.get("input_mapping", {})
        record_id_field = str(input_mapping.get("record_id", "record_id"))
        queries: list[dict[str, Any]] = []
        query_keys: list[tuple[Any, ...]] = []
        for _, row in frame.iterrows():
            record_id = str(_clean(row.get(record_id_field)) or self._record_id(row, occurrence["row"]))
            occurrence["row"] += 1
            query = self._query_record(row, record_id)
            normalized = normalize_runtime(query["w_name_o"], self.resource_dir / "admin_divisions.txt", self.resource_dir / "suffixes.txt")
            key_parts: list[Any] = [query["styr"], query["stmo"], query["fiyr"], query["fimo"], query["cityname"]]
            if bool(self.config["matching"].get("dedupe_include_distname", True)):
                key_parts.append(query["distname"])
            key_parts.append(normalized.norm)
            key = tuple(key_parts)
            queries.append(query)
            query_keys.append(key)

        unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]] = {}
        logs: list[dict[str, Any]] = []
        debug_rows: list[dict[str, Any]] = []
        for query, key in zip(queries, query_keys):
            if key not in unique:
                unique[key] = self._process_query(query)
            decision, debug, flags = unique[key]
            debug_rows.append({
                "record_id": query["record_id"],
                "query_key": "|".join("" if value is None else str(value) for value in key),
                "raw_name": query["w_name_o"],
                "norm_name": normalize_runtime(query["w_name_o"]).norm,
                "year": query["year"],
                "styr": query["styr"],
                "stmo": query["stmo"],
                "fiyr": query["fiyr"],
                "fimo": query["fimo"],
                "cityname": query["cityname"],
                "distname": query["distname"],
                "decision": decision.decision,
                "entity_id": decision.entity_id or "",
                "method": decision.method,
                "p_calibrated": decision.p_calibrated,
                "margin": decision.margin,
                "reason": decision.reason,
                "flags": flags,
                "top3_candidates": debug[:3],
                "candidate_ids": [candidate.get("entity_id", "") for candidate in debug],
            })

        result = frame.copy()
        for target in self.target_fields:
            if target not in result.columns:
                result[target] = ""
        logs.extend(self._build_logs(frame, queries, query_keys, unique, run_id))
        self._apply_decisions(result, queries, query_keys, unique)
        reviewed = output_dir / self.config["ui"]["export_files"]["content"]
        log_path = output_dir / self.config["ui"]["export_files"]["log"]
        self._export_content(input_path, reviewed, result, self.target_fields)
        self._validate_preserved_content(input_path, reviewed, self.target_fields)
        pd.DataFrame(logs).to_excel(log_path, index=False)
        (output_dir / "debug").mkdir(exist_ok=True)
        pd.DataFrame(debug_rows).to_json(output_dir / "debug" / "queries.jsonl", orient="records", lines=True, force_ascii=False)
        meta = {
            "run_id": run_id,
            "input": str(input_path),
            "input_version": _version(input_path),
            "codebook": str(self.codebook_path),
            "codebook_version": _version(self.codebook_path),
            "model_version": self.ranker.version,
            "model_metrics": self.ranker.metrics,
            "config": self.config,
            "audited_fields": list(self.target_fields),
            "field_mapping": self.field_mapping,
            "unique_queries": len(unique),
            "input_rows": len(frame),
        }
        (output_dir / "run_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return {"run_id": run_id, "output_dir": str(output_dir), "input_rows": len(frame), "unique_queries": len(unique), "reviewed_content": str(reviewed), "modification_log": str(log_path)}

    @staticmethod
    def _selected_row(decision: Decision, codebook: Codebook) -> CodebookRow | None:
        return codebook.row_for_entity(decision.entity_id) if decision.entity_id else None

    def _apply_decisions(self, result: pd.DataFrame, queries: list[dict[str, Any]], keys: list[tuple[Any, ...]], unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]]) -> None:
        for index, key in enumerate(keys):
            decision = unique[key][0]
            row = self._selected_row(decision, self.codebook)
            if row is None or decision.decision not in {"AUTO"}:
                continue
            for target in self.target_fields:
                source_field = self.field_mapping[target]
                result.at[index, target] = getattr(row, source_field)

    def _build_logs(self, frame: pd.DataFrame, queries: list[dict[str, Any]], keys: list[tuple[Any, ...]], unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]], run_id: str) -> list[dict[str, Any]]:
        logs: list[dict[str, Any]] = []
        for index, (query, key) in enumerate(zip(queries, keys)):
            decision = unique[key][0]
            row = self._selected_row(decision, self.codebook)
            values = {
                target: (getattr(row, self.field_mapping[target]) if row and decision.decision == "AUTO" else _clean(frame.iloc[index].get(target)))
                for target in self.target_fields
            }
            for column in self.target_fields:
                before = _clean(frame.iloc[index].get(column))
                after = values[column]
                added = column not in frame.columns
                change_type = "column_added" if added else ("auto_update" if before != after else "no_change")
                logs.append({"batch_id": run_id, "run_id": run_id, "record_id": query["record_id"], "column_name": column, "before_value": before, "system_suggested_value": after, "final_value": after, "change_type": change_type, "change_source": "system", "entity_id": decision.entity_id or "", "method": decision.method, "p_calibrated": decision.p_calibrated, "margin": decision.margin, "operator": "", "operated_at": datetime.now(timezone.utc).isoformat(), "reason": decision.reason, "ui_session_id": "", "codebook_version": _version(self.codebook_path), "model_version": self.ranker.version})
        return logs

    @staticmethod
    def _export_content(input_path: Path, output_path: Path, result: pd.DataFrame, target_fields: tuple[str, ...] | None = None) -> None:
        target_fields = target_fields or TARGET_FIELDS
        if input_path.suffix.lower() != ".xlsx":
            result.to_excel(output_path, index=False)
            return
        shutil.copy2(input_path, output_path)
        workbook = load_workbook(output_path)
        sheet = workbook[workbook.sheetnames[0]]
        headers = {sheet.cell(1, col).value: col for col in range(1, sheet.max_column + 1) if sheet.cell(1, col).value}
        next_col = sheet.max_column + 1
        for target in target_fields:
            if target not in headers:
                headers[target] = next_col
                sheet.cell(1, headers[target]).value = target
                next_col += 1
        for row_index, (_, row) in enumerate(result.iterrows(), start=2):
            for target in target_fields:
                sheet.cell(row_index, headers[target]).value = _clean(row.get(target))
        workbook.save(output_path)

    @staticmethod
    def _validate_preserved_content(input_path: Path, output_path: Path, target_fields: tuple[str, ...]) -> None:
        """Fail closed if any non-audited cell changed during export."""

        def read(path: Path) -> pd.DataFrame:
            if path.suffix.lower() == ".csv":
                return pd.read_csv(path, dtype=object)
            return pd.read_excel(path, dtype=object)

        source = read(input_path)
        exported = read(output_path)
        if len(source) != len(exported):
            raise ValueError("导出安全校验失败：行数发生变化")
        source_non_target = [column for column in source.columns if column not in target_fields]
        exported_non_target = [column for column in exported.columns if column not in target_fields]
        if source_non_target != exported_non_target:
            raise ValueError("导出安全校验失败：非目标列集合或顺序发生变化")
        for column in source_non_target:
            for row_index, (left, right) in enumerate(zip(source[column].tolist(), exported[column].tolist()), start=2):
                if _clean(left) != _clean(right):
                    raise ValueError(f"导出安全校验失败：{column}!{row_index} 被修改")
