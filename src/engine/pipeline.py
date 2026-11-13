"""End-to-end batch matching and safe two-file export."""

from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import load_workbook
from datetime import date, datetime

from .candidates import generate_candidates
from .codebook import Codebook, CodebookRow, normalize_name
from .config import load_config
from .features import build_feature_rows
from .index import NameIndex
from .llm_judge import JudgeError, LLMJudge
from .normalize import normalize_name as normalize_runtime
from .policy import Decision, decide
from .ranker import Ranker
from .sqlite_cache import QueryCache


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

    def run(self, input_path: str | Path, output_dir: str | Path = "outputs", run_id: str | None = None, progress_callback: Callable[[str, int, int], None] | None = None) -> dict[str, Any]:
        def progress(stage: str, done: int, total: int) -> None:
            if progress_callback is not None:
                progress_callback(stage, done, total)
        input_path = Path(input_path)
        run_id = run_id or datetime.now(timezone.utc).strftime("run_%Y%m%d_%H%M%S")
        output_dir = Path(output_dir) / run_id
        output_dir.mkdir(parents=True, exist_ok=True)
        cache_dir = Path(self.config["paths"].get("sqlite_cache_dir", "cache"))
        if not cache_dir.is_absolute():
            cache_dir = self.config_dir / cache_dir
        namespace_payload = {
            "codebook": _version(self.codebook_path),
            "model": self.ranker.version,
            "config": self.config,
        }
        namespace = hashlib.sha256(json.dumps(namespace_payload, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:24]
        query_cache = QueryCache(cache_dir / "query_cache.sqlite3", namespace)
        if input_path.suffix.lower() == ".csv":
            frame = pd.read_csv(input_path, dtype=object)
        else:
            frame = pd.read_excel(input_path, dtype=object)
        progress("读取输入文件", 1, 1)
        occurrence: Counter[str] = Counter()
        input_mapping = self.config.get("input_mapping", {})
        record_id_field = str(input_mapping.get("record_id", "record_id"))
        queries: list[dict[str, Any]] = []
        query_keys: list[tuple[Any, ...]] = []
        standard_w_names = {normalize_name(row.w_name) for row in self.codebook.valid_rows() if row.w_name}
        for row_number, (_, row) in enumerate(frame.iterrows(), start=1):
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
            if row_number == 1 or row_number == len(frame) or row_number % 250 == 0:
                progress("建立查询索引", row_number, len(frame))

        unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]] = {}
        logs: list[dict[str, Any]] = []
        debug_rows: list[dict[str, Any]] = []
        for query_number, (query, key) in enumerate(zip(queries, query_keys), start=1):
            if key not in unique:
                cached = query_cache.get(key)
                if cached is not None:
                    unique[key] = cached
                else:
                    processed = self._process_query(query)
                    unique[key] = processed
                    # LLM results are deliberately excluded from the durable
                    # cache. Rules and ranker decisions are stable under the
                    # namespace; external model decisions depend on the API
                    # key, prompt version and service state.
                    if processed[0].method != "llm":
                        query_cache.put(key, processed)
            decision, debug, flags = unique[key]
            source_row = frame.iloc[query_number - 1]
            existing = {
                "w_name": _clean(source_row.get("w_name")),
                "wrank_t": _clean(source_row.get("wrank_t")),
                "s1_sector": _clean(source_row.get("s1_sector")),
            }
            selected = self._selected_row(decision, self.codebook, query) if decision.decision == "AUTO" else None
            system = {
                "w_name": selected.w_name if selected else None,
                "wrank_t": selected.wrank if selected else None,
                "s1_sector": selected.s1_sector if selected else None,
            }
            issue_type = ""
            issue_fields: list[str] = []
            if existing["w_name"] not in (None, "") and normalize_name(existing["w_name"]) not in standard_w_names:
                issue_type = "错误输入"
            elif decision.decision == "AUTO":
                for field in ("w_name", "wrank_t", "s1_sector"):
                    if existing[field] not in (None, "") and str(existing[field]).strip() != str(system[field] or "").strip():
                        issue_fields.append(field)
                if issue_fields:
                    issue_type = "有冲突"
                elif any(existing[field] in (None, "") for field in ("w_name", "wrank_t", "s1_sector")):
                    issue_type = "缺失输入"
            else:
                issue_type = "多种可能"
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
                "existing_w_name": existing["w_name"],
                "existing_wrank_t": existing["wrank_t"],
                "existing_s1_sector": existing["s1_sector"],
                "issue_type": issue_type,
                "issue_fields": ";".join(issue_fields),
                "review_reason": (f"{issue_type}（字段：{';'.join(issue_fields)}）" if issue_type == "有冲突" and issue_fields else issue_type),
                "system_w_name": system["w_name"],
                "system_wrank_t": system["wrank_t"],
                "system_s1_sector": system["s1_sector"],
                "candidate_count": len(debug),
                "top1_probability": debug[0].get("probability", 0.0) if debug else 0.0,
                "top2_probability": debug[1].get("probability", 0.0) if len(debug) > 1 else 0.0,
                "flags": flags,
                "top3_candidates": debug[:3],
                "candidate_ids": [candidate.get("entity_id", "") for candidate in debug],
            })
            if query_number == 1 or query_number == len(queries) or query_number % 100 == 0:
                progress("匹配与候选裁决", query_number, len(queries))

        result = frame.copy()
        for target in self.target_fields:
            if target not in result.columns:
                result[target] = ""
        logs.extend(self._build_logs(frame, queries, query_keys, unique, run_id))
        # post_rank is deliberately deferred until after human review and the
        # temporal validity stage in the desktop workflow.
        self._apply_decisions(result, queries, query_keys, unique)
        reviewed = output_dir / self.config["ui"]["export_files"]["content"]
        log_path = output_dir / self.config["ui"]["export_files"]["log"]
        self._export_content(input_path, reviewed, result, self.target_fields)
        progress("导出 Excel", 1, 1)
        self._validate_preserved_content(input_path, reviewed, self.target_fields)
        progress("安全校验", 1, 1)
        pd.DataFrame(logs).to_excel(log_path, index=False)
        (output_dir / "debug").mkdir(exist_ok=True)
        debug_frame = pd.DataFrame(debug_rows)
        # Persist the two decisions needed by the review UI.  The UI must not
        # rescan the codebook every time the review page is opened.
        scope_pairs = set()
        wildcard_scope = []
        for code_row in self.codebook.valid_rows():
            cities = self.codebook.split_values(code_row.cityname)
            dists = self.codebook.split_values(code_row.distname)
            if not cities or not dists:
                wildcard_scope.append((set(normalize_name(city) for city in cities), set(normalize_name(dist) for dist in dists)))
            else:
                scope_pairs.update((normalize_name(city), normalize_name(dist)) for city in cities for dist in dists)
        if not debug_frame.empty:
            debug_frame["codebook_version"] = _version(self.codebook_path)
            debug_frame["model_version"] = str(self.ranker.version)
            debug_frame["pipeline_stage"] = "initial_matching"
            debug_frame["scope_match"] = [
                ((normalize_name(row.get("cityname", "")), normalize_name(row.get("distname", ""))) in scope_pairs)
                or any((not cities or normalize_name(row.get("cityname", "")) in cities) and (not dists or normalize_name(row.get("distname", "")) in dists) for cities, dists in wildcard_scope)
                for _, row in debug_frame.iterrows()
            ]
            group_fields = ["raw_name", "cityname", "distname", "existing_w_name", "existing_wrank_t", "existing_s1_sector", "issue_type", "issue_fields", "candidate_ids"]
            debug_frame["review_group_key"] = debug_frame.apply(
                lambda row: json.dumps([str(row.get(field, "") or "") for field in group_fields], ensure_ascii=False, sort_keys=True, default=str),
                axis=1,
            )
            debug_frame["review_group_count"] = debug_frame.groupby("review_group_key")["review_group_key"].transform("size")
        debug_frame.to_json(output_dir / "debug" / "queries.jsonl", orient="records", lines=True, force_ascii=False)
        debug_frame[debug_frame["issue_type"].astype(str).str.strip() != ""].to_excel(output_dir / "review_queue.xlsx", index=False)
        query_cache.close()
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
        return {"run_id": run_id, "output_dir": str(output_dir), "input_path": str(input_path), "input_rows": len(frame), "unique_queries": len(unique), "reviewed_content": str(reviewed), "modification_log": str(log_path), "review_queue": str(output_dir / "review_queue.xlsx")}

    @staticmethod
    def _selected_row(decision: Decision, codebook: Codebook, query: dict[str, Any] | None = None) -> CodebookRow | None:
        if not decision.entity_id:
            return None
        rows = codebook.rows_for_entity(decision.entity_id)
        if not rows:
            return None
        if query:
            try:
                year = int(float(query.get("year"))) if query.get("year") not in (None, "") else None
                month = int(float(query.get("stmo"))) if query.get("stmo") not in (None, "") else 6
            except (TypeError, ValueError):
                year, month = None, 6
            if year is not None:
                key = year * 12 + max(1, min(12, month)) - 1
                in_period = [row for row in rows if row.start_month <= key <= row.end_month]
                if in_period:
                    rows = in_period
        return rows[0]

    def _apply_decisions(self, result: pd.DataFrame, queries: list[dict[str, Any]], keys: list[tuple[Any, ...]], unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]]) -> None:
        for index, key in enumerate(keys):
            decision = unique[key][0]
            row = self._selected_row(decision, self.codebook, queries[index])
            if row is None or decision.decision not in {"AUTO"}:
                continue
            for target in self.target_fields:
                source_field = self.field_mapping[target]
                result.at[index, target] = getattr(row, source_field)

    def _build_logs(self, frame: pd.DataFrame, queries: list[dict[str, Any]], keys: list[tuple[Any, ...]], unique: dict[tuple[Any, ...], tuple[Decision, list[dict[str, Any]], list[str]]], run_id: str) -> list[dict[str, Any]]:
        logs: list[dict[str, Any]] = []
        for index, (query, key) in enumerate(zip(queries, keys)):
            decision = unique[key][0]
            row = self._selected_row(decision, self.codebook, queries[index])
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
        auxiliary_fields = [field for field in ["post_rank", "post_rank_original", "post_rank_adjustment_reason", "post_rank_adjustment_status"] if field in result.columns]
        for row_index, (_, row) in enumerate(result.iterrows(), start=2):
            for target in list(target_fields) + [field for field in auxiliary_fields if field not in target_fields]:
                if target not in headers:
                    headers[target] = next_col
                    sheet.cell(1, headers[target]).value = target
                    next_col += 1
                sheet.cell(row_index, headers[target]).value = _clean(row.get(target))
        workbook.save(output_path)

    @staticmethod
    def _validate_preserved_content(input_path: Path, output_path: Path, target_fields: tuple[str, ...]) -> None:
        """Fail closed if any non-audited cell changed during export."""

        # These columns are explicitly generated by the post_rank linkage
        # feature.  They are audited outputs, not accidental modifications to
        # the user's original non-target columns.
        generated_fields = {"post_rank", "post_rank_original", "post_rank_adjustment_reason", "post_rank_adjustment_status"}

        # For Excel inputs compare semantic cell values with openpyxl.  The
        # previous XML-level comparison was too strict: Excel formulas carry
        # cached results which openpyxl may drop on save, and shared strings,
        # blank cells, and formula storage forms can legitimately change
        # without changing the user's data.  Formula cells are compared by
        # formula text only; cached results are deliberately ignored.
        if input_path.suffix.lower() == ".xlsx" and output_path.suffix.lower() == ".xlsx":
            def semantic(value: Any, cell: Any) -> tuple[str, str]:
                if getattr(cell, "data_type", "") == "f" or (isinstance(value, str) and value.startswith("=")):
                    return ("formula", str(value).strip())
                if value is None:
                    return ("none", "")
                try:
                    if pd.isna(value):
                        return ("none", "")
                except (TypeError, ValueError):
                    pass
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    try:
                        number = float(value)
                        return ("number", str(int(number)) if number.is_integer() else format(number, ".15g"))
                    except (TypeError, ValueError):
                        pass
                if isinstance(value, (date, datetime)):
                    return ("date", value.isoformat())
                text = str(value)
                return ("none", "") if text == "" else ("str", text)

            try:
                source_book = load_workbook(input_path, data_only=False, read_only=True)
                output_book = load_workbook(output_path, data_only=False, read_only=True)
            except (ET.ParseError, zipfile.BadZipFile, OSError) as exc:
                raise ValueError(
                    "导出安全校验无法读取 Excel 内部结构。文件可能含有非法 XML 字符或已损坏；"
                    "请用 Excel/WPS 另存为新的 .xlsx 后重新上传。原文件不会被程序覆盖。"
                ) from exc
            source_sheet = source_book[source_book.sheetnames[0]]
            output_sheet = output_book[output_book.sheetnames[0]]
            source_iter = source_sheet.iter_rows()
            output_iter = output_sheet.iter_rows()
            try:
                source_header_cells = next(source_iter)
                output_header_cells = next(output_iter)
            except StopIteration as exc:
                raise ValueError("导出安全校验失败：Excel 缺少表头") from exc
            source_headers = [cell.value for cell in source_header_cells]
            output_headers = [cell.value for cell in output_header_cells]
            source_preserved_headers = [header for header in source_headers if header not in target_fields and header not in generated_fields]
            output_preserved_headers = [header for header in output_headers if header not in target_fields and header not in generated_fields]
            if source_preserved_headers != output_preserved_headers:
                raise ValueError("导出安全校验失败：非目标列集合或顺序发生变化")
            if source_sheet.max_row != output_sheet.max_row:
                raise ValueError("导出安全校验失败：行数发生变化")
            preserved_columns = [column for column, header in enumerate(source_headers, start=1) if header not in target_fields and header not in generated_fields]
            for row_number, (left_row, right_row) in enumerate(zip(source_iter, output_iter), start=2):
                for column in preserved_columns:
                    left_cell = left_row[column - 1]
                    right_cell = right_row[column - 1]
                    left = semantic(left_cell.value, left_cell)
                    right = semantic(right_cell.value, right_cell)
                    if left != right:
                        header = source_headers[column - 1] or f"Unnamed: {column - 1}"
                        raise ValueError(
                            f"导出安全校验失败：{header}!{row_number} 被修改；"
                            f"原值={left[1]!r}，导出值={right[1]!r}"
                        )
            return

        def read(path: Path) -> pd.DataFrame:
            if path.suffix.lower() == ".csv":
                return pd.read_csv(path, dtype=object)
            return pd.read_excel(path, dtype=object)

        source = read(input_path)
        exported = read(output_path)
        if len(source) != len(exported):
            raise ValueError("导出安全校验失败：行数发生变化")
        source_non_target = [column for column in source.columns if column not in target_fields and column not in generated_fields]
        exported_non_target = [column for column in exported.columns if column not in target_fields and column not in generated_fields]
        if source_non_target != exported_non_target:
            raise ValueError("导出安全校验失败：非目标列集合或顺序发生变化")
        for column in source_non_target:
            for row_index, (left, right) in enumerate(zip(source[column].tolist(), exported[column].tolist()), start=2):
                if _clean(left) != _clean(right):
                    raise ValueError(f"导出安全校验失败：{column}!{row_index} 被修改")
