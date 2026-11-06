"""Codebook loading, stable identifiers, and M0 audit reports.

This module deliberately does not interpret ``notes`` into runtime matching
rules.  Notes are only surfaced for offline review, as required by the
project contract.
"""

from __future__ import annotations

import csv
import hashlib
import math
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


REQUIRED_COLUMNS = [
    "cityname",
    "distname",
    "dtype",
    "drank",
    "s1_sector",
    "wrank",
    "w_name",
    "alias",
    "notes",
    "year_valid_from",
    "month_valid_from",
    "year_valid_to",
    "month_valid_to",
]
# ``dtype`` and ``drank`` are allowed to be blank by the loading contract, but
# are still audited explicitly because they affect level comparisons.
AUDIT_REQUIRED = ["cityname", "w_name", "s1_sector", "wrank", "dtype", "drank"]
ALIAS_SPLIT_RE = re.compile(r"[;；,，、|｜\n\r]+")
# The runtime normalizer removes bracketed explanatory text while retaining it
# in ``bracket_content``.  Codebook IDs and alias keys must use the same rule;
# otherwise a formal name such as ``部门（挂牌）`` can never match the runtime
# query ``部门``.
BRACKET_CONTENT_RE = re.compile(r"（[^）]*）|\([^)]*\)|【[^】]*】|\[[^]]*\]")
NOTE_REVIEW_RE = re.compile(r"(?:\d{4}|\d{4}[年./-]\d{1,2}|撤销|合并|更名|挂牌)")
MULTI_CITY_RE = re.compile(r"\s*[|｜、,，;；]\s*")


class CodebookError(ValueError):
    """Raised when a codebook cannot satisfy the declared data contract."""


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def as_text(value: Any) -> str:
    return "" if _is_missing(value) else str(value).strip()


def normalize_name(value: Any) -> str:
    """Stable conservative name normalization for IDs and audit joins."""

    text = unicodedata.normalize("NFKC", as_text(value))
    text = BRACKET_CONTENT_RE.sub("", text)
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[，,。；;、:：（）()【】\[\]《》“”\"'·]", "", text)
    return text


def split_aliases(value: Any) -> list[str]:
    text = as_text(value)
    if not text:
        return []
    return [item.strip() for item in ALIAS_SPLIT_RE.split(text) if item.strip()]


def _int_value(value: Any) -> int | None:
    if _is_missing(value) or as_text(value) == "":
        return None
    if isinstance(value, bool):
        return int(value)
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or int(number) != number:
        return None
    return int(number)


def _hash10(*parts: Any) -> str:
    payload = "|".join("NULL" if part is None else str(part) for part in parts)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:10]


@dataclass(frozen=True)
class CodebookRow:
    source_row: int
    cityname: str
    distname: str
    dtype: str
    drank: str
    s1_sector: str
    wrank: str
    w_name: str
    alias: str
    notes: str
    year_valid_from: int | None
    month_valid_from: int | None
    year_valid_to: int | None
    month_valid_to: int | None
    norm_w_name: str
    entity_id: str
    dept_key: str
    time_invalid: bool = False

    @property
    def aliases(self) -> list[str]:
        return split_aliases(self.alias)

    @property
    def all_time_empty(self) -> bool:
        return not self.time_invalid and all(
            value is None
            for value in (
                self.year_valid_from,
                self.month_valid_from,
                self.year_valid_to,
                self.month_valid_to,
            )
        )

    @property
    def start_month(self) -> int:
        if self.year_valid_from is None:
            return -10**9
        return self.year_valid_from * 12 + ((self.month_valid_from or 1) - 1)

    @property
    def end_month(self) -> int:
        if self.year_valid_to is None:
            return 10**9
        return self.year_valid_to * 12 + ((self.month_valid_to or 12) - 1)


@dataclass
class CodebookAudit:
    missing_required: list[dict[str, Any]]
    invalid_months: list[dict[str, Any]]
    year_month_incomplete: list[dict[str, Any]]
    reversed_periods: list[dict[str, Any]]
    period_conflicts: list[dict[str, Any]]
    duplicate_city_names: list[dict[str, Any]]
    ambiguous_aliases: list[dict[str, Any]]
    alias_name_collisions: list[dict[str, Any]]
    notes_need_review: list[dict[str, Any]]
    multi_city_values: list[dict[str, Any]]
    duplicate_entity_ids: list[dict[str, Any]]

    @property
    def summary(self) -> dict[str, int]:
        return {
            "missing_required": len(self.missing_required),
            "invalid_months": len(self.invalid_months),
            "year_month_incomplete": len(self.year_month_incomplete),
            "reversed_periods": len(self.reversed_periods),
            "period_conflicts": len(self.period_conflicts),
            "duplicate_city_names": len(self.duplicate_city_names),
            "ambiguous_aliases": len(self.ambiguous_aliases),
            "alias_name_collisions": len(self.alias_name_collisions),
            "notes_need_review": len(self.notes_need_review),
            "multi_city_values": len(self.multi_city_values),
            "duplicate_entity_ids": len(self.duplicate_entity_ids),
        }


class Codebook:
    """Load a contract-compliant codebook and expose stable row identities."""

    def __init__(self, path: str | Path, reports_dir: str | Path = "reports"):
        self.path = Path(path)
        self.reports_dir = Path(reports_dir)
        if not self.path.exists():
            raise CodebookError(f"找不到 codebook 文件: {self.path}")
        self.frame = self._read_frame()
        self.rows = self._build_rows(self.frame)
        self.audit_result = self._audit()

    def _read_frame(self) -> pd.DataFrame:
        suffix = self.path.suffix.lower()
        if suffix == ".csv":
            frame = pd.read_csv(self.path, dtype=object)
        elif suffix in {".xlsx", ".xlsm", ".xls"}:
            frame = pd.read_excel(self.path, dtype=object)
        else:
            raise CodebookError("codebook 仅支持 CSV 或 Excel 文件")
        missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
        if missing:
            raise CodebookError(f"codebook 缺少必要列: {missing}")
        return frame[REQUIRED_COLUMNS].copy()

    @staticmethod
    def _build_rows(frame: pd.DataFrame) -> list[CodebookRow]:
        rows: list[CodebookRow] = []
        for index, values in frame.iterrows():
            data = {column: values[column] for column in REQUIRED_COLUMNS}
            cityname = as_text(data["cityname"])
            norm_w_name = normalize_name(data["w_name"])
            year_from = _int_value(data["year_valid_from"])
            month_from = _int_value(data["month_valid_from"])
            year_to = _int_value(data["year_valid_to"])
            month_to = _int_value(data["month_valid_to"])
            raw_from_year = data["year_valid_from"]
            raw_from_month = data["month_valid_from"]
            raw_to_year = data["year_valid_to"]
            raw_to_month = data["month_valid_to"]
            time_invalid = (
                (not _is_missing(raw_from_month) and as_text(raw_from_month) != "" and month_from is None)
                or (not _is_missing(raw_to_month) and as_text(raw_to_month) != "" and month_to is None)
                or (year_from is None and month_from is not None)
                or (year_to is None and month_to is not None)
            )
            if not time_invalid:
                start_value = year_from * 12 + ((month_from or 1) - 1) if year_from is not None else -10**9
                end_value = year_to * 12 + ((month_to or 12) - 1) if year_to is not None else 10**9
                time_invalid = start_value > end_value
            rows.append(
                CodebookRow(
                    source_row=int(index) + 2,
                    cityname=cityname,
                    distname=as_text(data["distname"]),
                    dtype=as_text(data["dtype"]),
                    drank=as_text(data["drank"]),
                    s1_sector=as_text(data["s1_sector"]),
                    wrank=as_text(data["wrank"]),
                    w_name=as_text(data["w_name"]),
                    alias=as_text(data["alias"]),
                    notes=as_text(data["notes"]),
                    year_valid_from=year_from,
                    month_valid_from=month_from,
                    year_valid_to=year_to,
                    month_valid_to=month_to,
                    norm_w_name=norm_w_name,
                    entity_id=_hash10(cityname, norm_w_name, year_from, month_from),
                    dept_key=_hash10(cityname, norm_w_name),
                    time_invalid=time_invalid,
                )
            )
        return rows

    @staticmethod
    def _row_dict(row: CodebookRow, **extra: Any) -> dict[str, Any]:
        result = {
            "source_row": row.source_row,
            "cityname": row.cityname,
            "distname": row.distname,
            "w_name": row.w_name,
            "norm_w_name": row.norm_w_name,
            "entity_id": row.entity_id,
            "dept_key": row.dept_key,
            "year_valid_from": row.year_valid_from,
            "month_valid_from": row.month_valid_from,
            "year_valid_to": row.year_valid_to,
            "month_valid_to": row.month_valid_to,
            "time_invalid": row.time_invalid,
            "s1_sector": row.s1_sector,
            "wrank": row.wrank,
            "alias": row.alias,
            "notes": row.notes,
        }
        result.update(extra)
        return result

    def _audit(self) -> CodebookAudit:
        missing: list[dict[str, Any]] = []
        invalid_months: list[dict[str, Any]] = []
        year_month_incomplete: list[dict[str, Any]] = []
        reversed_periods: list[dict[str, Any]] = []
        duplicate_city_names: list[dict[str, Any]] = []
        notes_need_review: list[dict[str, Any]] = []
        multi_city_values: list[dict[str, Any]] = []

        for row in self.rows:
            for column, value in (
                ("cityname", row.cityname),
                ("w_name", row.w_name),
                ("s1_sector", row.s1_sector),
                ("wrank", row.wrank),
                ("dtype", row.dtype),
                ("drank", row.drank),
            ):
                if not value:
                    missing.append({"source_row": row.source_row, "column": column})
            for column, value in (
                ("month_valid_from", row.month_valid_from),
                ("month_valid_to", row.month_valid_to),
            ):
                raw = self.frame.iloc[row.source_row - 2][column]
                if not _is_missing(raw) and as_text(raw) and (value is None or not 1 <= value <= 12):
                    invalid_months.append(self._row_dict(row, column=column, raw_value=as_text(raw)))
            if row.year_valid_from is None and row.month_valid_from is not None:
                year_month_incomplete.append(self._row_dict(row, issue="from_month_without_year"))
            if row.year_valid_to is None and row.month_valid_to is not None:
                year_month_incomplete.append(self._row_dict(row, issue="to_month_without_year"))
            if not row.all_time_empty and row.start_month > row.end_month:
                reversed_periods.append(self._row_dict(row))
            if MULTI_CITY_RE.search(row.cityname):
                multi_city_values.append(self._row_dict(row, issue="multiple_city_values"))
            if NOTE_REVIEW_RE.search(row.notes):
                notes_need_review.append(self._row_dict(row, note_match=row.notes))

        groups: dict[tuple[str, str], list[CodebookRow]] = {}
        by_dept_key: dict[str, list[CodebookRow]] = {}
        for row in self.rows:
            for city_value in self.split_values(row.cityname):
                groups.setdefault((normalize_name(city_value), row.norm_w_name), []).append(row)
            by_dept_key.setdefault(row.dept_key, []).append(row)

        for (city, name), rows in groups.items():
            if city and name and len(rows) > 1:
                duplicate_city_names.extend(
                    self._row_dict(row, duplicate_group_size=len(rows)) for row in rows
                )

        period_conflicts: list[dict[str, Any]] = []
        for dept_key, rows in by_dept_key.items():
            # Rows missing city/name are already captured by the required-field
            # audit. They cannot provide a meaningful period conflict signal.
            valid_rows = [row for row in rows if row.cityname and row.norm_w_name]
            modes = {row.all_time_empty for row in valid_rows}
            if len(modes) > 1:
                for row in valid_rows:
                    period_conflicts.append(self._row_dict(row, issue="empty_nonempty_time_mix"))
            for left_index, left in enumerate(valid_rows):
                if any(
                    value is None
                    for value in (
                        left.year_valid_from,
                        left.month_valid_from,
                        left.year_valid_to,
                        left.month_valid_to,
                    )
                ) and not left.all_time_empty:
                    pass
                for right in valid_rows[left_index + 1 :]:
                    if left.start_month <= right.end_month and right.start_month <= left.end_month:
                        period_conflicts.extend(
                            [
                                self._row_dict(left, issue="time_overlap", other_source_row=right.source_row),
                                self._row_dict(right, issue="time_overlap", other_source_row=left.source_row),
                            ]
                        )

        alias_map: dict[tuple[str, str], set[str]] = {}
        name_map: dict[tuple[str, str], set[str]] = {}
        alias_entity_ids: dict[tuple[str, str], set[str]] = {}
        for row in self.rows:
            for city_value in self.split_values(row.cityname):
                city = normalize_name(city_value)
                name_map.setdefault((city, row.norm_w_name), set()).add(row.dept_key)
                for alias in row.aliases:
                    key = (city, normalize_name(alias))
                    alias_map.setdefault(key, set()).add(row.dept_key)
                    alias_entity_ids.setdefault(key, set()).add(row.entity_id)

        ambiguous_aliases: list[dict[str, Any]] = []
        for (city, alias), dept_keys in sorted(alias_map.items()):
            if alias and len(dept_keys) > 1:
                ambiguous_aliases.append(
                    {
                        "cityname_norm": city,
                        "alias_norm": alias,
                        "dept_keys": ",".join(sorted(dept_keys)),
                        "entity_ids": ",".join(sorted(alias_entity_ids.get((city, alias), set()))),
                    }
                )

        alias_name_collisions: list[dict[str, Any]] = []
        for (city, alias), dept_keys in sorted(alias_map.items()):
            name_keys = name_map.get((city, alias), set()) - dept_keys
            if alias and name_keys:
                alias_name_collisions.append(
                    {
                        "cityname_norm": city,
                        "alias_norm": alias,
                        "alias_dept_keys": ",".join(sorted(dept_keys)),
                        "name_dept_keys": ",".join(sorted(name_keys)),
                    }
                )

        duplicate_entity_ids: list[dict[str, Any]] = []
        by_entity_id: dict[str, list[CodebookRow]] = {}
        for row in self.rows:
            by_entity_id.setdefault(row.entity_id, []).append(row)
        for entity_id, rows in by_entity_id.items():
            if len(rows) > 1:
                for row in rows:
                    duplicate_entity_ids.append(
                        self._row_dict(row, duplicate_group_size=len(rows))
                    )

        return CodebookAudit(
            missing_required=missing,
            invalid_months=invalid_months,
            year_month_incomplete=year_month_incomplete,
            reversed_periods=reversed_periods,
            period_conflicts=period_conflicts,
            duplicate_city_names=duplicate_city_names,
            ambiguous_aliases=ambiguous_aliases,
            alias_name_collisions=alias_name_collisions,
            notes_need_review=notes_need_review,
            multi_city_values=multi_city_values,
            duplicate_entity_ids=duplicate_entity_ids,
        )

    def write_reports(self) -> CodebookAudit:
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        audit = self.audit_result
        self._write_csv(self.reports_dir / "period_conflicts.csv", audit.period_conflicts)
        self._write_csv(self.reports_dir / "notes_need_review.csv", audit.notes_need_review)
        report = self.reports_dir / "codebook_audit.md"
        lines = [
            "# Codebook audit",
            "",
            f"- source: `{self.path}`",
            f"- rows loaded: {len(self.rows)}",
            f"- unique entity IDs: {len({row.entity_id for row in self.rows})}",
            f"- unique department keys: {len({row.dept_key for row in self.rows})}",
            "",
            "## Summary",
            "",
            "| audit | count |",
            "|---|---:|",
        ]
        lines.extend(f"| {key} | {value} |" for key, value in audit.summary.items())
        lines.extend(
            [
                "",
                "## Time semantics",
                "",
                "Rows with all four validity fields empty apply to all dates. Missing start/end years are treated as negative/positive infinity; a missing start month becomes January and a missing end month becomes December when the corresponding year exists.",
                "",
                "## Findings",
                "",
                "- `period_conflicts.csv` contains overlap and empty/non-empty time mixing signals.",
                "- `notes_need_review.csv` contains notes with years or change-status keywords. Notes are not parsed into runtime rules.",
                "- `multi_city_values` identifies source cells such as `广州市 | 深圳市`; M0 preserves them as source values and defers city expansion to a later milestone.",
            ]
        )
        report.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return audit

    @staticmethod
    def _write_csv(path: Path, rows: Iterable[dict[str, Any]]) -> None:
        materialized = list(rows)
        if not materialized:
            # Keep a readable, machine-loadable header even when an audit has
            # no findings.  Downstream analysts can then use one schema for
            # all runs.
            if path.name == "period_conflicts.csv":
                path.write_text("source_row,issue\n", encoding="utf-8-sig")
            elif path.name == "notes_need_review.csv":
                path.write_text("source_row,note_match\n", encoding="utf-8-sig")
            else:
                path.write_text("\n", encoding="utf-8-sig")
            return
        fields: list[str] = []
        for row in materialized:
            for field in row:
                if field not in fields:
                    fields.append(field)
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(materialized)

    def enriched_frame(self) -> pd.DataFrame:
        """Return source rows plus stable M0 helper columns for later modules."""

        rows = [
            {
                **{column: getattr(row, column) for column in REQUIRED_COLUMNS},
                "norm_w_name": row.norm_w_name,
                "entity_id": row.entity_id,
                "dept_key": row.dept_key,
            }
            for row in self.rows
        ]
        return pd.DataFrame(rows)

    def rows_for_entity(self, entity_id: str) -> list[CodebookRow]:
        return [row for row in self.rows if row.entity_id == entity_id]

    def row_for_entity(self, entity_id: str) -> CodebookRow | None:
        rows = self.rows_for_entity(entity_id)
        return rows[0] if len(rows) == 1 else None

    def valid_rows(self) -> list[CodebookRow]:
        return [row for row in self.rows if row.cityname and row.norm_w_name and row.w_name]

    @staticmethod
    def split_values(value: str) -> tuple[str, ...]:
        """Split source cells such as ``广州市 | 深圳市`` without rewriting them."""

        return tuple(part.strip() for part in re.split(r"\s*[|｜、,，;；]\s*", value or "") if part.strip())

    @classmethod
    def row_applies_to_city(cls, row: CodebookRow, cityname: str) -> bool:
        if not row.cityname or not cityname:
            return False
        wanted = normalize_name(cityname)
        return wanted in {normalize_name(value) for value in cls.split_values(row.cityname)}
