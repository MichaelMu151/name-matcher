"""Candidate generation, input periods, and time-aware flags."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .codebook import Codebook, CodebookRow, normalize_name
from .index import IndexEntry, NameIndex
from .normalize import NormalizedName


def _is_broad_party_committee(normalized: NormalizedName) -> bool:
    """Recognize a generic local party committee by institutional form."""
    text = normalized.raw.replace(" ", "")
    excluded = ("纪律检查", "监察", "政法", "宣传", "组织", "统战", "农村", "直属机关", "党校")
    return "中共" in text and text.endswith("委员会") and not any(term in text for term in excluded)


@dataclass(frozen=True)
class InputPeriod:
    start: int | None
    end: int | None
    invalid: bool = False
    unknown: bool = False


@dataclass
class Candidate:
    entity_id: str
    dept_key: str
    row: CodebookRow
    matched_variant: str
    matched_field: str
    matched_core: str
    retrieval_path: str
    retrieval_score: float
    flags: list[str]


def _month(year: Any, month: Any, default: int) -> int | None:
    if year is None or str(year).strip() == "":
        return None
    try:
        year_number = float(year)
        if not year_number.is_integer():
            return None
        year_int = int(year_number)
    except (TypeError, ValueError, OverflowError):
        return None
    if month is None or str(month).strip() == "":
        month_int = default
    else:
        try:
            month_int = int(float(month))
        except (TypeError, ValueError, OverflowError):
            return None
    if not 1 <= month_int <= 12:
        return None
    return year_int * 12 + month_int - 1


def input_period(record: dict[str, Any]) -> InputPeriod:
    styr, stmo = record.get("styr"), record.get("stmo")
    fiyr, fimo = record.get("fiyr"), record.get("fimo")
    def present(value: Any) -> bool:
        if value is None:
            return False
        try:
            if value != value:  # NaN
                return False
        except Exception:
            pass
        return str(value).strip() != ""

    if not present(fiyr) and present(fimo):
        return InputPeriod(None, None, invalid=True)
    start = _month(styr, stmo, 1)
    end = _month(fiyr, fimo, 12)
    if (present(styr) and start is None) or (present(fiyr) and end is None):
        return InputPeriod(start, end, invalid=True)
    if start is not None and end is not None and start > end:
        return InputPeriod(start, end, invalid=True)
    unknown = start is None or end is None
    return InputPeriod(start, end, unknown=unknown)


def _overlap(period: InputPeriod, row: CodebookRow, month_tolerance: int = 0) -> tuple[bool | None, int | None]:
    if row.time_invalid:
        return None, None
    if row.all_time_empty:
        return True, None
    if period.invalid or period.unknown:
        return None, None
    tolerance = max(0, int(month_tolerance))
    start = max(period.start, row.start_month - tolerance)
    end = min(period.end, row.end_month + tolerance)
    return start <= end, max(0, end - start + 1)


def generate_candidates(
    record: dict[str, Any],
    index: NameIndex,
    normalized: NormalizedName,
    top_k: int = 10,
    distname_hard_filter: bool = False,
    month_tolerance: int = 0,
    year_tolerance: int = 0,
) -> tuple[list[Candidate], list[str]]:
    supplied_city = str(record.get("cityname") or "").strip()
    cityname = supplied_city or str(normalized.city_hint or "")
    flags = list(normalized.flags)
    period = input_period(record)
    if period.invalid:
        flags.append("input_period_invalid")
    if period.unknown:
        flags.append("period_unknown")
    if supplied_city == "" and normalized.city_hint:
        flags.append("city_inferred")
    elif not cityname:
        flags.append("city_missing")
    raw_candidates = index.retrieve(cityname, normalized, top_k=top_k)
    candidates: list[Candidate] = []
    seen: set[str] = set()
    for entry, score, path in raw_candidates:
        row = entry.row
        if row.entity_id in seen:
            continue
        candidate_flags = list(flags)
        if row.time_invalid:
            candidate_flags.append("codebook_period_invalid")
        applies, overlap_months = _overlap(
            period,
            row,
            month_tolerance=max(0, int(month_tolerance)) + 12 * max(0, int(year_tolerance)),
        )
        if applies is False:
            candidate_flags.append("out_of_period")
        elif applies is None and not row.all_time_empty:
            candidate_flags.append("period_unknown")
        if overlap_months is not None:
            candidate_flags.append(f"period_overlap_months={overlap_months}")
        query_dist = str(record.get("distname") or "").strip()
        if not query_dist:
            candidate_flags.append("dist_query_missing")
        if not row.distname:
            candidate_flags.append("dist_candidate_missing")
        dist_matches = query_dist and row.distname and any(
            normalize_name(query_dist) == normalize_name(value)
            for value in Codebook.split_values(row.distname)
        )
        if query_dist and row.distname and not dist_matches:
            candidate_flags.append("dist_mismatch")
            if distname_hard_filter:
                continue
        elif query_dist and row.distname and dist_matches:
            candidate_flags.append("dist_match")
        if entry.variant_type == "alias" and index.alias_is_ambiguous(entry, cityname):
            candidate_flags.append("alias_ambiguous")
        candidates.append(Candidate(row.entity_id, row.dept_key, row, entry.variant, entry.variant_type, entry.core, path, score, candidate_flags))
        seen.add(row.entity_id)
    # Fuzzy character similarity cannot reliably recover the semantic class
    # of names such as “中共广州市第七届委员会”. Add the existing generic
    # party-committee standard as a conservative human-review candidate.
    if _is_broad_party_committee(normalized):
        for row in index.codebook.valid_rows():
            if row.w_name != "地方党委" or row.entity_id in seen:
                continue
            if cityname and not index.codebook.row_applies_to_city(row, cityname):
                continue
            candidate_flags = list(flags) + ["policy_broad_party_committee"]
            applies, overlap_months = _overlap(
                period, row,
                month_tolerance=max(0, int(month_tolerance)) + 12 * max(0, int(year_tolerance)),
            )
            if applies is False:
                candidate_flags.append("out_of_period")
            elif applies is None and not row.all_time_empty:
                candidate_flags.append("period_unknown")
            if overlap_months is not None:
                candidate_flags.append(f"period_overlap_months={overlap_months}")
            query_dist = str(record.get("distname") or "").strip()
            if query_dist and row.distname:
                dist_matches = any(normalize_name(query_dist) == normalize_name(value) for value in Codebook.split_values(row.distname))
                candidate_flags.append("dist_match" if dist_matches else "dist_mismatch")
                if not dist_matches and distname_hard_filter:
                    continue
            candidates.append(Candidate(row.entity_id, row.dept_key, row, row.w_name, "w_name", row.norm_w_name, "policy", 0.99, candidate_flags))
            seen.add(row.entity_id)
            break
    dept_hits: dict[str, int] = {}
    for candidate in candidates:
        if "out_of_period" not in candidate.flags:
            dept_hits[candidate.dept_key] = dept_hits.get(candidate.dept_key, 0) + 1
    if any(count > 1 for count in dept_hits.values()):
        for candidate in candidates:
            if dept_hits.get(candidate.dept_key, 0) > 1:
                candidate.flags.append("period_ambiguous")
    return candidates, sorted(set(flags))
