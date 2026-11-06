"""Pairwise and within-query features for ranking."""

from __future__ import annotations

import math
import re
from collections import Counter
from pathlib import Path
from typing import Iterable

import numpy as np
from rapidfuzz import fuzz
from rapidfuzz.distance import JaroWinkler, LCSseq

from .candidates import Candidate
from .normalize import NormalizedName, normalize_name as runtime_normalize


FEATURE_NAMES = [
    "ratio", "partial_ratio", "token_sort_ratio", "token_set_ratio",
    "jaro_winkler", "lcs_ratio", "dice_1g", "dice_2g", "tfidf_cos", "len_ratio",
    "exact_norm", "exact_core", "exact_alias", "suffix_match",
    "suffix_same", "suffix_query_missing", "suffix_candidate_missing",
    "key_term_conflict", "digit_mismatch", "modifier_flag", "alias_ambiguous",
    "out_of_period", "period_unknown", "period_ambiguous", "codebook_period_invalid", "period_match",
    "input_period_invalid", "period_all_time", "period_overlap", "period_no_overlap",
    "city_missing", "city_inferred", "dist_mismatch", "dist_match", "dist_query_missing", "dist_candidate_missing",
    "group_ratio_rank", "group_tfidf_rank", "gap_to_best", "gap_to_second",
    "candidate_count",
]


def _load_conflicts() -> tuple[tuple[str, str], ...]:
    path = Path(__file__).resolve().parents[2] / "resources" / "conflict_terms.txt"
    if not path.exists():
        return (("公安", "国安"), ("党委", "政府"), ("发展改革", "改革"), ("纪委", "监察"))
    pairs: list[tuple[str, str]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        values = [item.strip() for item in line.split("|") if item.strip()]
        if len(values) == 2:
            pairs.append((values[0], values[1]))
    return tuple(pairs)


_CONFLICT_TERMS = _load_conflicts()


def _dice(a: str, b: str, n: int) -> float:
    if not a or not b:
        return 0.0
    grams_a = {a[i:i + n] for i in range(max(1, len(a) - n + 1))}
    grams_b = {b[i:i + n] for i in range(max(1, len(b) - n + 1))}
    return 2 * len(grams_a & grams_b) / max(1, len(grams_a) + len(grams_b))


def _digits(text: str) -> set[str]:
    return set(re.findall(r"\d+|[一二三四五六七八九十百]+", text))


def _conflict(a: str, b: str) -> int:
    for left, right in _CONFLICT_TERMS:
        if (left in a and right in b) or (right in a and left in b):
            return 1
    return 0


def _suffix(name: str) -> str:
    for suffix in ("委员会", "办公室", "管理局", "局", "委", "厅", "办", "中心", "院", "会"):
        if name.endswith(suffix):
            return suffix
    return ""


def _segment(value: str) -> str:
    try:
        import jieba  # type: ignore

        return " ".join(token for token in jieba.lcut(value) if token.strip())
    except Exception:
        return " ".join(value)


def _char_cosine(a: str, b: str, ngram: int = 2) -> float:
    def counts(value: str) -> Counter[str]:
        if not value:
            return Counter()
        if len(value) < ngram:
            return Counter({value: 1})
        return Counter(value[index : index + ngram] for index in range(len(value) - ngram + 1))

    left, right = counts(a), counts(b)
    if not left or not right:
        return 0.0
    dot = sum(left[key] * right.get(key, 0) for key in left)
    norm_left = math.sqrt(sum(value * value for value in left.values()))
    norm_right = math.sqrt(sum(value * value for value in right.values()))
    return dot / max(1e-12, norm_left * norm_right)


def pair_features(query: NormalizedName, candidate: Candidate) -> dict[str, float]:
    name = query.norm
    # Retrieval keeps the source spelling for evidence, but all similarity
    # features must compare the same normalized representation as the query.
    variant = runtime_normalize(candidate.matched_variant).norm or candidate.matched_variant
    core = query.core
    candidate_core = candidate.matched_core
    query_suffix = query.suffix_type or _suffix(name)
    candidate_suffix = _suffix(variant)
    segmented_name = _segment(name)
    segmented_variant = _segment(variant)
    ratio = fuzz.ratio(name, variant) / 100.0
    partial = fuzz.partial_ratio(name, variant) / 100.0
    token_sort = fuzz.token_sort_ratio(segmented_name, segmented_variant) / 100.0
    token_set = fuzz.token_set_ratio(segmented_name, segmented_variant) / 100.0
    flags = set(candidate.flags)
    overlap = 0.0
    for flag in flags:
        if flag.startswith("period_overlap_months="):
            overlap = float(flag.split("=", 1)[1])
    return {
        "ratio": ratio,
        "partial_ratio": partial,
        "token_sort_ratio": token_sort,
        "token_set_ratio": token_set,
        "jaro_winkler": JaroWinkler.normalized_similarity(name, variant),
        "lcs_ratio": LCSseq.normalized_similarity(name, variant),
        "dice_1g": _dice(name, variant, 1),
        "dice_2g": _dice(name, variant, 2),
        "tfidf_cos": _char_cosine(name, variant, 2),
        "len_ratio": min(len(name), len(variant)) / max(1, max(len(name), len(variant))),
        "exact_norm": float(name == variant and bool(name)),
        "exact_core": float(core == candidate_core and bool(core)),
        "exact_alias": float(candidate.matched_field == "alias" and name == variant),
        "suffix_match": float(query_suffix == candidate_suffix and bool(query_suffix)),
        "suffix_same": float(query_suffix == candidate_suffix and bool(query_suffix)),
        "suffix_query_missing": float(not query_suffix),
        "suffix_candidate_missing": float(not candidate_suffix),
        "key_term_conflict": float(_conflict(name, variant)),
        "digit_mismatch": float(bool(_digits(name) ^ _digits(variant))),
        "modifier_flag": float(bool(query.modifiers)),
        "alias_ambiguous": float("alias_ambiguous" in flags),
        "out_of_period": float("out_of_period" in flags),
        "period_unknown": float("period_unknown" in flags),
        "period_ambiguous": float("period_ambiguous" in flags),
        "codebook_period_invalid": float("codebook_period_invalid" in flags),
        "period_match": float("out_of_period" not in flags and "period_unknown" not in flags and "input_period_invalid" not in flags),
        "input_period_invalid": float("input_period_invalid" in flags),
        "period_all_time": float("period_overlap_months" not in flags and "out_of_period" not in flags and "period_unknown" not in flags),
        "period_overlap": float("period_overlap_months" in flags and "out_of_period" not in flags),
        "period_no_overlap": float("out_of_period" in flags),
        "period_overlap_months": overlap,
        "city_missing": float("city_missing" in flags),
        "city_inferred": float("city_inferred" in flags),
        "dist_mismatch": float("dist_mismatch" in flags),
        "dist_match": float("dist_match" in flags),
        "dist_query_missing": float("dist_query_missing" in flags),
        "dist_candidate_missing": float(not candidate.row.distname),
    }


def build_feature_rows(query: NormalizedName, candidates: list[Candidate]) -> list[dict[str, float | str]]:
    rows = [pair_features(query, candidate) for candidate in candidates]
    ratios = sorted((float(row["ratio"]) for row in rows), reverse=True)
    tfidf = sorted((float(candidate.retrieval_score) for candidate in candidates), reverse=True)
    best_ratio = ratios[0] if ratios else 0.0
    second_ratio = ratios[1] if len(ratios) > 1 else 0.0
    best_tfidf = tfidf[0] if tfidf else 0.0
    second_tfidf = tfidf[1] if len(tfidf) > 1 else 0.0
    for i, row in enumerate(rows):
        row["group_ratio_rank"] = float(1 + sum(value > float(row["ratio"]) for value in ratios)) if rows else 0.0
        score = float(candidates[i].retrieval_score)
        row["group_tfidf_rank"] = float(1 + sum(value > score for value in tfidf)) if rows else 0.0
        row["gap_to_best"] = best_ratio - float(row["ratio"])
        # This is the candidate's distance from the second-best score.  It is
        # positive when the candidate is below the runner-up and negative for
        # the top candidate, which preserves the direction explicitly.
        row["gap_to_second"] = second_ratio - float(row["ratio"]) if len(ratios) > 1 else 0.0
        row["best_field"] = candidates[i].matched_field
        row["candidate_count"] = float(len(candidates))
        row["entity_id"] = candidates[i].entity_id
    return rows


def feature_matrix(rows: list[dict[str, float | str]], feature_names: list[str] | None = None) -> np.ndarray:
    names = feature_names or FEATURE_NAMES
    return np.asarray([[float(row.get(name, 0.0)) for name in names] for row in rows], dtype=float)
