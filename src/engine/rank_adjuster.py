"""Rank normalization and post_rank adjustment utilities.

The source data uses a few historical aliases (for example ``地厅级`` for
``正厅级``).  All calculations happen on a single ordered scale and the
original text is never discarded.  Unknown labels are returned as explicit
exceptions so they can be reviewed rather than silently changed.
"""

from __future__ import annotations

from dataclasses import dataclass


# Lowest to highest.  The two股级 values are retained because a downward
# department-level change can legitimately move a post below副科级.
RANK_LEVELS = (
    "副股级", "股级", "副科级", "正科级", "副处级", "正处级",
    "副厅级", "正厅级", "副部级", "正部级",
)
RANK_SCORE = {name: index for index, name in enumerate(RANK_LEVELS)}
RANK_ALIASES = {
    "科级": "正科级",
    "处级": "正处级",
    "厅级": "正厅级",
    "地厅级": "正厅级",
    "部级": "正部级",
    "省部级": "正部级",
}
# Preferred spelling when writing back to the historical Guangzhou/Shenzhen
# files.  The canonical names remain available internally for comparisons.
DISPLAY_RANK = {
    "正科级": "科级",
    "正处级": "处级",
    "正厅级": "地厅级",
    "正部级": "省部级",
}


@dataclass(frozen=True)
class RankAdjustment:
    original: str
    adjusted: str
    delta: int
    reason: str
    status: str = "adjusted"


def canonical_rank(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return RANK_ALIASES.get(text, text)


def rank_score(value: object) -> int | None:
    return RANK_SCORE.get(canonical_rank(value))


def adjust_post_rank(post_rank: object, old_wrank: object, new_wrank: object) -> RankAdjustment:
    original = "" if post_rank is None else str(post_rank).strip()
    old = canonical_rank(old_wrank)
    new = canonical_rank(new_wrank)
    post = canonical_rank(original)
    old_score = RANK_SCORE.get(old)
    new_score = RANK_SCORE.get(new)
    post_score = RANK_SCORE.get(post)
    if not original or old_score is None or new_score is None:
        return RankAdjustment(original, original, 0, f"无法识别部门级别：{old_wrank} → {new_wrank}", "unresolved")
    delta = new_score - old_score
    if post_score is None:
        return RankAdjustment(original, original, delta, f"无法识别原 post_rank：{post_rank}", "unresolved")
    adjusted_score = post_score + delta
    if adjusted_score < 0 or adjusted_score >= len(RANK_LEVELS):
        return RankAdjustment(original, original, delta, f"调整后超出级别范围：{post_rank}，变化量 {delta:+d}", "unresolved")
    adjusted = DISPLAY_RANK.get(RANK_LEVELS[adjusted_score], RANK_LEVELS[adjusted_score])
    if adjusted == original:
        reason = "部门级别未发生可计算变化"
    else:
        reason = f"部门级别 {old_wrank} → {new_wrank}，post_rank 变化 {delta:+d}"
    return RankAdjustment(original, adjusted, delta, reason)

