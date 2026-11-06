"""Decision routing policy for automatic, LLM, and human review."""

from __future__ import annotations

from dataclasses import dataclass, field

from .candidates import Candidate


@dataclass
class Decision:
    decision: str
    entity_id: str | None
    method: str
    p_calibrated: float
    margin: float
    reason: str
    flags: list[str] = field(default_factory=list)


def decide(candidates: list[Candidate], probabilities: list[float], query_flags: list[str], tau_auto: float = 0.85, margin_auto: float = 0.10, tau_nil: float = 0.15) -> Decision:
    flags = sorted(set(query_flags + [flag for candidate in candidates for flag in candidate.flags]))
    if "multi_entity" in flags:
        return Decision("MULTIPLE", None, "policy", 0.0, 0.0, "输入名称包含多个实体", flags)
    if "input_period_invalid" in flags:
        return Decision("HUMAN", None, "policy", 0.0, 0.0, "输入任职时间组合非法", flags)
    if not candidates:
        return Decision("NONE", None, "policy", 0.0, 0.0, "没有候选", flags)
    if len(probabilities) != len(candidates):
        return Decision("HUMAN", None, "policy", 0.0, 0.0, "候选与模型分数数量不一致", flags)
    probabilities = [min(1.0, max(0.0, float(value))) for value in probabilities]
    order = sorted(range(len(candidates)), key=lambda index: probabilities[index], reverse=True)
    best = order[0]
    second = probabilities[order[1]] if len(order) > 1 else 0.0
    best_probability = float(probabilities[best])
    margin = best_probability - second
    blocked = {"out_of_period", "period_ambiguous", "period_unknown", "dist_mismatch", "city_missing", "codebook_period_invalid", "input_period_invalid"}
    exact_hits = [
        candidate
        for candidate in candidates
        if candidate.retrieval_path in {"exact", "alias"}
        and not (blocked & set(candidate.flags))
    ]
    exact_entities = {candidate.entity_id for candidate in exact_hits}
    clean_exact = (
        candidates[best].retrieval_path in {"exact", "alias"}
        and not (blocked & set(candidates[best].flags))
        and "alias_ambiguous" not in candidates[best].flags
        and len(exact_entities) == 1
    )
    if clean_exact:
        return Decision("AUTO", candidates[best].entity_id, candidates[best].retrieval_path, 1.0, margin, "唯一精确或别名命中", flags)
    if best_probability >= tau_auto and margin >= margin_auto and not (blocked & set(flags)):
        return Decision("AUTO", candidates[best].entity_id, "ranker", best_probability, margin, "排序模型高置信且间隔足够", flags)
    if best_probability <= tau_nil:
        return Decision("LLM", None, "policy", best_probability, margin, "最高候选概率低于 NIL 阈值", flags)
    return Decision("LLM", candidates[best].entity_id, "policy", best_probability, margin, "需要 LLM 或人工仲裁", flags)
