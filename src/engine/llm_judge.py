"""LLM arbitration interface with deterministic FakeJudge and safe validation."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from .schemas import LLMResponse


ALLOWED_DECISIONS = {"NONE", "MULTIPLE", "NOT_A_DEPARTMENT"}


@dataclass(frozen=True)
class Verdict:
    decision: str
    reason: str
    rule_refs: tuple[str, ...] = ()
    uncertain: bool = False


class JudgeError(ValueError):
    pass


class FakeJudge:
    def __init__(self, default: str = "NONE"):
        self.default = default

    def __call__(self, query: dict[str, Any], candidates: list[dict[str, Any]], rules: list[str], examples: list[dict[str, Any]]) -> dict[str, Any]:
        return {"decision": self.default, "reason": "FakeJudge 默认结果", "rule_refs": [], "uncertain": True}


class LLMJudge:
    def __init__(self, model: str = "fake", judge_fn: Callable[..., dict[str, Any]] | None = None, cache_path: str | Path = "reports/llm_cache.jsonl", budget: int = 100):
        self.model = model
        self.judge_fn = judge_fn or FakeJudge()
        self.cache_path = Path(cache_path)
        self.budget = budget
        self.calls = 0
        self.cache: dict[str, dict[str, Any]] = {}
        if self.cache_path.exists():
            for line in self.cache_path.read_text(encoding="utf-8").splitlines():
                try:
                    entry = json.loads(line)
                    self.cache[entry["key"]] = entry["result"]
                except (json.JSONDecodeError, KeyError):
                    continue

    def _key(self, query: dict[str, Any], candidates: list[dict[str, Any]], rules: list[str], examples: list[dict[str, Any]], order: list[int]) -> str:
        payload = json.dumps({"model": self.model, "query": query, "candidates": [candidates[index] for index in order], "rules": rules, "examples": examples}, ensure_ascii=False, sort_keys=True)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    @staticmethod
    def validate(raw: dict[str, Any], candidate_ids: set[str]) -> Verdict:
        if not isinstance(raw, dict):
            raise JudgeError("LLM 输出不是 JSON 对象")
        try:
            parsed = LLMResponse.model_validate(raw)
        except Exception as exc:
            raise JudgeError(f"LLM JSON schema 校验失败: {exc}") from exc
        decision = parsed.decision
        if decision not in candidate_ids and decision not in ALLOWED_DECISIONS:
            raise JudgeError("LLM decision 不在候选 entity_id 或固定选项中")
        return Verdict(decision, parsed.reason, tuple(parsed.rule_refs), parsed.uncertain)

    def judge(self, query: dict[str, Any], candidates: list[dict[str, Any]], rules: list[str] | None = None, examples: list[dict[str, Any]] | None = None) -> Verdict:
        rules = rules or []
        examples = examples or []
        candidate_ids = {str(candidate["entity_id"]) for candidate in candidates}
        orders = [list(range(len(candidates))), list(reversed(range(len(candidates))))]
        verdicts: list[Verdict] = []
        for order in orders:
            key = self._key(query, candidates, rules, examples, order)
            if key in self.cache:
                raw = self.cache[key]
            else:
                if self.calls >= self.budget:
                    raise JudgeError("LLM 预算已用尽")
                raw = self.judge_fn(query, [candidates[index] for index in order], rules, examples)
                self.calls += 1
                self.cache[key] = raw
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                with self.cache_path.open("a", encoding="utf-8") as handle:
                    handle.write(json.dumps({"key": key, "result": raw, "model": self.model, "operated_at": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False) + "\n")
            verdicts.append(self.validate(raw, candidate_ids))
        if verdicts[0] != verdicts[1]:
            raise JudgeError("两种候选顺序得到的 LLM 结果不一致")
        return verdicts[0]
