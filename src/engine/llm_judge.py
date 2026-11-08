"""LLM arbitration interface with deterministic FakeJudge and safe validation."""

from __future__ import annotations

import hashlib
import json
import urllib.error
import urllib.request
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


class DeepSeekJudge:
    """OpenAI-compatible DeepSeek judge with a strict JSON-only contract.

    The API key is held only in memory by the callable. It is never written to
    the cache or logs. Network failures deliberately raise ``JudgeError`` so
    the pipeline routes the case to HUMAN instead of guessing.
    """

    def __init__(
        self,
        api_key: str,
        model: str = "deepseek-chat",
        base_url: str = "https://api.deepseek.com",
        timeout: float = 30.0,
    ):
        self.api_key = api_key.strip()
        self.model = model.strip() or "deepseek-chat"
        self.base_url = base_url.rstrip("/")
        self.timeout = float(timeout)
        if not self.api_key:
            raise ValueError("DeepSeek API key 不能为空")

    @staticmethod
    def _system_prompt() -> str:
        return (
            "你是政府部门名称编码审核器。只能在候选 entity_id 中选择，或返回 NONE、MULTIPLE、"
            "NOT_A_DEPARTMENT。绝不能创造候选、改写正式部门名称、推断候选之外的 entity_id，"
            "也不能因为名称相似而忽略城市、级别、时间和部门类型冲突。若证据不足、候选冲突、"
            "或无法区分，必须将 uncertain 设为 true。只输出一个 JSON 对象，不要 Markdown。"
            "字段必须是 decision、reason、rule_refs、uncertain；reason 不超过 80 个汉字，"
            "rule_refs 是字符串数组。"
        )

    def __call__(self, query: dict[str, Any], candidates: list[dict[str, Any]], rules: list[str], examples: list[dict[str, Any]]) -> dict[str, Any]:
        user_payload = {"query": query, "candidates": candidates, "rules": rules, "examples": examples}
        body = {
            "model": self.model,
            "temperature": 0,
            "max_tokens": 300,
            "response_format": {"type": "json_object"},
            "messages": [
                {"role": "system", "content": self._system_prompt()},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False, default=str)},
            ],
        }
        request = urllib.request.Request(
            f"{self.base_url}/chat/completions",
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:300]
            raise JudgeError(f"DeepSeek API HTTP {exc.code}: {detail}") from exc
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            raise JudgeError(f"DeepSeek API 请求失败: {exc}") from exc
        try:
            content = payload["choices"][0]["message"]["content"]
            if isinstance(content, list):
                content = "".join(str(item.get("text", "")) if isinstance(item, dict) else str(item) for item in content)
            raw = json.loads(str(content).strip())
        except (KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
            raise JudgeError("DeepSeek 返回格式不是可解析的 JSON") from exc
        if not isinstance(raw, dict):
            raise JudgeError("DeepSeek 返回的 JSON 不是对象")
        return raw


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
