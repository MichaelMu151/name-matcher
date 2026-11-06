"""Pydantic boundary schemas with a minimal-runtime fallback."""

from __future__ import annotations

from typing import Any

try:
    from pydantic import BaseModel, ConfigDict, Field
except ImportError:  # pragma: no cover - only used before dependency install
    BaseModel = None  # type: ignore[assignment,misc]


if BaseModel is not None:

    class LLMResponse(BaseModel):
        model_config = ConfigDict(extra="forbid")

        decision: str
        reason: str = Field(default="", max_length=80)
        rule_refs: list[str] = Field(default_factory=list)
        uncertain: bool = False


    class InputRecord(BaseModel):
        model_config = ConfigDict(extra="allow")

        record_id: str | None = None
        year: Any | None = None
        styr: Any | None = None
        stmo: Any | None = None
        fiyr: Any | None = None
        fimo: Any | None = None
        cityname: str = ""
        distname: str = ""
        dtype: Any | None = None
        w_name_o: str = ""
        context: str = ""


    class DecisionRecord(BaseModel):
        model_config = ConfigDict(extra="forbid")

        decision: str
        entity_id: str | None = None
        method: str
        p_calibrated: float = 0.0
        margin: float = 0.0
        reason: str = ""
        flags: list[str] = Field(default_factory=list)

else:

    class LLMResponse:  # type: ignore[no-redef]
        def __init__(self, decision: str, reason: str = "", rule_refs: list[str] | None = None, uncertain: bool = False):
            if not isinstance(decision, str) or not isinstance(reason, str) or len(reason) > 80 or not isinstance(uncertain, bool):
                raise ValueError("invalid LLM response")
            if rule_refs is None:
                rule_refs = []
            if not isinstance(rule_refs, list) or not all(isinstance(item, str) for item in rule_refs):
                raise ValueError("invalid rule_refs")
            self.decision = decision
            self.reason = reason
            self.rule_refs = rule_refs
            self.uncertain = uncertain

        @classmethod
        def model_validate(cls, raw: Any) -> "LLMResponse":
            if not isinstance(raw, dict) or set(raw) - {"decision", "reason", "rule_refs", "uncertain"}:
                raise ValueError("invalid response object")
            if "decision" not in raw:
                raise ValueError("decision is required")
            return cls(raw["decision"], raw.get("reason", ""), raw.get("rule_refs", []), raw.get("uncertain", False))

    class InputRecord:  # type: ignore[no-redef]
        pass

    class DecisionRecord:  # type: ignore[no-redef]
        pass
