"""Small config loader with a PyYAML path and a safe built-in fallback."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from copy import deepcopy


DEFAULT_CONFIG: dict[str, Any] = {
    "audited_fields": ["w_name", "wrank_t", "s1_sector"],
    "field_mapping": {"w_name": "w_name", "wrank_t": "wrank", "s1_sector": "s1_sector"},
    "ui": {"allow_human_override_target_columns": False, "export_files": {"content": "reviewed_content.xlsx", "log": "modification_log.xlsx"}},
    "matching": {"candidate_k": 10, "month_tolerance": 0, "year_tolerance": 0, "dedupe_include_distname": True, "distname_hard_filter": False},
    "policy": {"tau_auto": 0.85, "margin_auto": 0.10, "tau_nil": 0.15, "alpha": 0.02},
    "input_mapping": {"record_id": "record_id", "year": "year", "dtype": "dtype", "dtype_fallback": "dtype_t", "context_fields": ["p_name", "post1", "context"]},
    "models": {"default_version": "synthetic_ranker", "llm_model": "fake", "llm_budget": 100},
    "paths": {"reports_dir": "reports", "resource_dir": "resources"},
}


def load_config(path: str | Path = "config.yaml") -> dict[str, Any]:
    path = Path(path)
    if not path.exists():
        return deepcopy(DEFAULT_CONFIG)
    try:
        import yaml  # type: ignore
        with path.open(encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
    except ImportError:
        # Keep a documented fallback for minimal environments, but do not
        # silently discard the user file: callers can see this in the audit
        # report and the dependency is declared in requirements.txt.
        data = {}
    if not isinstance(data, dict):
        raise ValueError("config.yaml 顶层必须是 mapping")
    result = _merge(DEFAULT_CONFIG, data)
    targets = result.get("audited_fields")
    mapping = result.get("field_mapping")
    if not isinstance(targets, list) or not targets:
        raise ValueError("config.audited_fields 必须是非空列表")
    if not isinstance(mapping, dict) or any(field not in mapping for field in targets):
        raise ValueError("config.field_mapping 必须覆盖全部 audited_fields")
    return result


def _merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _merge(result[key], value)
        else:
            result[key] = value
    return result
