"""Pure, conservative normalization for department names."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path


_PUNCT_RE = re.compile(r"[\s\u3000，,。；;：:、·•/\\|（）()【】\[\]《》“”\"'‘’—–_-]+")
_BRACKET_RE = re.compile(r"（[^）]*）|\([^)]*\)|【[^】]*】|\[[^]]*\]")
# Delimiters for a cell that contains more than one department.  ``和/及/与``
# are intentionally bounded by non-Chinese punctuation/word boundaries only
# through the post-split length check below; this avoids treating ordinary
# words such as ``委员会`` as two entities.
_MULTI_RE = re.compile(r"、|/|\\|;|；|和|及|与")
_MODIFIERS = ("党组", "机关党委", "党委", "政府", "机关")
_DEFAULT_SUFFIXES = ("管理委员会", "监督管理局", "委员会", "办公室", "管理局", "指挥部", "服务中心", "中心", "研究院", "学院", "大学", "医院", "集团公司", "有限公司", "总公司", "公司", "工委", "党委", "纪委", "局", "委", "厅", "办", "处", "署", "院", "会")


@dataclass(frozen=True)
class NormalizedName:
    raw: str
    norm: str
    core: str
    suffix_type: str
    city_hint: str | None
    modifiers: tuple[str, ...] = ()
    bracket_content: tuple[str, ...] = ()
    segments: tuple[str, ...] = ()
    flags: tuple[str, ...] = ()


def _read_lines(path: str | Path | None, defaults: tuple[str, ...]) -> tuple[str, ...]:
    if path is None or not Path(path).exists():
        return defaults
    return tuple(line.strip() for line in Path(path).read_text(encoding="utf-8").splitlines() if line.strip())


def normalize_name(raw: object, admin_divisions: str | Path | None = None, suffixes: str | Path | None = None) -> NormalizedName:
    text = "" if raw is None else str(raw)
    text = unicodedata.normalize("NFKC", text).strip()
    brackets = tuple(match.group(0)[1:-1] for match in _BRACKET_RE.finditer(text))
    main = _BRACKET_RE.sub("", text)
    divisions = sorted(_read_lines(admin_divisions, ("广州市", "深圳市", "广东省")), key=len, reverse=True)
    city_hint = next((city for city in divisions if main.startswith(city)), None)
    if city_hint:
        main = main[len(city_hint):]
    main = _PUNCT_RE.sub("", main)
    norm = main
    modifiers = tuple(item for item in _MODIFIERS if item in norm)
    suffix_values = sorted(_read_lines(suffixes, _DEFAULT_SUFFIXES), key=len, reverse=True)
    suffix_type = next((suffix for suffix in suffix_values if norm.endswith(suffix)), "")
    core = norm[:-len(suffix_type)] if suffix_type else norm
    raw_segments = [segment.strip() for segment in _MULTI_RE.split(text) if segment.strip()]
    segments = tuple(raw_segments)
    flags: list[str] = []
    if not norm:
        flags.append("empty")
    if len(segments) > 1 and all(len(_PUNCT_RE.sub("", part)) >= 2 for part in segments):
        flags.append("multi_entity")
    if brackets:
        flags.append("has_bracket_content")
    if modifiers:
        flags.append("has_modifier")
    return NormalizedName(
        raw=text,
        norm=norm,
        core=core,
        suffix_type=suffix_type,
        city_hint=city_hint,
        modifiers=modifiers,
        bracket_content=brackets,
        segments=segments,
        flags=tuple(sorted(set(flags))),
    )


def split_multi_entity(normalized: NormalizedName) -> tuple[str, ...]:
    if "multi_entity" not in normalized.flags:
        return (normalized.norm,) if normalized.norm else ()
    return tuple(normalize_name(part).norm for part in normalized.segments if normalize_name(part).norm)
