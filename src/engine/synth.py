"""Synthetic cold-start training data generation."""

from __future__ import annotations

import random
from dataclasses import asdict, dataclass

from .codebook import Codebook
from .normalize import normalize_name


@dataclass(frozen=True)
class SyntheticExample:
    raw_name: str
    entity_id: str
    cityname: str
    source: str = "synthetic"
    label: int = 1


def _delete_one(value: str, rng: random.Random) -> str:
    if len(value) <= 2:
        return value
    index = rng.randrange(len(value))
    return value[:index] + value[index + 1:]


def generate_examples(codebook: Codebook, seed: int = 42, include_negatives: bool = True) -> list[SyntheticExample]:
    rng = random.Random(seed)
    examples: list[SyntheticExample] = []
    for row in codebook.valid_rows():
        variants = {row.w_name, *row.aliases}
        for variant in variants:
            if not variant:
                continue
            examples.append(SyntheticExample(variant, row.entity_id, row.cityname))
            if len(variant) > 3:
                examples.append(SyntheticExample(_delete_one(variant, rng), row.entity_id, row.cityname))
            for city in codebook.split_values(row.cityname)[:1]:
                examples.append(SyntheticExample(city + variant, row.entity_id, row.cityname))
    if include_negatives:
        rows = codebook.valid_rows()
        for row in rows:
            other = next((item for item in rows if item.cityname == row.cityname and item.entity_id != row.entity_id), None)
            if other:
                examples.append(SyntheticExample(other.w_name, row.entity_id, row.cityname, label=0))
    return examples
