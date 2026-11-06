"""Multi-route retrieval index for codebook names and aliases."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import numpy as np
from rapidfuzz import fuzz, process
from sklearn.feature_extraction.text import TfidfVectorizer

from .codebook import Codebook, CodebookRow, normalize_name
from .normalize import NormalizedName, normalize_name as runtime_normalize


@dataclass(frozen=True)
class IndexEntry:
    entity_id: str
    row: CodebookRow
    variant: str
    variant_type: str
    variant_norm: str
    core: str


class NameIndex:
    def __init__(self, codebook: Codebook):
        self.codebook = codebook
        self.entries: list[IndexEntry] = []
        for row in codebook.valid_rows():
            # ``core`` is a property of the formal entity name, not of the
            # matched surface form.  Keeping it on every alias entry makes
            # core retrieval work for both formal names and aliases.
            row_runtime = runtime_normalize(row.w_name)
            row_norm = row_runtime.norm or row.norm_w_name
            row_core = row_runtime.core or row_norm
            self.entries.append(IndexEntry(row.entity_id, row, row.w_name, "w_name", row_norm, row_core))
            for alias in row.aliases:
                alias_norm = runtime_normalize(alias).norm or normalize_name(alias)
                if alias_norm:
                    self.entries.append(IndexEntry(row.entity_id, row, alias, "alias", alias_norm, row_core))
        self.exact: dict[tuple[str, str], list[IndexEntry]] = {}
        self.core: dict[tuple[str, str], list[IndexEntry]] = {}
        self.alias_entities: dict[tuple[str, str], set[str]] = {}
        for entry in self.entries:
            for city in codebook.split_values(entry.row.cityname):
                city_norm = normalize_name(city)
                self.exact.setdefault((city_norm, entry.variant_norm), []).append(entry)
                self.core.setdefault((city_norm, entry.core), []).append(entry)
                if entry.variant_type == "alias":
                    self.alias_entities.setdefault((city_norm, entry.variant_norm), set()).add(entry.entity_id)
        self.vectorizer = TfidfVectorizer(analyzer="char", ngram_range=(1, 3), sublinear_tf=True)
        self.matrix = self.vectorizer.fit_transform([entry.variant_norm for entry in self.entries]) if self.entries else None

    def exact_entries(self, cityname: str, normalized: NormalizedName) -> list[IndexEntry]:
        city = normalize_name(cityname)
        if not city:
            return [entry for entry in self.entries if entry.variant_norm == normalized.norm]
        return list(self.exact.get((city, normalized.norm), ()))

    def alias_entries(self, cityname: str, normalized: NormalizedName) -> list[IndexEntry]:
        return [entry for entry in self.exact_entries(cityname, normalized) if entry.variant_type == "alias"]

    def core_entries(self, cityname: str, normalized: NormalizedName) -> list[IndexEntry]:
        if not normalize_name(cityname):
            return [entry for entry in self.entries if entry.core == normalized.core]
        return list(self.core.get((normalize_name(cityname), normalized.core), ()))

    def alias_is_ambiguous(self, entry: IndexEntry, cityname: str) -> bool:
        """Whether an alias maps to more than one entity in the city."""

        if entry.variant_type != "alias":
            return False
        cities = self.codebook.split_values(entry.row.cityname) if not cityname else (cityname,)
        return any(
            len(self.alias_entities.get((normalize_name(city), entry.variant_norm), set())) > 1
            for city in cities
        )

    def retrieve(self, cityname: str, normalized: NormalizedName, top_k: int = 10) -> list[tuple[IndexEntry, float, str]]:
        if not normalized.norm or not self.entries:
            return []
        exact = self.exact_entries(cityname, normalized)
        core = self.core_entries(cityname, normalized)
        seen: set[str] = set()
        results: list[tuple[IndexEntry, float, str]] = []
        for entry in exact:
            if entry.entity_id not in seen:
                results.append((entry, 1.0, "exact" if entry.variant_type == "w_name" else "alias"))
                seen.add(entry.entity_id)
        for entry in core:
            if entry.entity_id not in seen:
                results.append((entry, 0.92, "core"))
                seen.add(entry.entity_id)
        # An unknown city must not erase all candidates.  When a city is
        # supplied, restrict retrieval to that city; otherwise search the
        # complete index and let policy route the result to review.
        city_entries = (
            [entry for entry in self.entries if self.codebook.row_applies_to_city(entry.row, cityname)]
            if cityname
            else list(self.entries)
        )
        choices: dict[str, list[IndexEntry]] = {}
        for entry in city_entries:
            choices.setdefault(entry.variant_norm, []).append(entry)
        for variant, score, _ in process.extract(
            normalized.norm,
            list(choices),
            scorer=fuzz.ratio,
            limit=max(top_k * 3, 20),
            score_cutoff=0,
        ):
            for entry in choices[variant]:
                if entry.entity_id not in seen:
                    results.append((entry, score / 100.0, "rapidfuzz"))
                    seen.add(entry.entity_id)
        if self.matrix is not None:
            query = self.vectorizer.transform([normalized.norm])
            scores = (query @ self.matrix.T).toarray()[0]
            eligible = [
                index
                for index, entry in enumerate(self.entries)
                if (not cityname or self.codebook.row_applies_to_city(entry.row, cityname))
            ]
            eligible.sort(key=lambda index: (-float(scores[index]), self.entries[index].entity_id))
            for index in eligible[: max(top_k * 3, 20)]:
                entry = self.entries[int(index)]
                if entry.entity_id in seen:
                    continue
                results.append((entry, float(scores[int(index)]), "tfidf"))
                seen.add(entry.entity_id)
        results.sort(key=lambda item: (-item[1], item[0].entity_id))
        return results[:top_k]
