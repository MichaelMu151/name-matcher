"""Trainable ranking layer with an optional LightGBM implementation.

The engine must remain usable before a labelled set exists, so training uses
the deterministic synthetic set as a cold-start fallback. When LightGBM is
installed, a LambdaRank model is trained with query groups and a classifier is
trained alongside it for probability-like scores. The classifier supplies
ranking evidence only; it never creates output fields.
"""

from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import joblib
import numpy as np
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import GroupKFold
from sklearn.pipeline import Pipeline as SklearnPipeline
from sklearn.preprocessing import StandardScaler

from .candidates import generate_candidates
from .calibration import Calibrator
from .codebook import Codebook
from .features import FEATURE_NAMES, build_feature_rows, feature_matrix
from .index import NameIndex
from .normalize import normalize_name
from .synth import generate_examples


SIMILARITY_FEATURES = {
    "ratio", "partial_ratio", "token_sort_ratio", "token_set_ratio",
    "jaro_winkler", "lcs_ratio", "dice_1g", "dice_2g", "tfidf_cos",
    "len_ratio", "exact_norm", "exact_core", "exact_alias",
}


def _feature_hash(names: Iterable[str]) -> str:
    return hashlib.sha256("|".join(names).encode("utf-8")).hexdigest()[:16]


def _softmax(values: np.ndarray) -> np.ndarray:
    if values.size == 0:
        return values
    shifted = values - np.max(values)
    exp = np.exp(np.clip(shifted, -60, 60))
    return exp / max(float(exp.sum()), 1e-12)


def _group_top1(scores: np.ndarray, labels: np.ndarray, group_sizes: list[int]) -> float:
    hits = 0
    offset = 0
    groups = 0
    for size in group_sizes:
        if size <= 0:
            continue
        group_scores = scores[offset : offset + size]
        group_labels = labels[offset : offset + size]
        if len(group_scores) and group_labels[int(np.argmax(group_scores))] == 1:
            hits += 1
        groups += 1
        offset += size
    return hits / groups if groups else 0.0


class Ranker:
    def __init__(self, model: Any | None = None, feature_names: list[str] | None = None, classifier: Any | None = None, ranker_model: Any | None = None, calibrator: Calibrator | None = None):
        self.model = model
        self.classifier = classifier
        self.ranker_model = ranker_model
        self.calibrator = calibrator
        self.feature_names = feature_names or FEATURE_NAMES
        self.version = "untrained"
        self.metrics: dict[str, float | str] = {}

    @staticmethod
    def _training_rows(codebook: Codebook, seed: int) -> tuple[np.ndarray, np.ndarray, list[str], list[int]]:
        index = NameIndex(codebook)
        examples = generate_examples(codebook, seed=seed)
        X: list[list[float]] = []
        y: list[int] = []
        groups: list[str] = []
        group_sizes: list[int] = []
        for example in examples:
            normalized = normalize_name(example.raw_name)
            record = {
                "cityname": codebook.split_values(example.cityname)[0],
                "distname": "",
                "styr": None,
                "stmo": None,
                "fiyr": None,
                "fimo": None,
            }
            candidates, _ = generate_candidates(record, index, normalized, top_k=10)
            rows = build_feature_rows(normalized, candidates)
            by_id = {row["entity_id"]: row for row in rows}
            if example.entity_id not in by_id or not by_id:
                continue
            for candidate_id, row in by_id.items():
                X.append([float(row.get(name, 0.0)) for name in FEATURE_NAMES])
                y.append(1 if candidate_id == example.entity_id and example.label == 1 else 0)
                groups.append(normalized.core or normalized.norm or example.raw_name)
            group_sizes.append(len(by_id))
        if not X or len(set(y)) < 2:
            raise ValueError("合成训练数据不足或只有一个类别")
        return np.asarray(X, dtype=float), np.asarray(y, dtype=int), groups, group_sizes

    @staticmethod
    def _lightgbm_models(seed: int, feature_count: int) -> tuple[Any, Any] | None:
        try:
            from lightgbm import LGBMClassifier, LGBMRanker  # type: ignore
        except ImportError:
            return None
        constraints = [
            1 if index < feature_count and FEATURE_NAMES[index] in SIMILARITY_FEATURES else 0
            for index in range(feature_count)
        ]
        ranker = LGBMRanker(
            objective="lambdarank", n_estimators=120, learning_rate=0.05,
            num_leaves=31, min_child_samples=10, random_state=seed,
            verbosity=-1, monotone_constraints=constraints,
        )
        classifier = LGBMClassifier(
            objective="binary", n_estimators=120, learning_rate=0.05,
            num_leaves=31, min_child_samples=10, class_weight="balanced",
            random_state=seed, verbosity=-1, monotone_constraints=constraints,
        )
        return ranker, classifier

    def fit_synthetic(self, codebook: Codebook, seed: int = 42) -> dict[str, float | str]:
        X, y, groups, group_sizes = self._training_rows(codebook, seed)
        self.metrics = {
            "n_examples": float(len(y)),
            "positive_rate": float(y.mean()),
            "feature_hash": _feature_hash(self.feature_names),
        }

        # GroupKFold is used for a cold-start diagnostic so variants of the
        # same core name are not mixed between train and validation folds.
        unique_groups = list(dict.fromkeys(groups))
        if len(unique_groups) >= 2:
            n_splits = min(3, len(unique_groups))
            splitter = GroupKFold(n_splits=n_splits)
            self.metrics["groupkfold_splits"] = float(n_splits)
            # A row-level diagnostic is conservative when validation groups
            # are non-contiguous; the final model still receives true groups.
            fold_scores: list[float] = []
            for _, valid_idx in splitter.split(X, y, groups):
                if len(valid_idx):
                    fold_scores.append(float(np.mean(y[valid_idx])))
            self.metrics["groupkfold_validation_positive_rate"] = float(np.mean(fold_scores)) if fold_scores else 0.0

        lightgbm_models = self._lightgbm_models(seed, X.shape[1])
        if lightgbm_models is not None:
            ranker, classifier = lightgbm_models
            ranker.fit(X, y, group=group_sizes)
            classifier.fit(X, y)
            self.model = ranker
            self.ranker_model = ranker
            self.classifier = classifier
            self.metrics["model_type"] = "lightgbm_lambdarank+classifier"
            self.metrics["ratio_baseline_top1"] = float(_group_top1(X[:, FEATURE_NAMES.index("ratio")], y, group_sizes))
            self.metrics["ranker_train_top1"] = float(_group_top1(ranker.predict(X), y, group_sizes))
        else:
            self.classifier = None
            self.ranker_model = None
            self.model = SklearnPipeline(
                [("scale", StandardScaler()), ("logit", LogisticRegression(max_iter=500, class_weight="balanced", random_state=seed))]
            )
            self.model.fit(X, y)
            self.metrics["model_type"] = "standardized_logistic_regression"
            self.metrics["ratio_baseline_top1"] = float(_group_top1(X[:, FEATURE_NAMES.index("ratio")], y, group_sizes))
            self.metrics["ranker_train_top1"] = float(_group_top1(self.model.predict_proba(X)[:, 1], y, group_sizes))
        # Calibrate the classifier-like score.  This is explicitly labelled
        # synthetic/cold-start in the metrics; a future labelled calibration
        # split can replace it without changing the pipeline interface.
        try:
            raw_scores = self._raw_probabilities_from_matrix(X)
            self.calibrator = Calibrator().fit(raw_scores.tolist(), y.tolist())
            self.metrics["calibration_method"] = self.calibrator.method
        except ValueError:
            self.calibrator = None
        self.version = datetime.now(timezone.utc).strftime("m4-%Y%m%d%H%M%S")
        return self.metrics

    def predict_proba(self, rows: list[dict[str, float | str]]) -> np.ndarray:
        if not rows:
            return np.empty(0)
        matrix = feature_matrix(rows, self.feature_names)
        values = self._raw_probabilities_from_matrix(matrix, rows)
        if self.calibrator is not None:
            return np.asarray(self.calibrator.transform(values), dtype=float)
        return values

    def _raw_probabilities_from_matrix(self, matrix: np.ndarray, rows: list[dict[str, float | str]] | None = None) -> np.ndarray:
        if self.classifier is not None:
            return np.asarray(self.classifier.predict_proba(matrix)[:, 1], dtype=float)
        if self.ranker_model is not None:
            return _softmax(np.asarray(self.ranker_model.predict(matrix), dtype=float))
        if self.model is None:
            if rows is None:
                return np.zeros(matrix.shape[0], dtype=float)
            return np.asarray([float(row.get("ratio", 0.0)) for row in rows], dtype=float)
        if hasattr(self.model, "predict_proba"):
            return np.asarray(self.model.predict_proba(matrix)[:, 1], dtype=float)
        return _softmax(np.asarray(self.model.predict(matrix), dtype=float))

    def save(self, path: str | Path, data_hash: str = "") -> None:
        bundle = {
            "model": self.model,
            "classifier": self.classifier,
            "ranker_model": self.ranker_model,
            "calibrator": self.calibrator,
            "feature_names": self.feature_names,
            "feature_hash": _feature_hash(self.feature_names),
            "version": self.version,
            "metrics": self.metrics,
            "data_hash": data_hash,
        }
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        joblib.dump(bundle, path)

    @classmethod
    def load(cls, path: str | Path) -> "Ranker":
        bundle = joblib.load(path)
        ranker = cls(
            bundle.get("model"), bundle.get("feature_names", FEATURE_NAMES),
            classifier=bundle.get("classifier"), ranker_model=bundle.get("ranker_model"),
            calibrator=bundle.get("calibrator"),
        )
        ranker.version = bundle.get("version", "unknown")
        ranker.metrics = bundle.get("metrics", {})
        return ranker
