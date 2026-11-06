"""Small independent probability calibration utility for M5."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
from sklearn.isotonic import IsotonicRegression
from sklearn.linear_model import LogisticRegression


@dataclass
class Calibrator:
    method: str = "platt"
    model: object | None = None

    def fit(self, scores: list[float], labels: list[int]) -> "Calibrator":
        if len(scores) != len(labels) or len(scores) < 2:
            raise ValueError("校准数据不足")
        if len(set(int(label) for label in labels)) < 2:
            raise ValueError("校准标签必须同时包含正负类")
        if len(scores) >= 300:
            self.method = "isotonic"
            self.model = IsotonicRegression(out_of_bounds="clip").fit(scores, labels)
        else:
            self.method = "platt"
            self.model = LogisticRegression(max_iter=500).fit(np.asarray(scores).reshape(-1, 1), labels)
        return self

    def transform(self, scores: list[float] | np.ndarray) -> np.ndarray:
        values = np.asarray(scores, dtype=float)
        if self.model is None:
            return np.clip(values, 0, 1)
        if self.method == "isotonic":
            return np.asarray(self.model.predict(values), dtype=float)
        return self.model.predict_proba(values.reshape(-1, 1))[:, 1]

    @staticmethod
    def choose_tau(scores: list[float], labels: list[int], alpha: float = 0.02) -> float:
        if len(scores) != len(labels) or not scores:
            return 1.0
        # Finite-sample correction: require at most floor(alpha*n) errors in
        # the accepted set, and apply a one-count safety margin.  This is a
        # conservative Learn-then-Test-style threshold rather than selecting
        # a threshold on a held-out test set.
        max_errors = max(0, int(np.floor(alpha * len(scores))) - 1)
        values = sorted(set(float(score) for score in scores), reverse=True)
        for threshold in values:
            selected = [label for score, label in zip(scores, labels) if score >= threshold]
            errors = len(selected) - sum(selected)
            if selected and errors <= max_errors:
                return threshold
        return 1.0

    def report(self, scores: list[float], labels: list[int], alpha: float = 0.02) -> dict[str, float | str]:
        calibrated = self.transform(scores)
        threshold = self.choose_tau(calibrated.tolist(), labels, alpha=alpha)
        selected = calibrated >= threshold
        return {
            "method": self.method,
            "tau": float(threshold),
            "coverage": float(selected.mean()) if len(selected) else 0.0,
            "accepted_error_rate": float(1 - np.asarray(labels)[selected].mean()) if selected.any() else 0.0,
        }
