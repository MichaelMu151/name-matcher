from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from src.engine.candidates import generate_candidates
from src.engine.codebook import Codebook
from src.engine.index import NameIndex
from src.engine.normalize import normalize_name
from src.engine.pipeline import Pipeline
from src.engine.policy import decide
from src.engine.features import build_feature_rows
from src.engine.llm_judge import JudgeError, LLMJudge
from src.engine.calibration import Calibrator
from src.engine.feedback import export_review, ingest_feedback


class AuditRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.fixture = Path(__file__).parent / "fixtures" / "synthetic_codebook.csv"
        cls.codebook = Codebook(cls.fixture, reports_dir=Path(tempfile.mkdtemp()))
        cls.index = NameIndex(cls.codebook)

    def test_shared_alias_is_flagged_and_never_clean_exact(self) -> None:
        normalized = normalize_name("共享别名")
        candidates, _ = generate_candidates({"cityname": "测试市", "distname": "甲区"}, self.index, normalized)
        self.assertGreaterEqual(len(candidates), 2)
        self.assertTrue(all("alias_ambiguous" in candidate.flags for candidate in candidates if candidate.matched_field == "alias"))
        decision = decide(candidates, [0.99] * len(candidates), [], tau_auto=0.85, margin_auto=0.1)
        self.assertNotEqual(decision.decision, "AUTO")

    def test_core_retrieval_uses_formal_core_for_aliases(self) -> None:
        normalized = normalize_name("市场")
        candidates, _ = generate_candidates({"cityname": "测试市"}, self.index, normalized)
        self.assertTrue(any(candidate.retrieval_path == "core" for candidate in candidates))
        rows = build_feature_rows(normalized, candidates)
        self.assertTrue(any(row["exact_core"] == 1.0 for row in rows))

    def test_missing_city_still_retrieves_candidates(self) -> None:
        normalized = normalize_name("完全不存在的机构")
        candidates, flags = generate_candidates({}, self.index, normalized)
        self.assertIn("city_missing", flags)
        self.assertIsInstance(candidates, list)

    def test_city_prefix_is_inferred_and_multi_entity_detected(self) -> None:
        normalized = normalize_name("广州市组织部和宣传部")
        self.assertEqual(normalized.city_hint, "广州市")
        self.assertIn("multi_entity", normalized.flags)
        candidates, flags = generate_candidates({"w_name_o": normalized.raw}, self.index, normalized)
        self.assertIn("city_inferred", flags)
        self.assertTrue(candidates or "city_missing" not in flags)

    def test_export_validation_rejects_non_target_change(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = directory / "source.xlsx"
            output = directory / "output.xlsx"
            frame = pd.DataFrame([{"keep": "x", "w_name": "old"}])
            frame.to_excel(source, index=False)
            altered = pd.DataFrame([{"keep": "x", "w_name": "new"}])
            Pipeline._export_content(source, output, altered, ("w_name",))
            workbook = load_workbook(output)
            workbook.active["A2"] = "changed"
            workbook.save(output)
            with self.assertRaises(ValueError):
                Pipeline._validate_preserved_content(source, output, ("w_name",))

    def test_invalid_codebook_time_is_not_clean_exact(self) -> None:
        normalized = normalize_name("组织部")
        candidates, _ = generate_candidates({"cityname": "测试市", "distname": "甲区"}, self.index, normalized)
        target = [candidate for candidate in candidates if candidate.row.w_name == "组织部"]
        self.assertTrue(target)
        self.assertIn("codebook_period_invalid", target[0].flags)
        decision = decide(candidates, [0.99] * len(candidates), [])
        self.assertNotEqual(decision.decision, "AUTO")

    def test_llm_schema_rejects_overlong_reason(self) -> None:
        with self.assertRaises(JudgeError):
            LLMJudge.validate({"decision": "NONE", "reason": "x" * 81, "rule_refs": [], "uncertain": False}, set())

    def test_calibration_has_conservative_threshold_report(self) -> None:
        calibrator = Calibrator().fit([0.1, 0.2, 0.8, 0.9], [0, 0, 1, 1])
        report = calibrator.report([0.1, 0.2, 0.8, 0.9], [0, 0, 1, 1], alpha=0.25)
        self.assertIn("coverage", report)
        self.assertGreaterEqual(float(report["tau"]), 0.0)

    def test_feedback_confirm_uses_system_entity_and_stays_pending(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = directory / "input.csv"
            debug = directory / "debug.jsonl"
            review = directory / "review.xlsx"
            pd.DataFrame([{"record_id": "r1", "w_name_o": "市场监管局"}]).to_csv(source, index=False)
            pd.DataFrame([{"record_id": "r1", "decision": "AUTO", "entity_id": "abc", "reason": "ok", "top3_candidates": [], "candidate_ids": ["abc"]}]).to_json(debug, orient="records", lines=True, force_ascii=False)
            export_review(source, debug, review)
            frame = pd.read_excel(review, dtype=object)
            frame.loc[0, "human_decision"] = "confirm"
            frame.to_excel(review, index=False)
            aliases = directory / "aliases.csv"
            ingest_feedback(review, aliases)
            aliases_frame = pd.read_csv(aliases, dtype=object)
            self.assertEqual(aliases_frame.loc[0, "entity_id"], "abc")
            self.assertEqual(aliases_frame.loc[0, "status"], "pending_second_review")


if __name__ == "__main__":
    unittest.main()
