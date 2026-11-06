from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd

from src.engine.candidates import generate_candidates, input_period
from src.engine.codebook import Codebook
from src.engine.llm_judge import FakeJudge, LLMJudge
from src.engine.normalize import normalize_name
from src.engine.pipeline import Pipeline
from src.engine.index import NameIndex


class EngineSmokeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.fixture = Path(__file__).parent / "fixtures" / "synthetic_codebook.csv"
        self.codebook = Codebook(self.fixture, reports_dir=Path(tempfile.mkdtemp()))

    def test_normalization_preserves_brackets_and_city_hint(self) -> None:
        value = normalize_name("广州市市场监督管理局（挂牌后）")
        self.assertEqual(value.city_hint, "广州市")
        self.assertEqual(value.norm, "市场监督管理局")
        self.assertEqual(value.bracket_content, ("挂牌后",))
        self.assertEqual(value.suffix_type, "监督管理局")

    def test_exact_candidate_and_period_flags(self) -> None:
        index = NameIndex(self.codebook)
        normalized = normalize_name("市场监管局")
        candidates, flags = generate_candidates({"cityname": "测试市", "distname": "甲区", "styr": 2021, "fiyr": None, "fimo": None}, index, normalized)
        self.assertTrue(candidates)
        self.assertIn("alias", {candidate.retrieval_path for candidate in candidates})
        self.assertIn("period_unknown", flags)

    def test_invalid_input_period_is_flagged(self) -> None:
        period = input_period({"styr": 2020, "stmo": 1, "fiyr": None, "fimo": 3})
        self.assertTrue(period.invalid)

    def test_llm_candidate_whitelist(self) -> None:
        judge = LLMJudge(judge_fn=FakeJudge("NONE"), cache_path=Path(tempfile.mkdtemp()) / "cache.jsonl")
        verdict = judge.judge({"w_name_o": "x"}, [{"entity_id": "abc", "w_name": "x"}])
        self.assertEqual(verdict.decision, "NONE")

    def test_pipeline_exports_target_columns_only(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            input_path = directory / "input.xlsx"
            pd.DataFrame([
                {"record_id": "r1", "year": 2020, "styr": 2019, "stmo": 1, "fiyr": None, "fimo": None, "cityname": "测试市", "distname": "甲区", "w_name_o": "市场监管局", "keep": "unchanged", "w_name": "旧值", "wrank_t": "旧级别", "s1_sector": "旧类别"},
            ]).to_excel(input_path, index=False)
            pipeline = Pipeline(self.fixture, config_path="config.yaml", llm_judge=LLMJudge(judge_fn=FakeJudge("NONE"), cache_path=directory / "cache.jsonl"))
            result = pipeline.run(input_path, directory / "outputs", run_id="smoke")
            output = pd.read_excel(result["reviewed_content"], dtype=object)
            source = pd.read_excel(input_path, dtype=object)
            self.assertEqual(output["keep"].tolist(), source["keep"].tolist())
            self.assertIn(output.loc[0, "w_name"], {"旧值", "市场监督管理局"})
            self.assertTrue(Path(result["modification_log"]).exists())


if __name__ == "__main__":
    unittest.main()
