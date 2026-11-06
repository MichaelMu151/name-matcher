from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from src.engine.codebook import Codebook, CodebookError, REQUIRED_COLUMNS


def make_fixture(path: Path) -> None:
    rows = [
        {
            "cityname": "测试市",
            "distname": "甲区",
            "dtype": "地级市",
            "drank": "副部级",
            "s1_sector": "02政府",
            "wrank": "副厅级",
            "w_name": "市场监督管理局",
            "alias": "市监局；市场监管局",
            "notes": "合成样本：2020年后调整",
            "year_valid_from": None,
            "month_valid_from": None,
            "year_valid_to": None,
            "month_valid_to": None,
        },
        {
            "cityname": "测试市",
            "distname": "甲区",
            "dtype": "地级市",
            "drank": "副部级",
            "s1_sector": "02政府",
            "wrank": "处级",
            "w_name": "住房公积金管理中心",
            "alias": "公积金中心|住房公积金中心",
            "notes": "合成样本",
            "year_valid_from": 2000,
            "month_valid_from": 1,
            "year_valid_to": 2010,
            "month_valid_to": 12,
        },
        {
            "cityname": "测试市",
            "distname": "甲区",
            "dtype": "地级市",
            "drank": "副部级",
            "s1_sector": "02政府",
            "wrank": "副厅级",
            "w_name": "住房公积金管理中心",
            "alias": "公积金中心；共享别名",
            "notes": "合成样本",
            "year_valid_from": 2010,
            "month_valid_from": 1,
            "year_valid_to": 2020,
            "month_valid_to": 12,
        },
        {
            "cityname": "测试市",
            "distname": "甲区",
            "dtype": "地级市",
            "drank": "副部级",
            "s1_sector": "01党委",
            "wrank": "副厅级",
            "w_name": "组织部",
            "alias": "共享别名；市场监督管理局",
            "notes": "撤销后更名",
            "year_valid_from": 2025,
            "month_valid_from": 13,
            "year_valid_to": 2024,
            "month_valid_to": 1,
        },
        {
            "cityname": "测试市 | 另一市",
            "distname": "甲区",
            "dtype": "地级市",
            "drank": "副部级",
            "s1_sector": "02政府",
            "wrank": "副厅级",
            "w_name": "政务服务中心",
            "alias": "政务中心",
            "notes": "合成样本",
            "year_valid_from": None,
            "month_valid_from": 3,
            "year_valid_to": None,
            "month_valid_to": None,
        },
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(REQUIRED_COLUMNS)
    for row in rows:
        sheet.append([row[column] for column in REQUIRED_COLUMNS])
    workbook.save(path)


class CodebookM0Tests(unittest.TestCase):
    def test_synthetic_fixture_is_loadable(self) -> None:
        fixture = Path(__file__).parent / "fixtures" / "synthetic_codebook.csv"
        codebook = Codebook(fixture)
        self.assertEqual(len(codebook.rows), 5)
        self.assertEqual(codebook.rows[0].entity_id, Codebook(fixture).rows[0].entity_id)

    def test_contract_and_stable_ids(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fixture.xlsx"
            make_fixture(path)
            first = Codebook(path)
            second = Codebook(path)
            self.assertEqual([row.entity_id for row in first.rows], [row.entity_id for row in second.rows])
            self.assertTrue(all(len(row.entity_id) == 10 for row in first.rows))
            self.assertTrue(all(len(row.dept_key) == 10 for row in first.rows))

    def test_audit_detects_required_m0_findings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fixture.xlsx"
            reports = Path(directory) / "reports"
            make_fixture(path)
            codebook = Codebook(path, reports_dir=reports)
            audit = codebook.write_reports()
            self.assertGreaterEqual(len(audit.period_conflicts), 2)
            self.assertGreaterEqual(len(audit.invalid_months), 1)
            self.assertGreaterEqual(len(audit.year_month_incomplete), 1)
            self.assertGreaterEqual(len(audit.reversed_periods), 1)
            self.assertGreaterEqual(len(audit.ambiguous_aliases), 1)
            self.assertGreaterEqual(len(audit.alias_name_collisions), 1)
            self.assertGreaterEqual(len(audit.notes_need_review), 2)
            self.assertGreaterEqual(len(audit.multi_city_values), 1)
            self.assertTrue((reports / "codebook_audit.md").exists())
            self.assertTrue((reports / "period_conflicts.csv").exists())
            self.assertTrue((reports / "notes_need_review.csv").exists())

    def test_missing_contract_columns_fail_fast(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "legacy.xlsx"
            workbook = Workbook()
            workbook.active.append(["dept_id", "level", "standard_name", "aliases"])
            workbook.save(path)
            with self.assertRaises(CodebookError):
                Codebook(path)


if __name__ == "__main__":
    unittest.main()
