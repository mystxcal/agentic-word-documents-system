import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from agentic_docs.sources_excel import export_dataset


class GenericExcelTests(unittest.TestCase):
    def _workbook(self, rows) -> Path:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        path = Path(directory.name) / "arbitrary-input.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Any Sheet"
        for row in rows:
            sheet.append(row)
        workbook.save(path)
        workbook.close()
        return path

    def test_arbitrary_columns_are_not_hardcoded(self):
        path = self._workbook(
            [
                ["Odd column name", "Another Field", "Free Text"],
                ["A-2", 20, "Keep this"],
                ["A-1", 10, "Ignore this"],
            ]
        )
        dataset = export_dataset(
            path,
            {"sheet": "Any Sheet", "range": "A1:C3"},
            {
                "columns": [
                    {"source": "Free Text", "heading": "Description"},
                    "Odd column name",
                ],
                "filters": [{"column": "Free Text", "op": "contains", "value": "keep"}],
                "sort": [{"column": "Another Field", "direction": "desc"}],
            },
            component_id="arbitrary-table",
        )
        self.assertEqual([column["heading"] for column in dataset["columns"]], ["Description", "Odd column name"])
        self.assertEqual(dataset["records"][0]["values"], ["Keep this", "A-2"])
        self.assertEqual(dataset["selected_row_count"], 1)

    def test_duplicate_source_headers_are_refused(self):
        path = self._workbook([["Same", "Same"], [1, 2]])
        with self.assertRaisesRegex(ValueError, "unique"):
            export_dataset(path, {"sheet": "Any Sheet", "range": "A1:B2"})

    def test_formula_policy_exposes_or_rejects_uncalculated_formulas(self):
        path = self._workbook([["Item", "Total"], ["A", "=1+1"]])
        dataset = export_dataset(path, {"sheet": "Any Sheet", "range": "A1:B2"})
        self.assertEqual(dataset["formula_evidence"]["formula_count"], 1)
        with self.assertRaisesRegex(ValueError, "require_no_formulas"):
            export_dataset(
                path,
                {"sheet": "Any Sheet", "range": "A1:B2"},
                formula_policy="require_no_formulas",
            )
        with self.assertRaisesRegex(ValueError, "without cached results"):
            export_dataset(
                path,
                {"sheet": "Any Sheet", "range": "A1:B2"},
                formula_policy="require_cached_results",
            )
