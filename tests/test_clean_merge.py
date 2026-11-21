import csv
import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import clean_merge as cm  # noqa: E402

SAMPLES = sorted((ROOT / "sample_data").glob("*.csv"))


class CleanMergeTests(unittest.TestCase):
    def test_header_normalisation(self):
        self.assertEqual(cm.normalize_header(" E-mail "), "email")
        self.assertEqual(cm.normalize_header("Full Name"), "name")
        self.assertEqual(cm.normalize_header("Tel"), "phone")
        self.assertEqual(cm.normalize_header(""), "column")

    def test_value_cleaning(self):
        self.assertEqual(cm.clean_value("email", " Ivan@Example.COM "), "ivan@example.com")
        self.assertEqual(cm.clean_value("phone", "+380 (50) 123-45-67"), "+380501234567")
        self.assertEqual(cm.clean_value("phone", "050 222 33 44"), "0502223344")
        self.assertEqual(cm.clean_value("name", "  Ivan   Petrenko "), "Ivan Petrenko")

    def test_reads_different_encodings_and_delimiters(self):
        for path in SAMPLES:
            headers, rows = cm.read_csv(path)
            self.assertIn("email", headers, path.name)
            self.assertIn("name", headers, path.name)
        _, rows_c = cm.read_csv(ROOT / "sample_data" / "customers_c.csv")
        self.assertEqual(rows_c[0]["name"], "Петро Мельник")  # cp1251 decoded correctly

    def test_merge_by_email_removes_duplicates_and_empty_rows(self):
        headers, rows, stats = cm.merge(SAMPLES, key=["email"])
        emails = [r["email"] for r in rows if r.get("email")]
        self.assertEqual(len(emails), len(set(emails)))
        self.assertEqual(stats["empty_removed"], 2)       # one empty row in a, one in b
        self.assertEqual(stats["rows_in"], 10)
        self.assertEqual(stats["duplicates_removed"], 3)  # Ivan and Olena (file b), Maria (file c)
        self.assertEqual(stats["rows_out"], 5)
        self.assertIn("source_file", headers)

    def test_save_csv_and_xlsx(self):
        headers, rows, _ = cm.merge(SAMPLES, key=["email"])
        with tempfile.TemporaryDirectory() as tmp:
            out_csv = Path(tmp) / "r.csv"
            cm.save(out_csv, headers, rows)
            with open(out_csv, encoding="utf-8-sig", newline="") as f:
                self.assertEqual(len(list(csv.DictReader(f))), len(rows))
            out_xlsx = Path(tmp) / "r.xlsx"
            cm.save(out_xlsx, headers, rows)
            from openpyxl import load_workbook
            ws = load_workbook(out_xlsx).active
            self.assertEqual(ws.max_row, len(rows) + 1)


if __name__ == "__main__":
    unittest.main()
