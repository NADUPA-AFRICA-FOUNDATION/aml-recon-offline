import json
import tempfile
import unittest
from datetime import time
from pathlib import Path

from openpyxl import Workbook

import pandas as pd

from recon_verifier import Verifier, apply_derived_checks, main, write_report


ROOT = Path(__file__).resolve().parents[1]


def save_book(path, sheet, rows, extra=None):
    workbook = Workbook()
    ws = workbook.active
    ws.title = sheet
    for row in rows:
        ws.append(row)
    for name, values in (extra or {}).items():
        tab = workbook.create_sheet(name)
        for row in values:
            tab.append(row)
    workbook.save(path)


class VerifierTests(unittest.TestCase):
    def test_aml_config_describes_delivered_workbooks(self):
        config = json.loads((ROOT / "recon_config_aml_training.json").read_text())
        self.assertEqual(config["raw"]["glob"], "Training_Progress_Report_*.xlsx")
        self.assertEqual(config["raw"]["sheet_name"], "Training Progress Report")
        self.assertEqual(config["raw"]["header_mode"], "grouped")
        self.assertEqual(config["recon"]["path"], "AML_Training_Full_Reconciliation_Reaudited_Final.xlsx")
        self.assertEqual(config["recon"]["header_row"], 3)
        self.assertEqual(config["recon"]["source_file_transform"]["repl"], "Training Progress Report_")
        self.assertEqual(config["field_map"]["Assignment Info::Assignement title"], "Assignment Title")
        self.assertEqual(config["summary_checks"][1]["data_end_row"], 8)
        self.assertEqual(config["cell_checks"][-1]["cell"], "I9")

    def fixture(self, root, mismatch=False):
        save_book(root / "raw_one.xlsx", "Data", [
            ["People", "People", "Training", "Training"],
            ["ID", "Name", "Status", "Assigned"],
            ["1", "Ada", "Complete", "01.05.2025"],
            ["2", "Lin", "Open", "02.05.2025"],
        ])
        save_book(root / "recon.xlsx", "Master", [
            ["ID", "Person", "Status", "Assigned", "Done", "Year", "Source"],
            ["1", "Wrong" if mismatch else "Ada", "Complete", "2025-05-01", 1, 2025, "raw_one.xlsx"],
            ["2", "Lin", "Open", "2025-05-02", 0, 2025, "raw_one.xlsx"],
        ], {"Summary": [["Status", "Rows"], ["Complete", 1], ["Open", 1]], "Dashboard": [[2]]})
        return {
            "raw": {"glob": "raw_*.xlsx", "sheet_name": "Data", "header_mode": "grouped", "group_header_row": 1, "field_header_row": 2, "data_start_row": 3},
            "recon": {"path": "recon.xlsx", "records_sheet": "Master", "header_row": 1, "data_start_row": 2, "source_file_column": "Source"},
            "join": {"mode": "keys", "raw_keys": ["People::ID"], "master_keys": ["ID"]},
            "field_map": {"People::Name": "Person", "Training::Status": "Status"},
            "date_field_map": {"Training::Assigned": "Assigned"}, "raw_date_format": "%d.%m.%Y",
            "derived_checks": [{"name": "Completion rule", "type": "value_equals", "source_col": "Status", "match_value": "Complete", "target_col": "Done"}, {"type": "year_of_date", "source_col": "Assigned", "target_col": "Year"}],
            "summary_checks": [{"name": "status", "sheet_name": "Summary", "header_row": 1, "data_start_row": 2, "group_by": ["Status"], "metrics": {"Rows": {"col": "ID", "agg": "count"}}}],
            "cell_checks": [{"name": "records", "sheet": "Dashboard", "cell": "A1", "metric": {"col": "ID", "agg": "count"}}],
        }

    def test_all_check_types_pass_and_report(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            results = Verifier(self.fixture(root), root).run()
            self.assertTrue(all(result.passed for result in results), [(r.name, r.detail, r.samples) for r in results])
            self.assertTrue(any(result.name == "Derived: Completion rule" for result in results))
            write_report(root / "report.md", results)
            self.assertIn("**Overall: PASS**", (root / "report.md").read_text())

    def test_mismatch_returns_one_and_writes_samples(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self.fixture(root, mismatch=True)
            (root / "config.json").write_text(json.dumps(config))
            self.assertEqual(main(["--config", str(root / "config.json"), "--report", str(root / "report.md")]), 1)
            report = (root / "report.md").read_text()
            self.assertIn("**Overall: FAIL**", report)
            self.assertIn("Wrong", report)

    def test_no_raw_files_is_configuration_error(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = {"raw": {"glob": "none*.xlsx"}, "recon": {}}
            (root / "config.json").write_text(json.dumps(config))
            self.assertEqual(main(["--config", str(root / "config.json")]), 2)

    def test_all_derived_rule_types_and_null_role_evidence(self):
        frame = pd.DataFrame({
            "Role": [None, "Learner", "Learner\nAML analyst"],
            "Status": ["Done", "Open", "done"],
            "When": ["2025-01-02", None, "2026-03-04"],
            "Seconds": [None, time(0, 0, 30), "00:01:00"],
        })
        derived = apply_derived_checks(frame, [
            {"type": "role_evidence", "source_col": "Role", "generic_marker": "Learner", "target_col": "Internal", "true_value": "Internal", "false_value": "External"},
            {"type": "value_in_set", "source_col": "Status", "match_values": ["Done"], "target_col": "In set"},
            {"type": "value_equals", "source_col": "Status", "match_value": "Open", "target_col": "Equal"},
            {"type": "year_of_date", "source_col": "When", "target_col": "Year"},
            {"type": "duration_bucket", "source_col": "Seconds", "min_seconds": 20, "max_seconds": 60, "target_col": "Bucket"},
        ])
        self.assertEqual(derived["Internal"].tolist(), ["External", "External", "Internal"])
        self.assertEqual(derived["In set"].tolist(), [1, 0, 1])
        self.assertEqual(derived["Equal"].tolist(), [0, 1, 0])
        self.assertTrue(pd.isna(derived["Year"].iloc[1]))
        self.assertEqual(derived["Bucket"].tolist(), [0, 1, 0])

    def test_cli_path_overrides(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self.fixture(root)
            config["raw"]["glob"] = "not-the-real-location/*.xlsx"
            config["recon"]["path"] = "not-the-real-workbook.xlsx"
            (root / "config.json").write_text(json.dumps(config))
            exit_code = main([
                "--config", str(root / "config.json"),
                "--raw-glob", str(root / "raw_*.xlsx"),
                "--recon-path", str(root / "recon.xlsx"),
            ])
            self.assertEqual(exit_code, 0)

    def test_key_join_catches_swapped_missing_and_extra_keys(self):
        verifier = Verifier({"join": {"mode": "keys", "raw_keys": ["ID"], "master_keys": ["ID"]}})
        verifier.raw = pd.DataFrame({"ID": ["raw-only"], "Value": [1]})
        verifier.master = pd.DataFrame({"ID": ["master-only"], "Value": [1]})
        verifier.align()
        result = next(result for result in verifier.results if result.name == "Join integrity")
        self.assertFalse(result.passed)
        self.assertIn("unmatched=2", result.detail)

    def test_positional_join_includes_master_only_sources(self):
        verifier = Verifier({
            "join": {"mode": "positional"},
            "recon": {"source_file_column": "Source"},
        })
        verifier.raw = pd.DataFrame({"__source_file__": ["raw.xlsx"], "Value": [1]})
        verifier.master = pd.DataFrame({
            "Source": ["raw.xlsx", "master-only.xlsx"],
            "Value": [1, 2],
        })
        aligned = verifier.align()
        self.assertEqual(len(aligned), 2)
        self.assertFalse(verifier.results[-1].passed)

    def test_invalid_dates_do_not_reconcile_as_two_missing_dates(self):
        verifier = Verifier({
            "date_field_map": {"Raw date": "Recon date"},
            "raw_date_format": "%d.%m.%Y",
        })
        aligned = pd.DataFrame({
            "Raw date__raw": ["not-a-date"],
            "Recon date__master": ["also-not-a-date"],
        })
        verifier.compare_fields(aligned)
        self.assertFalse(verifier.results[0].passed)
        self.assertEqual(len(verifier.results[0].samples), 1)
        self.assertEqual(verifier.results[0].samples[0]["raw"], "not-a-date")


if __name__ == "__main__":
    unittest.main()
