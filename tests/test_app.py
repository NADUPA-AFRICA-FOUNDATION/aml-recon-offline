"""Run with: python -m unittest discover -s tests -v."""
import io
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]


class Upload(io.BytesIO):
    def __init__(self, contents, name="training_2026.csv"):
        super().__init__(contents)
        self.name = name
        self.size = len(contents)


def export(user="001", missing_name=False):
    return Upload(pd.DataFrame([{
        "User ID": user, "Username": f"{user}@safaricom.co.ke",
        "Full Name": "" if missing_name else "Example Learner",
        "Assignment ID": "AML-2026", "Assignment Title": "AML refresher",
        "Assignment Begin": "2026-01-01", "Assignment End": "2026-12-31",
        "Completion Status": "Completed", "Result Status": "Passed",
    }]).to_csv(index=False).encode())


class StudioTests(unittest.TestCase):
    def test_empty_screen(self):
        app = AppTest.from_file(str(ROOT / "app.py")).run()
        self.assertFalse(app.exception)
        self.assertTrue(app.button[0].disabled)
        self.assertTrue(any("<h1>Training reconciliation</h1>" in item.value for item in app.markdown))
        self.assertTrue(any("Your results will appear here" in item.value for item in app.markdown))

    def test_reconcile_export_and_invalidate_changed_rules(self):
        # Same filename must not silently overwrite a separate uploaded export.
        with patch("streamlit.file_uploader", return_value=[export(), export("002")]):
            app = AppTest.from_file(str(ROOT / "app.py")).run()
            app.button[0].click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertEqual(app.session_state.summary["assignments"], 2)
            self.assertEqual(app.session_state.summary["participants"], 2)
            workbook = pd.ExcelFile(io.BytesIO(app.session_state.result_bytes))
            self.assertEqual(len(workbook.sheet_names), 7)
            self.assertEqual(len(pd.read_excel(workbook, "All Records")), 2)
            self.assertEqual(app.session_state.validation_errors, [])
            app.text_area(key="internal_domains").set_value("example.org").run()
            self.assertIsNone(app.session_state.result_bytes)
            self.assertIsNone(app.session_state.summary)

    def test_exception_search_and_processing_error(self):
        with patch("streamlit.file_uploader", return_value=[export(missing_name=True)]):
            app = AppTest.from_file(str(ROOT / "app.py")).run()
            app.button[0].click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state.summary["exceptions"], 1)
            app.text_input[0].set_value("does-not-exist").run()
            self.assertTrue(any("Showing 0 of 1" in item.value for item in app.caption))
        with patch("streamlit.file_uploader", return_value=[Upload(b"irrelevant\n\n")]):
            app = AppTest.from_file(str(ROOT / "app.py")).run()
            app.button[0].click().run(timeout=30)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertIsNone(app.session_state.result_bytes)

    def test_invalid_domains_block_run_and_recover(self):
        with patch("streamlit.file_uploader", return_value=[export()]):
            app = AppTest.from_file(str(ROOT / "app.py")).run()
            app.text_area(key="internal_domains").set_value("https://example.org").run()
            self.assertFalse(app.exception)
            self.assertTrue(app.button[0].disabled)
            self.assertTrue(any("Some domains are invalid" in item.value for item in app.error))
            app.text_area(key="internal_domains").set_value("@safaricom.co.ke; partner.org").run()
            self.assertFalse(app.button[0].disabled)
            app.button[0].click().run(timeout=30)
            self.assertEqual(app.session_state.summary["internal"], 1)
            original = app.session_state.result_bytes
            app.toggle(key="increased_contrast").set_value(True).run()
            self.assertEqual(app.session_state.result_bytes, original)
            self.assertFalse(app.exception)

    def test_filter_recovery_and_text_details_preserve_workbook(self):
        with patch("streamlit.file_uploader", return_value=[export(missing_name=True)]):
            app = AppTest.from_file(str(ROOT / "app.py")).run()
            app.button[0].click().run(timeout=30)
            original = app.session_state.result_bytes
            app.selectbox(key="exception_issue").select("Possible repeats").run()
            self.assertTrue(any("No exceptions match" in item.value for item in app.info))
            next(button for button in app.button if button.label == "Clear filters").click().run()
            self.assertEqual(app.selectbox(key="exception_issue").value, "All exceptions")
            self.assertTrue(any("Not provided" == item.value for item in app.text))
            self.assertEqual(app.session_state.result_bytes, original)
            self.assertFalse(app.exception)


if __name__ == "__main__":
    unittest.main()
