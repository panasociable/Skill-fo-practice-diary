import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "dnevnik-praktiki" / "scripts" / "fill_diary.py"
EXAMPLE = ROOT / "skills" / "dnevnik-praktiki" / "assets" / "data_example.json"


def load_fill_diary():
    spec = importlib.util.spec_from_file_location("fill_diary", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class FillDiaryTest(unittest.TestCase):
    def test_example_generates_docx_without_placeholders(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "diary.docx"
            result = subprocess.run(
                [sys.executable, str(SCRIPT), str(EXAMPLE), str(out)],
                cwd=ROOT,
                text=True,
                capture_output=True,
                check=False,
            )

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(out.exists())

            doc = Document(out)
            text = "\n".join(
                [p.text for p in doc.paragraphs]
                + [c.text for t in doc.tables for r in t.rows for c in r.cells]
            )
            self.assertNotRegex(text, re.compile(r"\{\{[^}]+}}"))

    def test_rejects_end_date_before_start_date(self):
        data = json.loads(EXAMPLE.read_text(encoding="utf-8"))
        data["end_date"] = "2026-06-01"

        with tempfile.TemporaryDirectory() as tmp:
            input_path = Path(tmp) / "data.json"
            output_path = Path(tmp) / "diary.docx"
            input_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

            result = subprocess.run(
                [sys.executable, str(SCRIPT), str(input_path), str(output_path)],
                cwd=ROOT,
                text=True,
                capture_output=True,
                check=False,
            )

            self.assertNotEqual(result.returncode, 0)
            self.assertIn("end_date не может быть раньше start_date", result.stderr)

    def test_short_practice_periods_do_not_reverse_dates(self):
        fill_diary = load_fill_diary()
        start = fill_diary.parse("2026-06-29", "start_date")
        periods = fill_diary.split_periods(start, start)

        self.assertEqual(len(periods), 8)
        self.assertEqual(set(periods), {"29.06-29.06"})


if __name__ == "__main__":
    unittest.main()
