import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "audit_pdf.py"
READY = importlib.util.find_spec("fitz") is not None and importlib.util.find_spec("PIL") is not None


@unittest.skipUnless(READY, "PyMuPDF/Pillow not installed")
class PdfAuditTest(unittest.TestCase):
    def test_a4_blank_and_hash_detection_and_rendering(self):
        import fitz
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pdf, report, renders = root / "sheet.pdf", root / "audit.json", root / "renders"
            doc = fitz.open()
            page = doc.new_page(width=595.28, height=841.89)
            page.insert_text((40, 80), "Readable 123.45", fontsize=11)
            page = doc.new_page(width=595.28, height=841.89)
            page.insert_text((40, 80), "###", fontsize=10)
            doc.new_page(width=595.28, height=841.89)
            doc.save(pdf)
            doc.close()
            original_bytes = pdf.read_bytes()
            blocked = subprocess.run([sys.executable, str(SCRIPT), str(pdf), "--output", str(pdf)],
                                     capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(blocked.returncode, 2)
            self.assertEqual(pdf.read_bytes(), original_bytes)
            result = subprocess.run([sys.executable, str(SCRIPT), str(pdf), "--output", str(report),
                                     "--render-dir", str(renders)], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            pages = json.loads(report.read_text(encoding="utf-8"))["documents"][0]["pages"]
            self.assertTrue(all(p["a4"] for p in pages))
            self.assertFalse(pages[0]["hash_marks_detected"])
            self.assertTrue(pages[1]["hash_marks_detected"])
            self.assertTrue(pages[2]["blank_by_raster"])
            self.assertEqual(len(list(renders.glob("*.png"))), 3)


if __name__ == "__main__":
    unittest.main()
