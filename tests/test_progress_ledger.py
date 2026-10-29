import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "progress_ledger.py"


class ProgressLedgerTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "source.xlsx"
        self.output = self.root / "output.xlsx"
        self.ledger = self.root / "progress.json"
        self.plan = self.root / "plan.json"
        self.evidence = self.root / "evidence.txt"
        self.evidence.write_text("test evidence", encoding="utf-8")
        with zipfile.ZipFile(self.source, "w") as workbook:
            workbook.writestr("xl/workbook.xml", '''<?xml version="1.0" encoding="UTF-8"?>
              <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheets><sheet name="Visible" state="visible"/><sheet name="Hidden" state="hidden"/></sheets>
              </workbook>''')
        self.output.write_bytes(b"revised workbook")
        self.request = hashlib.sha256(b"A4, preserve hidden").hexdigest()
        self.source_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()

    def tearDown(self):
        self.temp.cleanup()

    def call(self, *args, code=0):
        result = subprocess.run([sys.executable, str(SCRIPT), *map(str, args)],
                                capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(result.returncode, code, result.stdout + result.stderr)
        return json.loads(result.stdout or result.stderr)

    def initialize(self):
        return self.call("init", "--ledger", self.ledger, "--source", self.source,
                         "--output", self.output, "--request-fingerprint", self.request)

    def record(self, gate, sheet=None, state="passed"):
        command = ["record", "--ledger", self.ledger, "--gate", gate,
                   "--state", state, "--evidence", f"evidence for {gate}"]
        if state == "passed":
            command += ["--evidence-file", self.plan if gate == "plan" else self.evidence]
        if sheet:
            command += ["--sheet", sheet]
        return self.call(*command)

    def create_plan(self):
        self.plan.write_text(json.dumps({"schema_version": 1, "source_sha256": self.source_hash,
            "request_fingerprint": self.request, "sheets": [{"name": "Visible", "paper": "A4",
            "orientation": "landscape", "print_area": "$A$1:$C$8",
            "scale": {"mode": "zoom", "percent": 100},
            "margins_pt": {"left": 18, "right": 18, "top": 18, "bottom": 18}}]}), encoding="utf-8")

    def test_gate_requirements_idempotence_and_staleness(self):
        self.initialize()
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Visible", "--scope", "printable")
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Hidden", "--scope", "hidden")
        self.create_plan()
        failure = self.call("finalize", "--ledger", self.ledger, "--plan", self.plan, code=2)
        self.assertIn("gate", failure["error"])
        for gate in ("inventory", "baseline", "reopen", "content", "links"):
            self.record(gate, state="not_applicable" if gate == "links" else "passed")
        for gate in ("plan", "pdf", "visual"):
            self.record(gate, sheet="Visible")
        before_repeat = self.ledger.read_bytes()
        self.assertEqual(self.record("visual", sheet="Visible")["status"], "unchanged")
        self.assertEqual(self.ledger.read_bytes(), before_repeat)
        self.assertEqual(self.call("finalize", "--ledger", self.ledger, "--plan", self.plan)["status"], "verified")
        original_ledger = self.ledger.read_bytes()
        self.assertEqual(self.initialize()["status"], "already_exists")
        self.assertEqual(self.call("finalize", "--ledger", self.ledger, "--plan", self.plan)["unchanged"], True)
        self.assertEqual(self.ledger.read_bytes(), original_ledger)
        status = self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request)
        self.assertEqual(status["status"], "verified")
        self.assertEqual((status["gate_progress"]["passed"], status["gate_progress"]["required"]), (8, 8))
        another = hashlib.sha256(b"different request").hexdigest()
        self.assertEqual(self.call("status", "--ledger", self.ledger, "--request-fingerprint", another, code=1)["status"], "stale")
        self.output.write_bytes(b"changed after inspection")
        self.assertEqual(self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request, code=1)["status"], "stale")
        stale = self.call("finalize", "--ledger", self.ledger, "--plan", self.plan, code=2)
        self.assertIn("stale", stale["error"])
        for gate in ("reopen", "content", "links"):
            self.record(gate, state="not_applicable" if gate == "links" else "passed")
        for gate in ("pdf", "visual"):
            self.record(gate, sheet="Visible")
        self.call("finalize", "--ledger", self.ledger, "--plan", self.plan)
        self.evidence.write_text("changed evidence", encoding="utf-8")
        self.assertEqual(self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request, code=1)["status"], "stale")
        self.evidence.write_text("test evidence", encoding="utf-8")
        self.plan.write_text(self.plan.read_text(encoding="utf-8") + " ", encoding="utf-8")
        self.assertEqual(self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request, code=1)["status"], "stale")
        self.source.write_bytes(b"changed source")
        self.assertEqual(self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request, code=1)["status"], "stale")

    def test_reject_same_file_and_scope_change(self):
        self.call("init", "--ledger", self.ledger, "--source", self.source, "--output", self.source,
                  "--request-fingerprint", self.request, code=2)
        self.call("init", "--ledger", self.source, "--source", self.source, "--output", self.output,
                  "--request-fingerprint", self.request, code=2)
        self.call("init", "--ledger", self.output, "--source", self.source, "--output", self.output,
                  "--request-fingerprint", self.request, code=2)
        self.initialize()
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Hidden", "--scope", "hidden")
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Hidden", "--scope", "printable", code=2)

    def test_failed_gate_without_output_and_new_sheet_invalidates_inventory(self):
        self.initialize()
        self.output.unlink()
        self.record("reopen", state="failed")
        self.record("inventory")
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Visible", "--scope", "printable")
        ledger = json.loads(self.ledger.read_text(encoding="utf-8"))
        self.assertNotIn("inventory", ledger["gates"])
        self.assertEqual(ledger["gates"]["reopen"]["state"], "failed")

    def test_catalog_prevents_omitted_or_invented_sheets(self):
        self.initialize()
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Invented", "--scope", "printable", code=2)
        self.call("add-sheet", "--ledger", self.ledger, "--sheet", "Visible", "--scope", "printable")
        status = self.call("status", "--ledger", self.ledger, "--request-fingerprint", self.request, code=1)
        self.assertEqual(status["unclassified_sheet_count"], 1)
        self.assertIsNone(status["gate_progress"]["percent"])
        self.create_plan()
        for gate in ("inventory", "baseline", "reopen", "content", "links"):
            self.record(gate, state="not_applicable" if gate == "links" else "passed")
        for gate in ("plan", "pdf", "visual"):
            self.record(gate, sheet="Visible")
        result = self.call("finalize", "--ledger", self.ledger, "--plan", self.plan, code=2)
        self.assertIn("not all source-catalog sheets", result["error"])


if __name__ == "__main__":
    unittest.main()
