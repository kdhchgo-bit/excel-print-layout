import json
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]


class SkillStructureTest(unittest.TestCase):
    def test_skill_references_exist(self):
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertTrue(skill.startswith("---\nname: excel-print-layout\n"))
        for target in re.findall(r"\]\(([^)#]+)(?:#[^)]*)?\)", skill):
            if "://" not in target:
                self.assertTrue((ROOT / target).is_file(), target)

    def test_example_plan_is_valid_json_and_marked_as_example(self):
        plan = json.loads((ROOT / "assets" / "print-plan.example.json").read_text(encoding="utf-8"))
        self.assertEqual(plan["schema_version"], 1)
        self.assertEqual(plan["source_sha256"], "0" * 64)
        self.assertIn("replace-with", plan["request_fingerprint"])

    def test_bilingual_detailed_references_and_usage_exist(self):
        for name in ("architecture.ko.md", "architecture.en.md", "ledger-usage.md"):
            self.assertTrue((ROOT / "references" / name).is_file())
        for name in ("architecture.ko.md", "architecture.en.md"):
            text = (ROOT / "references" / name).read_text(encoding="utf-8")
            self.assertEqual(len(re.findall(r"^\| L\d{2} ", text, flags=re.MULTILINE)), 19)


if __name__ == "__main__":
    unittest.main()
