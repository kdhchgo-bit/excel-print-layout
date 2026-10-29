import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest import mock

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "compare_excel_com.py"
sys.path.insert(0, str(SCRIPT.parent))
spec = importlib.util.spec_from_file_location("excel_com_under_test", SCRIPT)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class Count:
    Count = 2


class Range:
    def __init__(self, formula, values):
        self.Formula = formula
        self.Value2 = values


class Sheet:
    Name = "Visible"
    Type = -4167
    Visible = -1

    def __init__(self, formula, values):
        self.formula = formula
        self.values = values
        self.UsedRange = types.SimpleNamespace(Row=1, Column=1, Rows=Count(), Columns=Count())

    def Cells(self, row, column):
        return (row, column)

    def Range(self, start, end):
        return Range(self.formula, self.values)


class Sheets:
    Count = 1

    def __init__(self, sheet):
        self.sheet = sheet

    def __call__(self, index):
        return self.sheet


class Workbook:
    def __init__(self, sheet, link):
        self.sheet = sheet
        self.Sheets = Sheets(sheet)
        self.link = link
        self.closed = False

    def Worksheets(self, name):
        if name != self.sheet.Name:
            raise KeyError(name)
        return self.sheet

    def LinkSources(self, kind):
        return (self.link,) if kind == 1 else None

    def Close(self, SaveChanges=False):
        self.closed = True


class Workbooks:
    def __init__(self, source, output):
        self.works = {"source.xlsx": source, "output.xlsx": output}
        self.blank = Workbook(Sheet(None, None), "")
        self.calls = []

    def Add(self):
        return self.blank

    def Open(self, path, **kwargs):
        self.calls.append(kwargs)
        return self.works[Path(path).name]


class Excel:
    Hwnd = 99

    def __init__(self, source, output):
        self.Workbooks = Workbooks(source, output)
        self.quit = False

    def Quit(self):
        self.quit = True


class ExcelComparisonTest(unittest.TestCase):
    def run_compare(self, original_formula, revised_formula, link1, link2, max_cells=20):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source, output, report = root / "source.xlsx", root / "output.xlsx", root / "report.json"
            source.write_bytes(b"original")
            output.write_bytes(b"revised")
            values = ((1, 2), (3, 4))
            fake = Excel(Workbook(Sheet(original_formula, values), link1),
                         Workbook(Sheet(revised_formula, values), link2))
            client = types.ModuleType("win32com.client")
            client.DispatchEx = lambda _: fake
            package = types.ModuleType("win32com")
            package.client = client
            process = types.ModuleType("win32process")
            process.GetWindowThreadProcessId = lambda _: (1, 222)
            argv = [str(SCRIPT), str(source), str(output), "--output", str(report), "--max-cells", str(max_cells)]
            with mock.patch.dict(sys.modules, {"win32com": package, "win32com.client": client, "win32process": process}), \
                    mock.patch.object(module.sys, "platform", "win32"), mock.patch.object(module.sys, "argv", argv):
                code = module.main()
            self.assertTrue(fake.quit)
            self.assertTrue(fake.Workbooks.blank.closed)
            self.assertTrue(all(call["UpdateLinks"] == 0 and call["ReadOnly"] for call in fake.Workbooks.calls))
            self.assertFalse(fake.EnableEvents)
            self.assertEqual(fake.AutomationSecurity, 3)
            self.assertEqual(fake.Calculation, -4135)
            return code, json.loads(report.read_text(encoding="utf-8"))

    def test_formula_and_link_match(self):
        matrix = (("=1+1", None), (3, 4))
        code, report = self.run_compare(matrix, matrix, "C:\\data\\source.xls", "C:\\data\\source.xls")
        self.assertEqual(code, 0)
        self.assertEqual(report["status"], "match_within_checked_fields")
        self.assertEqual(report["excel_pid"], 222)

    def test_formula_and_link_changes_fail(self):
        code, report = self.run_compare((("=1+1", None), (3, 4)), ((2, None), (3, 4)),
                                        "C:\\data\\a.xls", "C:\\data\\b.xls")
        self.assertEqual(code, 1)
        self.assertEqual(len(report["differences"]), 2)

    def test_too_many_cells_is_incomplete(self):
        matrix = (("=1+1", None), (3, 4))
        code, report = self.run_compare(matrix, matrix, "C:\\data\\a.xls", "C:\\data\\a.xls", max_cells=1)
        self.assertEqual(code, 2)
        self.assertEqual(report["status"], "incomplete")


if __name__ == "__main__":
    unittest.main()
