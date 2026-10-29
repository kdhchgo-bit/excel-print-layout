#!/usr/bin/env python3
"""Read-only Excel COM formula/value/link comparison in a private instance."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
from safe_output import ensure_not_input, write_bytes

XL_WORKSHEET = -4167
XL_CALC_MANUAL = -4135
MSO_FORCE_DISABLE = 3


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def matrix_hash(value) -> str:
    data = json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def flatten_links(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (tuple, list)):
        return [part for item in value for part in flatten_links(item)]
    text = str(value)
    # Normalize Windows paths only; do not assume URLs are case-insensitive.
    if (len(text) > 1 and text[1] == ":") or text.startswith("\\\\"):
        text = os.path.normcase(os.path.normpath(text))
    return [text]


def read_links(workbook) -> dict:
    result = {}
    for kind, label in ((1, "excel"), (2, "ole")):
        try:
            result[label] = {"targets": sorted(flatten_links(workbook.LinkSources(kind)))}
        except Exception as exc:
            result[label] = {"error": f"{type(exc).__name__}: {exc}"}
    return result


def open_readonly(excel, path: Path):
    return excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True,
                                IgnoreReadOnlyRecommended=True, AddToMru=False)


def metadata(excel, path: Path) -> dict:
    workbook = open_readonly(excel, path)
    try:
        report = {"links": read_links(workbook), "sheets": []}
        for index in range(1, workbook.Sheets.Count + 1):
            sheet = workbook.Sheets(index)
            item = {"name": str(sheet.Name), "type": int(sheet.Type), "visible": int(sheet.Visible)}
            if item["type"] == XL_WORKSHEET:
                used = sheet.UsedRange
                item["bounds"] = [int(used.Row), int(used.Column),
                                  int(used.Row + used.Rows.Count - 1),
                                  int(used.Column + used.Columns.Count - 1)]
            report["sheets"].append(item)
        return report
    finally:
        workbook.Close(SaveChanges=False)


def matrices(excel, path: Path, bounds: dict[str, list[int]], max_cells: int) -> dict:
    workbook = open_readonly(excel, path)
    results = {}
    try:
        for name, (r1, c1, r2, c2) in bounds.items():
            count = (r2 - r1 + 1) * (c2 - c1 + 1)
            if count > max_cells:
                results[name] = {"skipped": f"{count} cells exceeds --max-cells"}
                continue
            sheet = workbook.Worksheets(name)
            cells = sheet.Range(sheet.Cells(r1, c1), sheet.Cells(r2, c2))
            results[name] = {"bounds": [r1, c1, r2, c2], "cell_count": count,
                             "formula_sha256": matrix_hash(cells.Formula),
                             "value2_sha256": matrix_hash(cells.Value2)}
        return results
    finally:
        workbook.Close(SaveChanges=False)


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("original", type=Path)
    p.add_argument("revised", type=Path)
    p.add_argument("--output", type=Path, help="write JSON report instead of stdout")
    p.add_argument("--overwrite", action="store_true", help="replace a different existing report")
    p.add_argument("--max-cells", type=int, default=500_000,
                   help="maximum cells per worksheet for COM array comparison")
    args = p.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    excel = None
    blank = None
    try:
        if sys.platform != "win32":
            raise RuntimeError("native Excel COM comparison requires Windows")
        if args.max_cells < 1 or args.original.suffix.lower() != args.revised.suffix.lower():
            raise ValueError("use matching workbook extensions and a positive --max-cells")
        if os.path.samefile(args.original, args.revised):
            raise ValueError("original and revised must be different files")
        if args.output:
            ensure_not_input(args.output, [args.original, args.revised])
        import win32com.client
        import win32process
        original_hash, revised_hash = sha(args.original), sha(args.revised)
        excel = win32com.client.DispatchEx("Excel.Application")
        pid = int(win32process.GetWindowThreadProcessId(excel.Hwnd)[1])
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = MSO_FORCE_DISABLE
        blank = excel.Workbooks.Add()
        excel.Calculation = XL_CALC_MANUAL
        excel.CalculateBeforeSave = False
        before, after = metadata(excel, args.original), metadata(excel, args.revised)
        old = {item["name"]: item for item in before["sheets"]}
        new = {item["name"]: item for item in after["sheets"]}
        difference = []
        incomplete = []
        if [x["name"] for x in before["sheets"]] != [x["name"] for x in after["sheets"]]:
            difference.append("sheet names or order differ")
        bounds = {}
        for name in old.keys() & new.keys():
            a, b = old[name], new[name]
            if a["type"] != b["type"] or a["visible"] != b["visible"]:
                difference.append(f"sheet type or visibility differs: {name}")
            if a["type"] == b["type"] == XL_WORKSHEET:
                x, y = a["bounds"], b["bounds"]
                bounds[name] = [min(x[0], y[0]), min(x[1], y[1]), max(x[2], y[2]), max(x[3], y[3])]
            else:
                incomplete.append(f"non-worksheet content not checked: {name}")
        for label in ("excel", "ole"):
            a, b = before["links"][label], after["links"][label]
            if "error" in a or "error" in b:
                incomplete.append(f"could not read {label} links")
            elif a["targets"] != b["targets"]:
                difference.append(f"{label} link targets differ")
        a_matrices = matrices(excel, args.original, bounds, args.max_cells)
        b_matrices = matrices(excel, args.revised, bounds, args.max_cells)
        for name in bounds:
            if "skipped" in a_matrices[name] or "skipped" in b_matrices[name]:
                incomplete.append(f"formula/value comparison skipped: {name}")
            elif a_matrices[name] != b_matrices[name]:
                difference.append(f"formula or value matrix differs: {name}")
        if sha(args.original) != original_hash or sha(args.revised) != revised_hash:
            incomplete.append("file changed during comparison")
        status = "different" if difference else ("incomplete" if incomplete else "match_within_checked_fields")
        report = {"status": status, "excel_pid": pid, "original_sha256": original_hash,
                  "revised_sha256": revised_hash, "original_metadata": before,
                  "revised_metadata": after, "sheets": {"original": a_matrices, "revised": b_matrices},
                  "differences": difference, "incomplete": incomplete,
                  "limitations": ["hidden row/column states and merges require structural comparison",
                                  "visual layout, external-data refresh and drawings are not proven",
                                  "read-only comparison does not recalculate formulas"]}
        text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            write_bytes(args.output, text.encode("utf-8"), overwrite=args.overwrite)
        else:
            print(text, end="")
        return {"match_within_checked_fields": 0, "different": 1, "incomplete": 2}[status]
    except Exception as exc:
        print(json.dumps({"status": "error", "error": f"{type(exc).__name__}: {exc}"},
                         ensure_ascii=False), file=sys.stderr)
        return 2
    finally:
        if blank is not None:
            try:
                blank.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(main())
