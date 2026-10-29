#!/usr/bin/env python3
"""Single-writer, per-workbook progress ledger; it never edits a workbook."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import tempfile
from datetime import datetime, timezone
import xml.etree.ElementTree as ET
import zipfile

BOOK_GATES = ("inventory", "baseline", "reopen", "content", "links")
SHEET_GATES = ("plan", "pdf", "visual")
SCOPES = ("printable", "hidden", "empty", "excluded")
STATES = ("passed", "failed", "not_applicable")
OUTPUT_GATES = {"reopen", "content", "links", "pdf", "visual"}
HEX = re.compile(r"^[0-9a-f]{64}$")
XML_FORMATS = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def absolute(value: str) -> str:
    return str(Path(value).expanduser().resolve())


def same_path(a: str, b: str) -> bool:
    a_path, b_path = Path(a), Path(b)
    if a_path.exists() and b_path.exists():
        return os.path.samefile(a_path, b_path)
    return os.path.normcase(absolute(a)) == os.path.normcase(absolute(b))


def sheet_catalog(source: Path, supplied: str | None) -> tuple[list[dict], str]:
    """Capture source sheet names. XLS chart/macro sheets also need native Excel review."""
    if supplied:
        raw = json.loads(Path(absolute(supplied)).read_text(encoding="utf-8"))
        sheets = raw.get("sheets")
        method = "user-supplied catalog; validate against native Excel"
    elif source.suffix.lower() in XML_FORMATS:
        with zipfile.ZipFile(source) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
        namespace = root.tag.split("}", 1)[0] + "}" if root.tag.startswith("{") else ""
        container = root.find(namespace + "sheets")
        if container is None:
            raise ValueError("workbook XML does not contain a sheet catalog")
        sheets = [{"name": item.get("name"), "visibility": item.get("state", "visible")}
                  for item in container.findall(namespace + "sheet")]
        method = "OOXML workbook sheet catalog"
    elif source.suffix.lower() == ".xls":
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError(".xls sheet inventory requires xlrd or --catalog") from exc
        book = xlrd.open_workbook(str(source), on_demand=True)
        try:
            states = {0: "visible", 1: "hidden", 2: "veryHidden"}
            sheets = [{"name": name, "visibility": states.get(book.sheet_by_name(name).visibility, "unknown")}
                      for name in book.sheet_names()]
        finally:
            book.release_resources()
        method = "xlrd worksheet catalog; native Excel must also check non-worksheet sheets"
    else:
        raise ValueError("unsupported source catalog format; provide --catalog from native Excel")
    if not isinstance(sheets, list) or not sheets:
        raise ValueError("sheet catalog must contain a nonempty sheets list")
    names = [item.get("name") for item in sheets if isinstance(item, dict)]
    if (len(names) != len(sheets) or any(not isinstance(name, str) or not name for name in names)
            or len(set(names)) != len(names)):
        raise ValueError("sheet catalog names must be unique nonempty strings")
    allowed_states = {"visible", "hidden", "veryHidden", "unknown"}
    if any(item.get("visibility") not in allowed_states for item in sheets):
        raise ValueError("sheet visibility must be visible, hidden, veryHidden, or unknown")
    return [{"name": item["name"], "visibility": item["visibility"]} for item in sheets], method


def load(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1:
        raise ValueError("unsupported ledger schema")
    return data


def write(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent,
                                     prefix=".ledger-", delete=False) as stream:
        temp = Path(stream.name)
        stream.write(text)
    try:
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def init(args: argparse.Namespace, path: Path) -> dict:
    source, output = absolute(args.source), absolute(args.output)
    if same_path(source, output):
        raise ValueError("source and output must be different files")
    if same_path(str(path), source) or same_path(str(path), output):
        raise ValueError("ledger must not be the source or output workbook")
    if Path(source).suffix.lower() != Path(output).suffix.lower():
        raise ValueError("source and output must retain the same extension")
    if not HEX.fullmatch(args.request_fingerprint):
        raise ValueError("request fingerprint must be a lowercase SHA-256 digest")
    source_hash = digest(Path(source))
    if path.exists():
        data = load(path)
        if (data["source"] == source and data["output"] == output
                and data["source_sha256"] == source_hash
                and data["request_fingerprint"] == args.request_fingerprint):
            return {"status": "already_exists", "ledger": str(path)}
        raise ValueError("existing ledger differs; create a new ledger rather than overwrite it")
    expected_sheets, catalog_method = sheet_catalog(Path(source), args.catalog)
    data = {"schema_version": 1, "created_utc": utc_now(), "source": source,
            "source_sha256": source_hash, "output": output,
            "request_fingerprint": args.request_fingerprint,
            "expected_sheets": expected_sheets, "catalog_method": catalog_method,
            "gates": {}, "sheets": {}, "final": None}
    write(path, data)
    return {"status": "initialized", "ledger": str(path), "source_sha256": source_hash}


def add_sheet(args: argparse.Namespace, path: Path) -> dict:
    data = load(path)
    if args.sheet not in {item["name"] for item in data["expected_sheets"]}:
        raise ValueError("sheet is not in the source catalog")
    prior = data["sheets"].get(args.sheet)
    if prior:
        if prior["scope"] != args.scope:
            raise ValueError("sheet already has a different scope; review the ledger")
        return {"status": "already_exists", "sheet": args.sheet}
    data["sheets"][args.sheet] = {"scope": args.scope, "gates": {}}
    # The old inventory evidence cannot cover a sheet added afterwards.
    data["gates"].pop("inventory", None)
    data["final"] = None
    write(path, data)
    return {"status": "added", "sheet": args.sheet, "scope": args.scope}


def record(args: argparse.Namespace, path: Path) -> dict:
    data = load(path)
    if not args.evidence.strip():
        raise ValueError("evidence must describe an actual file, check, or finding")
    if args.sheet:
        if args.sheet not in data["sheets"]:
            raise ValueError("sheet must be added before recording a gate")
        if data["sheets"][args.sheet]["scope"] != "printable":
            raise ValueError("per-sheet print gates apply only to printable sheets")
        if args.gate not in SHEET_GATES:
            raise ValueError("invalid per-sheet gate")
        gates = data["sheets"][args.sheet]["gates"]
    else:
        if args.gate not in BOOK_GATES:
            raise ValueError("invalid workbook gate")
        gates = data["gates"]
    if args.state == "not_applicable" and (args.sheet or args.gate != "links"):
        raise ValueError("not_applicable is allowed only for the workbook links gate")
    current = {"state": args.state, "evidence": args.evidence}
    evidence_files = []
    for item in args.evidence_file:
        evidence_path = Path(absolute(item))
        if not evidence_path.is_file():
            raise ValueError(f"evidence file does not exist: {evidence_path}")
        if same_path(str(evidence_path), str(path)):
            raise ValueError("the ledger cannot be its own evidence file")
        evidence_files.append({"path": str(evidence_path), "sha256": digest(evidence_path)})
    if args.state == "passed" and not evidence_files:
        raise ValueError("a passed gate requires at least one --evidence-file")
    current["evidence_files"] = sorted(evidence_files, key=lambda item: item["path"])
    if args.gate in OUTPUT_GATES:
        output = Path(data["output"])
        if not output.is_file() and args.state != "failed":
            raise ValueError("output must exist before recording this gate")
        current["output_sha256"] = digest(output) if output.is_file() else None
    previous = gates.get(args.gate)
    if previous and all(previous.get(k) == v for k, v in current.items()):
        return {"status": "unchanged", "gate": args.gate, "sheet": args.sheet}
    current["recorded_utc"] = utc_now()
    gates[args.gate] = current
    data["final"] = None
    write(path, data)
    return {"status": "recorded", "gate": args.gate, "sheet": args.sheet}


def valid_gate(gate: dict | None, name: str, output_hash: str) -> bool:
    if not gate:
        return False
    allowed = {"passed", "not_applicable"} if name == "links" else {"passed"}
    files = gate.get("evidence_files", [])
    evidence_current = all(Path(item["path"]).is_file() and digest(Path(item["path"])) == item["sha256"]
                           for item in files)
    return (gate.get("state") in allowed and bool(gate.get("evidence", "").strip())
            and (gate.get("state") != "passed" or bool(files)) and evidence_current
            and (name not in OUTPUT_GATES or gate.get("output_sha256") == output_hash))


def plan_errors(data: dict, plan: dict) -> list[str]:
    errors = []
    if plan.get("schema_version") != 1:
        errors.append("plan schema_version must be 1")
    if plan.get("source_sha256") != data["source_sha256"]:
        errors.append("plan source hash does not match the ledger")
    if plan.get("request_fingerprint") != data["request_fingerprint"]:
        errors.append("plan request fingerprint does not match the ledger")
    sheets = plan.get("sheets")
    if not isinstance(sheets, list):
        return errors + ["plan sheets must be a list"]
    names = [s.get("name") for s in sheets if isinstance(s, dict)]
    printable = {name for name, item in data["sheets"].items() if item["scope"] == "printable"}
    if len(names) != len(sheets) or len(names) != len(set(names)) or set(names) != printable:
        errors.append("plan sheet names must match printable sheets exactly")
    for sheet in sheets:
        if not isinstance(sheet, dict):
            continue
        for key in ("name", "paper", "orientation", "print_area", "scale", "margins_pt"):
            if key not in sheet or sheet[key] in (None, ""):
                errors.append(f"sheet {sheet.get('name')!r} lacks {key}")
        scale = sheet.get("scale")
        if sheet.get("orientation") not in ("portrait", "landscape"):
            errors.append(f"sheet {sheet.get('name')!r} needs portrait or landscape orientation")
        margins = sheet.get("margins_pt")
        if (not isinstance(margins, dict) or any(
                not isinstance(margins.get(k), (int, float)) or isinstance(margins.get(k), bool)
                or margins[k] < 0 for k in ("top", "bottom", "left", "right"))):
            errors.append(f"sheet {sheet.get('name')!r} needs four nonnegative margins in points")
        if not isinstance(scale, dict) or scale.get("mode") not in ("zoom", "fit"):
            errors.append(f"sheet {sheet.get('name')!r} needs an explicit zoom or fit scale")
        elif scale["mode"] == "zoom" and not (isinstance(scale.get("percent"), int)
                                                 and 10 <= scale["percent"] <= 400):
            errors.append(f"sheet {sheet.get('name')!r} has an invalid zoom percentage")
        elif scale["mode"] == "fit":
            fits = [scale.get("wide"), scale.get("tall")]
            if (not all(key in scale for key in ("wide", "tall")) or
                    any(x is not None and x is not False and
                        (isinstance(x, bool) or not isinstance(x, int) or x < 1) for x in fits) or
                    all(x is None or x is False for x in fits)):
                errors.append(f"sheet {sheet.get('name')!r} needs valid explicit fit width and height")
        for field in ("visible_column_widths", "visible_row_heights_pt"):
            dims = sheet.get(field, {})
            if (not isinstance(dims, dict) or any(isinstance(v, bool) or not isinstance(v, (int, float))
                                                  or v <= 0 for v in dims.values())):
                errors.append(f"sheet {sheet.get('name')!r} has invalid {field}")
        for field in ("horizontal_breaks_before", "vertical_breaks_before", "border_repairs"):
            if not isinstance(sheet.get(field, []), list):
                errors.append(f"sheet {sheet.get('name')!r} needs a list for {field}")
    return errors


def finalize(args: argparse.Namespace, path: Path) -> dict:
    data = load(path)
    source, output, plan_path = Path(data["source"]), Path(data["output"]), Path(absolute(args.plan))
    if same_path(str(source), str(output)):
        raise ValueError("source and output now refer to the same file")
    if digest(source) != data["source_sha256"]:
        raise ValueError("source changed; start a new baseline")
    if not output.is_file():
        raise ValueError("output does not exist")
    output_hash = digest(output)
    plan_hash = digest(plan_path)
    errors = plan_errors(data, json.loads(plan_path.read_text(encoding="utf-8")))
    if not data["sheets"]:
        errors.append("sheet inventory is empty")
    if set(data["sheets"]) != {item["name"] for item in data["expected_sheets"]}:
        errors.append("not all source-catalog sheets are classified")
    for gate in BOOK_GATES:
        if not valid_gate(data["gates"].get(gate), gate, output_hash):
            errors.append(f"workbook gate incomplete or stale: {gate}")
    for name, sheet in data["sheets"].items():
        if sheet["scope"] == "printable":
            for gate in SHEET_GATES:
                if not valid_gate(sheet["gates"].get(gate), gate, output_hash):
                    errors.append(f"sheet {name!r} gate incomplete or stale: {gate}")
            plan_gate = sheet["gates"].get("plan", {})
            if not any(item.get("path") == str(plan_path) and item.get("sha256") == plan_hash
                       for item in plan_gate.get("evidence_files", [])):
                errors.append(f"sheet {name!r} plan evidence is not the final plan")
    if errors:
        raise ValueError("; ".join(errors))
    final = {"output_sha256": output_hash, "plan": str(plan_path), "plan_sha256": plan_hash}
    previous = data.get("final")
    if previous and all(previous.get(k) == v for k, v in final.items()):
        return {"status": "verified", "output_sha256": output_hash, "unchanged": True}
    final["verified_utc"] = utc_now()
    data["final"] = final
    write(path, data)
    return {"status": "verified", "output_sha256": output_hash}


def status(args: argparse.Namespace, path: Path) -> dict:
    data = load(path)
    reasons = []
    if not HEX.fullmatch(args.request_fingerprint):
        raise ValueError("request fingerprint must be a lowercase SHA-256 digest")
    if args.request_fingerprint != data["request_fingerprint"]:
        reasons.append("current request fingerprint differs")
    final = data.get("final")
    if not final:
        reasons.append("not finalized")
    for key, expected in (("source", data["source_sha256"]),
                          ("output", final.get("output_sha256") if final else None),
                          ("plan", final.get("plan_sha256") if final else None)):
        candidate = Path(final["plan"] if key == "plan" and final else data.get(key, ""))
        if not expected:
            continue
        if not candidate.is_file() or digest(candidate) != expected:
            reasons.append(f"{key} missing or changed")
    if final:
        all_gates = list(data["gates"].values())
        all_gates += [gate for sheet in data["sheets"].values() for gate in sheet["gates"].values()]
        for gate in all_gates:
            for evidence in gate.get("evidence_files", []):
                evidence_path = Path(evidence["path"])
                if not evidence_path.is_file() or digest(evidence_path) != evidence["sha256"]:
                    reasons.append(f"evidence missing or changed: {evidence_path}")
    state = "verified" if final and not reasons else ("stale" if final else "incomplete")
    printable = sum(item["scope"] == "printable" for item in data["sheets"].values())
    unclassified = len({item["name"] for item in data["expected_sheets"]} - set(data["sheets"]))
    output_path = Path(data["output"])
    current_output = digest(output_path) if output_path.is_file() else ""
    passed = sum(valid_gate(data["gates"].get(gate), gate, current_output) for gate in BOOK_GATES)
    passed += sum(valid_gate(item["gates"].get(gate), gate, current_output)
                  for item in data["sheets"].values() if item["scope"] == "printable" for gate in SHEET_GATES)
    required = len(BOOK_GATES) + printable * len(SHEET_GATES)
    return {"status": state, "reasons": reasons, "sheet_count": len(data["sheets"]),
            "printable_sheet_count": printable,
            "unclassified_sheet_count": unclassified,
            "gate_progress": {"passed": passed, "required": required,
                              "percent": round(100 * passed / required, 1) if not unclassified else None,
                              "note": "gate coverage, not elapsed-time or quality prediction; no percentage until all sheets are classified"},
            "workbook_gates": data["gates"],
            "sheets": data["sheets"]}


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest="command", required=True)
    start = sub.add_parser("init")
    start.add_argument("--ledger", required=True)
    start.add_argument("--source", required=True)
    start.add_argument("--output", required=True)
    start.add_argument("--request-fingerprint", required=True)
    start.add_argument("--catalog", help="JSON sheet catalog from native Excel when automatic inventory is unsupported")
    sheet = sub.add_parser("add-sheet")
    sheet.add_argument("--ledger", required=True)
    sheet.add_argument("--sheet", required=True)
    sheet.add_argument("--scope", choices=SCOPES, required=True)
    item = sub.add_parser("record")
    item.add_argument("--ledger", required=True)
    item.add_argument("--sheet")
    item.add_argument("--gate", choices=BOOK_GATES + SHEET_GATES, required=True)
    item.add_argument("--state", choices=STATES, required=True)
    item.add_argument("--evidence", required=True)
    item.add_argument("--evidence-file", action="append", default=[],
                      help="evidence file to hash; repeat as needed (required for passed)")
    finish = sub.add_parser("finalize")
    finish.add_argument("--ledger", required=True)
    finish.add_argument("--plan", required=True)
    check = sub.add_parser("status")
    check.add_argument("--ledger", required=True)
    check.add_argument("--request-fingerprint", required=True,
                       help="SHA-256 of the current normalized request and scope")
    return p


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    args = parser().parse_args()
    path = Path(absolute(args.ledger))
    try:
        result = {"init": init, "add-sheet": add_sheet, "record": record,
                  "finalize": finalize, "status": status}[args.command](args, path)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("status") not in ("incomplete", "stale") else 1
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
