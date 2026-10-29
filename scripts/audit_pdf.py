#!/usr/bin/env python3
"""Triage Excel-exported PDFs and optionally render pages for human review."""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import unicodedata
from safe_output import ensure_not_input, write_bytes


def tokens(document) -> set[str]:
    result = set()
    for page in document:
        for word in page.get_text("words"):
            text = unicodedata.normalize("NFKC", str(word[4])).strip()
            if text:
                result.add(text)
    return result


def weighted_median(counter: Counter) -> float | None:
    total = sum(counter.values())
    if not total:
        return None
    seen = 0
    for size in sorted(counter):
        seen += counter[size]
        if seen * 2 >= total:
            return round(size, 2)
    return None


def audit(path: Path, render_dir: Path | None, dpi: int, small_font: float,
          overwrite: bool, report_path: Path | None) -> dict:
    import fitz
    from PIL import Image
    doc = fitz.open(path)
    try:
        if doc.needs_pass:
            raise ValueError("encrypted PDF requires a password")
        pages = []
        key = hashlib.sha256(str(path.resolve()).encode("utf-8")).hexdigest()[:8]
        for page in doc:
            rect = page.rect
            spans = [span for block in page.get_text("dict")["blocks"]
                     for line in block.get("lines", []) for span in line.get("spans", [])
                     if span.get("text", "").strip()]
            sizes = Counter()
            outside = []
            for span in spans:
                size = round(float(span["size"]), 2)
                sizes[size] += len(span.get("text", "").strip())
                bbox = fitz.Rect(span["bbox"])
                if bbox.x0 < -0.2 or bbox.y0 < -0.2 or bbox.x1 > rect.width + 0.2 or bbox.y1 > rect.height + 0.2:
                    outside.append({"text": span.get("text", "")[:50], "bbox_pt": list(bbox)})
            pixmap = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5), colorspace=fitz.csGRAY, alpha=False)
            image = Image.frombytes("L", (pixmap.width, pixmap.height), pixmap.samples)
            ink = image.point(lambda pixel: 255 if pixel < 235 else 0).getbbox()
            ink_bbox_pt = [round(i * 2, 1) for i in ink] if ink else None
            a4 = abs(min(rect.width, rect.height) - 595.28) < 3 and abs(max(rect.width, rect.height) - 841.89) < 3
            text = page.get_text("text")
            page_report = {"page": page.number + 1, "size_pt": [round(rect.width, 2), round(rect.height, 2)],
                           "a4": a4, "blank_by_raster": ink is None,
                           "ink_bbox_pt": ink_bbox_pt,
                           "ink_width_fraction": round((ink[2] - ink[0]) / pixmap.width, 3) if ink else 0,
                           "ink_height_fraction": round((ink[3] - ink[1]) / pixmap.height, 3) if ink else 0,
                           "hash_marks_detected": bool(re.search(r"#{3,}", text)),
                           "text_chars": sum(sizes.values()), "min_font_pt": min(sizes) if sizes else None,
                           "weighted_median_font_pt": weighted_median(sizes),
                           "chars_below_threshold": sum(count for size, count in sizes.items() if size < small_font),
                           "text_spans_outside_page": outside[:20]}
            if render_dir:
                render_dir.mkdir(parents=True, exist_ok=True)
                image_name = f"{path.stem}_{key}_p{page.number + 1:03d}.png"
                image_path = render_dir / image_name
                ensure_not_input(image_path, [path] + ([report_path] if report_path else []))
                image_bytes = page.get_pixmap(dpi=dpi, alpha=False).tobytes("png")
                write_bytes(image_path, image_bytes, overwrite=overwrite)
                page_report["render"] = str(image_path)
            pages.append(page_report)
        return {"path": str(path), "pages": pages}
    finally:
        doc.close()


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("paths", nargs="+", type=Path, help="PDF files or directories")
    p.add_argument("--recursive", action="store_true", help="include PDF files in child directories")
    p.add_argument("--output", type=Path, help="write the JSON report to this path")
    p.add_argument("--overwrite", action="store_true", help="replace different existing reports or renders")
    p.add_argument("--render-dir", type=Path, help="render every page to this evidence directory")
    p.add_argument("--dpi", type=int, default=150)
    p.add_argument("--small-font", type=float, default=7.0, help="report characters below this size")
    p.add_argument("--baseline", type=Path, help="compare distinct text tokens against one baseline PDF")
    args = p.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        import fitz
        from PIL import Image  # noqa: F401
        if not 50 <= args.dpi <= 600 or args.small_font <= 0:
            raise ValueError("use --dpi 50..600 and a positive --small-font")
        paths = []
        for candidate in args.paths:
            if candidate.is_file() and candidate.suffix.lower() == ".pdf":
                paths.append(candidate)
            elif candidate.is_dir():
                children = candidate.rglob("*") if args.recursive else candidate.iterdir()
                paths.extend(p for p in children if p.is_file() and p.suffix.lower() == ".pdf")
            else:
                raise ValueError(f"not a PDF or directory: {candidate}")
        paths = sorted(set(p.resolve() for p in paths))
        if not paths:
            raise ValueError("no PDFs found")
        if args.baseline and len(paths) != 1:
            raise ValueError("--baseline requires exactly one revised PDF")
        if args.baseline and os.path.samefile(args.baseline, paths[0]):
            raise ValueError("baseline and revised PDF must be different files")
        if args.output:
            ensure_not_input(args.output, paths + ([args.baseline] if args.baseline else []))
        reports = [audit(path, args.render_dir, args.dpi, args.small_font, args.overwrite, args.output)
                   for path in paths]
        report = {"documents": reports,
                  "limitations": ["raster and text checks do not prove cell borders, glyph clipping or readability",
                                  "text embedded in a raster image may not be extractable",
                                  "font size and token checks are review aids, not pass/fail proof"]}
        if args.baseline:
            with fitz.open(args.baseline) as baseline, fitz.open(paths[0]) as revised:
                missing = sorted(tokens(baseline) - tokens(revised))
            report["baseline_possible_missing_tokens"] = {"count": len(missing), "sample": missing[:30],
                                                          "note": "tokenization can differ; inspect manually"}
        text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            write_bytes(args.output, text.encode("utf-8"), overwrite=args.overwrite)
        else:
            print(text, end="")
        flagged = any(not page["a4"] or page["blank_by_raster"] or page["hash_marks_detected"]
                      or page["text_spans_outside_page"]
                      for doc in reports for page in doc["pages"])
        return 1 if flagged else 0
    except Exception as exc:
        print(json.dumps({"status": "error", "error": f"{type(exc).__name__}: {exc}"},
                         ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
