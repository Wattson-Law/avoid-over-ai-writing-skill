#!/usr/bin/env python3
"""Convert a local Markdown directory to DOCX without network access."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from markdown_to_docx import MODES, convert


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--mode", choices=MODES, required=True)
    parser.add_argument("--pattern", default="*.md", help="input glob, default: *.md")
    parser.add_argument("--report-dir", type=Path, help="write one JSON report per output")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    source_dir = args.input_dir.resolve()
    out_dir = args.out_dir.resolve()
    if not source_dir.is_dir():
        parser.error(f"input directory does not exist: {source_dir}")
    files = sorted(path for path in source_dir.glob(args.pattern) if path.is_file())
    if not files:
        parser.error(f"no files matched {args.pattern!r} in {source_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)
    if args.report_dir:
        args.report_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict] = []
    failed = False
    for source in files:
        output = out_dir / f"{source.stem}.docx"
        if output.exists() and not args.overwrite:
            result = {"ok": False, "input": str(source), "error": f"output exists: {output}"}
            failed = True
        else:
            try:
                result = convert(source.read_text(encoding="utf-8-sig"), output, args.mode, source.parent)
                result["input"] = str(source)
            except Exception as exc:  # keep processing other files and report every failure
                result = {"ok": False, "input": str(source), "output": str(output), "error": f"{type(exc).__name__}: {exc}"}
                failed = True
        results.append(result)
        if args.report_dir:
            (args.report_dir / f"{source.stem}.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False))
    summary = {"ok": not failed, "count": len(results), "failed": sum(not item.get("ok") for item in results)}
    print(json.dumps(summary, ensure_ascii=False))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
