#!/usr/bin/env python3
"""Format structured local records as common GB/T 7714-2015 references.

The tool is deliberately offline. It formats supplied fields and never looks up
missing authors, dates, pages, identifiers, or URLs.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

SUPPORTED = {"J", "M", "D", "C", "R", "S", "P", "EB/OL"}


def load_records(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="utf-8-sig")
    if path.suffix.lower() in {".jsonl", ".ndjson"}:
        values = [json.loads(line) for line in text.splitlines() if line.strip()]
    else:
        value = json.loads(text)
        values = value if isinstance(value, list) else value.get("references", [])
    if not isinstance(values, list) or not all(isinstance(item, dict) for item in values):
        raise ValueError("input must be a JSON array, an object with 'references', or JSONL records")
    return values


def required(record: dict[str, Any], key: str) -> str:
    value = str(record.get(key, "")).strip()
    if not value:
        raise ValueError(f"record {record.get('id', '?')}: missing {key}")
    return value


def authors(value: Any) -> str:
    if isinstance(value, str):
        names = [part.strip() for part in value.replace("；", ";").split(";") if part.strip()]
    elif isinstance(value, list):
        names = [str(part).strip() for part in value if str(part).strip()]
    else:
        names = []
    if not names:
        raise ValueError("each record needs at least one author")
    if len(names) <= 3:
        return ", ".join(names)
    # The caller appends the sentence-ending period after the author list.
    # Keep the English abbreviation period out here to avoid ``et al..``.
    suffix = "等" if any("\u4e00" <= char <= "\u9fff" for char in "".join(names)) else "et al"
    return ", ".join(names[:3]) + ", " + suffix


def optional(record: dict[str, Any], key: str) -> str:
    return str(record.get(key, "")).strip()


def format_record(record: dict[str, Any], number: int) -> str:
    kind = required(record, "type").upper()
    if kind not in SUPPORTED:
        raise ValueError(f"record {number}: unsupported type {kind}; use one of {sorted(SUPPORTED)}")
    creator = authors(record.get("authors"))
    title = required(record, "title")
    if kind == "J":
        source = required(record, "source")
        year = required(record, "year")
        volume = optional(record, "volume")
        issue = optional(record, "issue")
        pages = optional(record, "pages")
        volume_issue = (
            f", {volume}({issue})" if volume and issue
            else f", {volume}" if volume
            else f"({issue})" if issue
            else ""
        )
        page_text = f": {pages}" if pages else ""
        result = f"{creator}. {title}[J]. {source}, {year}{volume_issue}{page_text}."
    elif kind == "M":
        place = required(record, "place")
        publisher = required(record, "publisher")
        result = f"{creator}. {title}[M]. {place}: {publisher}, {required(record, 'year')}."
    elif kind == "D":
        place = required(record, "place")
        institution = required(record, "institution")
        result = f"{creator}. {title}[D]. {place}: {institution}, {required(record, 'year')}."
    elif kind == "C":
        collection = required(record, "collection")
        place = required(record, "place")
        publisher = required(record, "publisher")
        editor = optional(record, "editor")
        editor_text = f"//{editor}. " if editor else "//"
        pages = f": {optional(record, 'pages')}" if optional(record, "pages") else ""
        result = f"{creator}. {title}[C]{editor_text}{collection}. {place}: {publisher}, {required(record, 'year')}{pages}."
    elif kind == "EB/OL":
        url = required(record, "url")
        access = required(record, "accessed")
        published = optional(record, "published")
        date_part = f". ({published})[{access}]" if published else f"[{access}]"
        result = f"{creator}. {title}[EB/OL]{date_part}. {url}."
    else:
        source_key = "standard" if kind == "S" else "patent" if kind == "P" else "source"
        source = optional(record, source_key)
        result = f"{creator}. {title}[{kind}]."
        if source:
            result = f"{result} {source}."
        year = optional(record, "year")
        if year:
            result = f"{result[:-1]}, {year}."
    doi = optional(record, "doi")
    if doi:
        result = f"{result[:-1]}. DOI: {doi}."
    return f"[{number}] {result}"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="local JSON or JSONL records")
    parser.add_argument("--out", type=Path, help="Markdown output; defaults to stdout")
    args = parser.parse_args()
    try:
        lines = [format_record(record, index) for index, record in enumerate(load_records(args.input), 1)]
        output = "\n".join(lines) + ("\n" if lines else "")
        if args.out:
            args.out.write_text(output, encoding="utf-8")
        else:
            sys.stdout.write(output)
        return 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
