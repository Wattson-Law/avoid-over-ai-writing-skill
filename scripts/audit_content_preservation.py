#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"

DIRECTION_RE = re.compile(
    r"不低于|不高于|不少于|不超过|至少|至多|大于|小于|高于|低于|"
    r"≥|≤|>|<|±"
)
NUMBER_RE = re.compile(
    r"(?<![A-Za-z0-9])"
    r"\d+(?:\.\d+)?"
    r"(?:\s*[-~至—]\s*\d+(?:\.\d+)?)?"
    r"\s*(?:%|％|万元|元|亿元|年|个月|月|日|天|小时|分钟|秒|毫秒|ms|s|"
    r"人|台|套|项|篇|个|次|家|份|章|页|m²|m2|㎡|mm|cm|m|km|GB|TB|核|G)?"
)
LATIN_RE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9+_.-]*(?![A-Za-z0-9])")
ORG_RE = re.compile(
    r"[\u4e00-\u9fffA-Za-z0-9]{2,30}(?:有限公司|集团|大学|研究院|实验室|研究中心)"
)
SEQUENCE_CELL_RE = re.compile(r"^(?:序号|\d+)$")
LIST_PREFIX_RE = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十百]+[章节部分]\s*|"
    r"[一二三四五六七八九十百]+[、.]\s*|"
    r"[（(]\d+[）)]\s*|"
    r"\d+[）)]\s*|"
    r"\d+[.、]\s+"
    r")"
)
HEADING_PREFIX_RE = re.compile(r"^\d+(?:\.\d+)+(?:[.、]\s*|\s*)")
TABLE_CAPTION_RE = re.compile(r"^表\s*\d+$")


def normalize(text: str) -> str:
    text = text.replace("\u3000", " ").replace("％", "%")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_token(text: str) -> str:
    return re.sub(r"\s+", "", normalize(text))


def visible_character_count(text: str) -> int:
    return len(re.sub(r"\s+", "", text))


def load_package(path: Path):
    with zipfile.ZipFile(path) as package:
        document_root = etree.fromstring(package.read("word/document.xml"))
        styles_root = etree.fromstring(package.read("word/styles.xml"))
    style_names = {}
    for style in styles_root.xpath("./w:style", namespaces=NS):
        names = style.xpath("./w:name/@w:val", namespaces=NS)
        style_id = style.get(W + "styleId")
        if style_id and names:
            style_names[style_id] = names[0].strip()
    toc_style_ids = {
        style_id
        for style_id, name in style_names.items()
        if name.lower().startswith("toc")
    }
    return document_root, style_names, toc_style_ids


def visible_text(element) -> str:
    parts = []
    for node in element.xpath(".//w:t", namespaces=NS):
        if any(ancestor.tag in {W + "del", W + "moveFrom"} for ancestor in node.iterancestors()):
            continue
        parts.append(node.text or "")
    return "".join(parts)


def paragraph_style(paragraph) -> str:
    values = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return values[0] if values else ""


def strip_structural_prefix(text: str, style_name: str) -> str:
    text = LIST_PREFIX_RE.sub("", text, count=1)
    normalized_style = style_name.strip().lower()
    if normalized_style.startswith("heading") or "标题" in normalized_style:
        text = HEADING_PREFIX_RE.sub("", text, count=1)
    return text


def body_text(root, style_names: dict[str, str], toc_style_ids: set[str]) -> str:
    values = []
    for paragraph in root.xpath(".//w:body//w:p", namespaces=NS):
        if any(ancestor.tag == W + "tbl" for ancestor in paragraph.iterancestors()):
            continue
        style_id = paragraph_style(paragraph)
        if style_id in toc_style_ids or style_id.lower().startswith("toc"):
            continue
        text = normalize(visible_text(paragraph))
        if text:
            if TABLE_CAPTION_RE.fullmatch(text):
                continue
            values.append(strip_structural_prefix(text, style_names.get(style_id, "")))
    return "\n".join(values)


def table_business_text(root) -> str:
    values = []
    for cell in root.xpath(".//w:body//w:tbl//w:tc", namespaces=NS):
        text = normalize(visible_text(cell))
        if text and not SEQUENCE_CELL_RE.fullmatch(text):
            values.append(text)
    return "\n".join(values)


def table_cells(root) -> Counter:
    cells = Counter()
    for cell in root.xpath(".//w:body//w:tbl//w:tc", namespaces=NS):
        text = normalize(visible_text(cell))
        if text and not SEQUENCE_CELL_RE.fullmatch(text):
            cells[text] += 1
    return cells


def protected_tokens(text: str) -> dict[str, Counter]:
    return {
        "directions": Counter(normalize_token(x) for x in DIRECTION_RE.findall(text)),
        "numbers": Counter(normalize_token(x) for x in NUMBER_RE.findall(text) if x.strip()),
        "latin_terms": Counter(LATIN_RE.findall(text)),
        "organizations": Counter(ORG_RE.findall(text)),
    }


def counter_diff(source: Counter, output: Counter) -> dict:
    return {
        "missing": dict(source - output),
        "added": dict(output - source),
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Audit protected content between a source and edited DOCX."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument(
        "--min-text-retention",
        type=float,
        default=0.85,
        help="Minimum output/source narrative character ratio (default: 0.85).",
    )
    args = parser.parse_args()
    if not 0 <= args.min_text_retention <= 1:
        parser.error("--min-text-retention must be between 0 and 1.")

    source_root, source_styles, source_toc_styles = load_package(args.source)
    output_root, output_styles, output_toc_styles = load_package(args.output)
    source_body_text = body_text(source_root, source_styles, source_toc_styles)
    output_body_text = body_text(output_root, output_styles, output_toc_styles)
    source_text = "\n".join(
        (
            source_body_text,
            table_business_text(source_root),
        )
    )
    output_text = "\n".join(
        (
            output_body_text,
            table_business_text(output_root),
        )
    )
    source_tokens = protected_tokens(source_text)
    output_tokens = protected_tokens(output_text)

    differences = {
        category: counter_diff(source_tokens[category], output_tokens[category])
        for category in source_tokens
    }
    differences["table_cells"] = counter_diff(
        table_cells(source_root), table_cells(output_root)
    )
    protected_content_ok = all(
        not detail[side]
        for detail in differences.values()
        for side in ("missing", "added")
    )
    source_characters = visible_character_count(source_body_text)
    output_characters = visible_character_count(output_body_text)
    retention_ratio = (
        output_characters / source_characters if source_characters else 1.0
    )
    text_retention_ok = retention_ratio >= args.min_text_retention
    ok = protected_content_ok and text_retention_ok
    report = {
        "ok": ok,
        "source": str(args.source.resolve()),
        "output": str(args.output.resolve()),
        "differences": differences,
        "text_retention": {
            "ok": text_retention_ok,
            "source_characters": source_characters,
            "output_characters": output_characters,
            "ratio": round(retention_ratio, 4),
            "minimum": args.min_text_retention,
            "note": (
                "Mechanical over-compression guardrail only; passing does not prove "
                "that every semantic unit was preserved."
            ),
        },
        "note": "Internal preservation audit only; do not publish as a user-facing issue list.",
    }

    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    if not ok:
        print("Protected-content audit failed.", file=sys.stderr)
        return 1
    print("Protected-content audit passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
