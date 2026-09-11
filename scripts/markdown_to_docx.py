#!/usr/bin/env python3
"""Parse Markdown into a faithful, styled DOCX. Editorial work is done by the host.

The same function handles files and pasted text. It does not execute embedded code,
invent headings, rewrite prose, or update Word TOC fields.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

sys.dont_write_bytecode = True
SKILL_ROOT = Path(__file__).resolve().parent.parent
for vendor in ("python_docx_1_2_0", "mistune_2_0_4"):
    sys.path.insert(0, str(SKILL_ROOT / "vendor" / vendor))

import mistune
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Mm, Pt, RGBColor

MODES = ("requirements-list", "decision-proposal", "rd-application", "rd-implementation-outline")
CHAPTER_WRAPPER_RE = re.compile(
    r"^\s*第([一二三四五六七八九十百千万零〇两0-9]+)章\s*(.+?)\s*$"
)


def element(tag, **attrs):
    node = OxmlElement(f"w:{tag}")
    for name, value in attrs.items():
        node.set(qn(f"w:{name}"), str(value))
    return node


def set_font(target, family, size, bold=False):
    target.font.name = "Times New Roman"
    target.font.size = Pt(size)
    target.font.bold = bold
    target.font.color.rgb = RGBColor(0, 0, 0)
    rpr = target._element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    for key, value in (("ascii", "Times New Roman"), ("hAnsi", "Times New Roman"), ("eastAsia", family)):
        fonts.set(qn(f"w:{key}"), value)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{attr}"), None)


def plain(tokens):
    return "".join(t.get("text", t.get("alt", "")) if not t.get("children") else plain(t["children"]) for t in tokens)


def body_outline(paragraph):
    paragraph._p.get_or_add_pPr().append(element("outlineLvl", val=9))


class MarkdownWriter:
    def __init__(self, mode, base_dir):
        self.mode = mode
        self.rd = mode.startswith("rd-")
        self.size = 14 if self.rd else 16
        self.base_dir = Path(base_dir).resolve()
        self.doc = Document()
        self.tables = 0
        self.images = []
        self.links = []
        self.records = []
        self.bookmarks = {}
        self.heading_offset = 0
        self.heading_normalizations = []
        self.setup()

    def setup(self):
        section = self.doc.sections[0]
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1.25)
        normal = self.doc.styles["Normal"]
        set_font(normal, "仿宋", self.size)
        fmt = normal.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt.first_line_indent = Pt(self.size * 2)
        fmt.space_before = fmt.space_after = Pt(0)
        fmt.line_spacing = 1.5
        fmt.widow_control = True
        for level in range(1, 7):
            style = self.doc.styles[f"Heading {level}"]
            req_detail = self.mode == "requirements-list" and level > 1
            family = "仿宋_GB2312" if self.rd and level == 1 else "仿宋" if self.rd or req_detail else "宋体"
            size = 14 if self.rd and level > 1 else 16
            set_font(style, family, size, not req_detail)
            fmt = style.paragraph_format
            fmt.left_indent = fmt.right_indent = fmt.first_line_indent = Pt(0)
            fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            fmt.line_spacing = 1.5 if self.rd or req_detail else 1.2
            fmt.space_before = Pt(12 if self.rd and level == 1 else 0 if req_detail else 6)
            fmt.space_after = Pt(12 if self.rd and level == 1 else 0 if req_detail else 4)
            fmt.keep_with_next = True
        for name in ("AOW List", "AOW Quote", "AOW Code"):
            style = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.base_style = normal
            set_font(style, "仿宋" if name != "AOW Code" else "宋体", self.size if name != "AOW Code" else 12)
            fmt = style.paragraph_format
            fmt.first_line_indent = Pt(0)
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.line_spacing = 1.5 if name != "AOW Code" else 1.0
            if name == "AOW Code":
                style.font.name = "Consolas"
        set_font(self.doc.styles["Title"], "方正小标宋简体" if self.rd else "宋体", 26 if self.rd else 18, True)
        for style in self.doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                for border in style._element.xpath("./w:pPr/w:pBdr"):
                    border.getparent().remove(border)
        footer = self.doc.styles.add_style("AOW Footer", WD_STYLE_TYPE.PARAGRAPH)
        footer.base_style = normal
        set_font(footer, "宋体", 9)
        footer.paragraph_format.first_line_indent = Pt(0)
        footer.paragraph_format.line_spacing = 1.0

    def paragraph(self, style=None):
        p = self.doc.add_paragraph(style=style)
        body_outline(p)
        return p

    def inline(self, p, tokens, bold=False, italic=False, strike=False, parent=None):
        for token in tokens:
            kind = token["type"]
            if kind in {"strong", "emphasis", "strikethrough"}:
                self.inline(p, token["children"], bold or kind == "strong", italic or kind == "emphasis", strike or kind == "strikethrough", parent)
            elif kind in {"text", "codespan", "linebreak", "inline_html"}:
                if kind == "inline_html" and not re.fullmatch(r"<br\s*/?>", token["text"], re.I):
                    raise ValueError("HTML requires a host conversion path that preserves its structure; nothing was saved")
                text = "\n" if kind in {"linebreak", "inline_html"} else token["text"]
                if kind == "text":
                    text = text.replace("\n", " ")
                run = p.add_run(text)
                if parent is not None:
                    # Hyperlink runs have a hyperlink parent rather than a paragraph
                    # in python-docx; give them the paragraph's resolved typography.
                    style = p.style
                    set_font(run, style._element.rPr.rFonts.get(qn("w:eastAsia")) or "仿宋", (style.font.size or Pt(self.size)).pt, bool(style.font.bold))
                if bold:
                    run.bold = True
                run.italic, run.font.strike = italic, strike
                if parent is not None:
                    parent.append(run._r)
            elif kind == "link":
                url = token["link"]
                link = element("hyperlink")
                if url.startswith("#"):
                    anchor = unquote(url[1:])
                    link.set(qn("w:anchor"), self.bookmarks.get(anchor, anchor))
                else:
                    link.set(qn("r:id"), self.doc.part.relate_to(url, RT.HYPERLINK, is_external=True))
                p._p.append(link)
                self.inline(p, token["children"], bold, italic, strike, link)
                self.links.append({"text": plain(token["children"]), "target": url})
            elif kind == "image":
                src = token["src"]
                if re.match(r"^[A-Za-z]:[\\/]", src):
                    path = Path(unquote(src))
                else:
                    parsed = urlparse(src)
                    if parsed.scheme or parsed.netloc:
                        raise ValueError(f"Image needs an available local asset before conversion: {src}")
                    path = self.base_dir / unquote(parsed.path)
                if not path.is_file():
                    raise FileNotFoundError(f"Markdown image is missing: {path}")
                section = self.doc.sections[-1]
                width = section.page_width - section.left_margin - section.right_margin
                shape = p.add_run().add_picture(str(path), width=width)
                max_height = section.page_height - section.top_margin - section.bottom_margin - Pt(50)
                if shape.height > max_height:
                    ratio = max_height / shape.height
                    shape.width, shape.height = int(shape.width * ratio), int(max_height)
                shape._inline.docPr.set("descr", token.get("alt", ""))
                self.images.append({"path": str(path.resolve()), "alt": token.get("alt", ""), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
            else:
                raise ValueError(f"Unsupported Markdown inline node: {kind}; use an equivalent host converter without dropping it")

    def numbering(self, ordered, start, depth):
        root = self.doc.part.numbering_part.element
        aid = max([int(x.get(qn("w:abstractNumId"))) for x in root.findall(qn("w:abstractNum"))] + [-1]) + 1
        nid = max([int(x.get(qn("w:numId"))) for x in root.findall(qn("w:num"))] + [0]) + 1
        abstract = element("abstractNum", abstractNumId=aid)
        abstract.append(element("multiLevelType", val="singleLevel"))
        lvl = element("lvl", ilvl=0)
        lvl.extend([element("start", val=start), element("numFmt", val="decimal" if ordered else "bullet"), element("lvlText", val="%1." if ordered else "•"), element("lvlJc", val="left")])
        lvl.append(element("pPr"))
        lvl[-1].append(element("ind", left=480 * (depth + 1), hanging=240))
        abstract.append(lvl)
        # OOXML requires all abstract definitions before numbering instances.
        first_num = root.find(qn("w:num"))
        if first_num is None:
            root.append(abstract)
        else:
            first_num.addprevious(abstract)
        num = element("num", numId=nid)
        num.append(element("abstractNumId", val=aid))
        root.append(num)
        return nid

    def table(self, token):
        self.tables += 1
        caption = self.paragraph()
        caption.paragraph_format.first_line_indent = Pt(0)
        caption.paragraph_format.space_before = Pt(6)
        caption.paragraph_format.space_after = Pt(4)
        caption.paragraph_format.keep_with_next = True
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(caption.add_run(f"表{self.tables}"), "宋体", 12)
        rows = [token["children"][0]["children"]] + [r["children"] for r in token["children"][1]["children"]]
        count = len(rows[0])
        if any(len(row) != count for row in rows):
            raise ValueError("Markdown table has inconsistent column counts")
        table = self.doc.add_table(rows=len(rows), cols=count)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        section = self.doc.sections[-1]
        width = section.page_width - section.left_margin - section.right_margin
        weights = [max(8, min(22, max(len(plain(row[i]["children"])) for row in rows))) for i in range(count)]
        widths = [int(width * w / sum(weights)) for w in weights]
        for column, col_width in zip(table.columns, widths):
            column.width = col_width
        borders = element("tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            borders.append(element(edge, val="single", sz=4, color="000000"))
        table._tbl.tblPr.append(borders)
        table.rows[0]._tr.get_or_add_trPr().append(element("tblHeader", val="true"))
        for ri, row in enumerate(rows):
            for ci, cell_token in enumerate(row):
                cell = table.cell(ri, ci)
                cell.width = widths[ci]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                tcpr = cell._tc.get_or_add_tcPr()
                tcpr.append(element("shd", val="clear", fill="FFFFFF"))
                margins = element("tcMar")
                for edge in ("top", "start", "bottom", "end"):
                    margins.append(element(edge, w=15 if self.mode == "requirements-list" else 80, type="dxa"))
                tcpr.append(margins)
                p = cell.paragraphs[0]
                pf = p.paragraph_format
                pf.first_line_indent = Pt(0)
                pf.space_before = pf.space_after = Pt(6 if self.mode == "requirements-list" else 3)
                pf.line_spacing = 1.0
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ri == 0 or self.mode == "requirements-list" or len(plain(cell_token["children"])) <= 12 else WD_ALIGN_PARAGRAPH.LEFT
                self.inline(p, cell_token["children"], bold=ri == 0)
                for node in p._p.xpath(".//w:r"):
                    from docx.text.run import Run
                    run = Run(node, p)
                    set_font(run, "仿宋" if self.rd else "宋体", 12, ri == 0 or bool(run.bold))

    def blocks(self, tokens, depth=0, quote=False):
        for token in tokens:
            kind = token["type"]
            if kind in {"newline", "thematic_break"}:
                continue
            if kind == "heading":
                level = max(1, token["level"] - self.heading_offset)
                p = self.doc.add_paragraph(style=f"Heading {level}")
                heading_children = token["children"]
                heading_text = plain(heading_children)
                match = CHAPTER_WRAPPER_RE.match(heading_text)
                if match:
                    normalized = f"{match.group(1)}、{match.group(2)}"
                    heading_children = [{"type": "text", "text": normalized}]
                    self.heading_normalizations.append(
                        {"from": heading_text, "to": normalized}
                    )
                self.inline(p, heading_children)
                bookmark = self.bookmarks.get(self.slug(heading_text))
                if bookmark:
                    ident = len(self.doc.element.xpath(".//w:bookmarkStart")) + 1
                    p._p.insert(1, element("bookmarkStart", id=ident, name=bookmark))
                    p._p.append(element("bookmarkEnd", id=ident))
            elif kind in {"paragraph", "block_text"}:
                p = self.paragraph("AOW Quote" if quote else None)
                if quote:
                    p.paragraph_format.left_indent = Pt(24 * (depth + 1))
                self.inline(p, token["children"])
            elif kind == "block_quote":
                self.blocks(token["children"], depth, quote=True)
            elif kind == "list":
                nid = self.numbering(token.get("ordered", False), token.get("start", 1), depth)
                for item in token["children"]:
                    children = item["children"]
                    p = self.paragraph("AOW List")
                    p.paragraph_format.left_indent = Pt(24 * (depth + 1))
                    p.paragraph_format.first_line_indent = Pt(-12)
                    if item["type"] == "task_list_item":
                        p.add_run("[x] " if item["checked"] else "[ ] ")
                    else:
                        numpr = p._p.get_or_add_pPr().get_or_add_numPr()
                        numpr.get_or_add_ilvl().val = 0
                        numpr.get_or_add_numId().val = nid
                    if children and children[0]["type"] in {"block_text", "paragraph"}:
                        self.inline(p, children[0]["children"])
                        children = children[1:]
                    self.blocks(children, depth + 1, quote)
            elif kind == "block_code":
                p = self.paragraph("AOW Code")
                # Language identifiers are metadata, not business prose. Never execute.
                p.add_run(token["text"])
            elif kind == "table":
                self.table(token)
            else:
                raise ValueError(f"Unsupported Markdown block: {kind}; use an equivalent host converter without dropping it")

    @staticmethod
    def slug(text):
        return re.sub(r"[^\w\- ]", "", text.lower()).replace(" ", "-")

    def build(self, text, title_mode="auto"):
        # An outer markdown fence is a transport wrapper only if it spans all input.
        wrapper = re.fullmatch(r"\s*(`{3,}|~{3,})(?:markdown|md)\s*\n([\s\S]*)\n\1\s*", text)
        if wrapper:
            text = wrapper.group(2)
        tokens = mistune.create_markdown(renderer="ast", plugins=["table", "strikethrough", "task_lists"])(text)
        meaningful = [t for t in tokens if t["type"] not in {"newline", "thematic_break"}]
        headings = [t for t in meaningful if t["type"] == "heading"]
        first_is_h1 = meaningful and meaningful[0]["type"] == "heading" and meaningful[0]["level"] == 1
        take_title = first_is_h1 and (title_mode == "first-h1" or (title_mode == "auto" and sum(t["level"] == 1 for t in headings) == 1))
        if take_title:
            title_token = meaningful[0]
            p = self.doc.add_paragraph(style="Title")
            body_outline(p)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.first_line_indent = Pt(0)
            pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(4), 1.2
            self.inline(p, title_token["children"])
            tokens.remove(title_token)
            self.heading_offset = 1
            if self.mode == "decision-proposal" and len(headings) > 1:
                pf.space_before = Pt(156)
                self.doc.add_section(WD_SECTION.NEW_PAGE)
        for idx, token in enumerate(headings):
            slug = self.slug(plain(token["children"]))
            self.bookmarks.setdefault(slug, f"MDHeading{idx+1}")
        self.blocks(tokens)
        if self.mode != "requirements-list":
            section = self.doc.sections[-1]
            section.footer.is_linked_to_previous = False
            p = section.footer.paragraphs[0]
            p.style = "AOW Footer"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            r = p.add_run()
            set_font(r, "宋体", 9)
            field = element("fldSimple", instr=" PAGE ")
            field.append(r._r)
            r.add_text("1")
            p._p.append(field)
        self.doc.core_properties.title = plain(meaningful[0].get("children", [])) if take_title else ""
        self.doc.core_properties.author = ""
        return {"headings": len(headings) - int(bool(take_title)), "title_extracted": bool(take_title), "tables": self.tables, "images": self.images, "links": self.links, "heading_normalizations": self.heading_normalizations, "ast": tokens}


def convert(text, output, mode, base_dir, title_mode="auto"):
    writer = MarkdownWriter(mode, base_dir)
    report = writer.build(text, title_mode)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    writer.doc.save(output)
    return {"ok": True, "output": str(output.resolve()), "source_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(), "editorial_done": False, "needs_toc": mode == "decision-proposal" and report["headings"] > 0, **report}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Markdown file, or '-' to read UTF-8 Markdown from stdin")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--mode", choices=MODES, required=True)
    parser.add_argument("--base-dir", type=Path, help="Base for relative images; defaults to the input file's directory")
    parser.add_argument("--title-mode", choices=("auto", "first-h1", "none"), default="auto")
    parser.add_argument("--encoding", default="utf-8-sig")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    try:
        path = None if args.input == "-" else Path(args.input)
        if path is not None and path.resolve() == args.out.resolve():
            raise ValueError("Output must not overwrite the Markdown input")
        text = sys.stdin.buffer.read().decode(args.encoding) if path is None else path.read_text(encoding=args.encoding)
        report = convert(text, args.out, args.mode, args.base_dir or (path.parent if path else Path.cwd()), args.title_mode)
    except Exception as exc:
        report = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "ast"}, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
