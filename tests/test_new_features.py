import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from format_references import format_record  # noqa: E402
from markdown_to_docx import convert, Document  # noqa: E402


class NewFeatureTests(unittest.TestCase):
    def test_gbt7714_journal_record(self):
        result = format_record(
            {
                "type": "J",
                "authors": ["张三", "李四"],
                "title": "测试方法",
                "source": "测试学报",
                "year": "2026",
                "volume": "10",
                "issue": "2",
                "pages": "1-8",
            },
            1,
        )
        self.assertEqual(result, "[1] 张三, 李四. 测试方法[J]. 测试学报, 2026, 10(2): 1-8.")

    def test_gbt7714_doi_has_sentence_boundary(self):
        result = format_record(
            {
                "type": "J",
                "authors": ["A"],
                "title": "T",
                "source": "S",
                "year": "2026",
                "pages": "1-2",
                "doi": "10.1234/example",
            },
            1,
        )
        self.assertEqual(result, "[1] A. T[J]. S, 2026: 1-2. DOI: 10.1234/example.")

    def test_batch_cli_converts_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            source = base / "drafts"
            output = base / "build"
            source.mkdir()
            (source / "one.md").write_text("# 标题\n\n## 技术领域\n\n正文。\n", encoding="utf-8")
            command = [
                sys.executable,
                str(SCRIPTS / "batch_convert.py"),
                str(source),
                "--out-dir",
                str(output),
                "--mode",
                "patent",
            ]
            completed = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertTrue((output / "one.docx").is_file())
            validated = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS / "validate_standardized_docx.py"),
                    str(output / "one.docx"),
                    "--mode",
                    "patent",
                    "--processing-strength",
                    "format-only",
                ],
                cwd=ROOT,
                capture_output=True,
                text=True,
            )
            self.assertEqual(validated.returncode, 0, validated.stderr + validated.stdout)

    def test_missing_reference_field_is_explicit(self):
        with self.assertRaisesRegex(ValueError, "missing source"):
            format_record({"type": "J", "authors": ["A"], "title": "T", "year": "2026"}, 1)

    def test_gbt7714_journal_issue_without_volume_is_preserved(self):
        result = format_record(
            {
                "type": "J",
                "authors": ["张三"],
                "title": "标题",
                "source": "学报",
                "year": "2026",
                "issue": "2",
                "pages": "1-8",
            },
            1,
        )
        self.assertEqual(result, "[1] 张三. 标题[J]. 学报, 2026(2): 1-8.")

    def test_gbt7714_electronic_resource_without_publication_date(self):
        result = format_record(
            {
                "type": "EB/OL",
                "authors": ["张三"],
                "title": "网页",
                "accessed": "2026-01-01",
                "url": "https://example.com",
            },
            1,
        )
        self.assertEqual(result, "[1] 张三. 网页[EB/OL][2026-01-01]. https://example.com.")

    def test_gbt7714_electronic_resource_with_publication_date(self):
        result = format_record(
            {
                "type": "EB/OL",
                "authors": ["张三"],
                "title": "网页",
                "published": "2025-12-01",
                "accessed": "2026-01-01",
                "url": "https://example.com",
            },
            1,
        )
        self.assertEqual(result, "[1] 张三. 网页[EB/OL]. (2025-12-01)[2026-01-01]. https://example.com.")

    def test_gbt7714_book_thesis_and_conference_templates(self):
        cases = [
            (
                {"type": "M", "authors": ["甲"], "title": "书", "place": "北京", "publisher": "出版社", "year": "2026"},
                "[1] 甲. 书[M]. 北京: 出版社, 2026.",
            ),
            (
                {"type": "D", "authors": ["乙"], "title": "论文", "place": "上海", "institution": "某大学", "year": "2026"},
                "[1] 乙. 论文[D]. 上海: 某大学, 2026.",
            ),
            (
                {"type": "C", "authors": ["丙"], "title": "会议论文", "collection": "论文集", "place": "广州", "publisher": "出版社", "year": "2026", "pages": "10-12"},
                "[1] 丙. 会议论文[C]//论文集. 广州: 出版社, 2026: 10-12.",
            ),
        ]
        for record, expected in cases:
            with self.subTest(kind=record["type"]):
                self.assertEqual(format_record(record, 1), expected)

    def test_gbt7714_required_fields_fail_for_each_core_type(self):
        incomplete = [
            {"type": "M", "authors": ["甲"], "title": "书", "publisher": "出版社", "year": "2026"},
            {"type": "D", "authors": ["乙"], "title": "论文", "place": "上海", "year": "2026"},
            {"type": "C", "authors": ["丙"], "title": "会议论文", "collection": "论文集", "place": "广州", "year": "2026"},
            {"type": "EB/OL", "authors": ["丁"], "title": "网页", "url": "https://example.com"},
        ]
        for record in incomplete:
            with self.subTest(kind=record["type"]):
                with self.assertRaises(ValueError):
                    format_record(record, 1)

    def test_gbt7714_english_author_folding(self):
        result = format_record(
            {"type": "J", "authors": ["A", "B", "C", "D"], "title": "T", "source": "S", "year": "2026"},
            1,
        )
        self.assertEqual(result, "[1] A, B, C, et al. T[J]. S, 2026.")

    def test_plain_patent_claim_gets_hanging_indent(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "claim.docx"
            convert(
                "# 标题\n\n## 权利要求书\n\n1、一种方法。\n",
                output,
                "patent",
                Path(tmp),
            )
            document = Document(output)
            claim = next(paragraph for paragraph in document.paragraphs if "一种方法" in paragraph.text)
            self.assertAlmostEqual(claim.paragraph_format.left_indent.pt, 28.0, places=1)
            self.assertAlmostEqual(claim.paragraph_format.first_line_indent.pt, -28.0, places=1)


if __name__ == "__main__":
    unittest.main()
