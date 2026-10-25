# -*- coding: utf-8 -*-
"""Reviewed Word files of long, error-dense theses must stay openable.

A real thesis produced thousands of comments that Word could not open; most
were the same few mistakes repeated. Repeats beyond the first three are
highlighted where they occur and counted in the first comment, and every part
of the package is compressed.
"""
import os
os.environ["BALE_SKIP_DOTENV"] = "1"
from pathlib import Path
import tempfile
import unittest
import zipfile

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

import checker
from safe_annotator import FOLD_REPEATED_AFTER
from unified_word_review import finalize_word_review

Z = "‌"


class LargeWordOutputTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="large_word_")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / "thesis.docx"
        document = Document()
        for index in range(40):
            # The same two mistakes in every paragraph, plus one unique typo.
            number = str(index + 1).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
            document.add_paragraph(f"داده ها در مرحلهٔ {number} گردآوری شد و نتایج نشان میدهد که روش مناسب است.")
        document.add_paragraph("این روش بعنوان روش اصلی انتخاب شد.")
        document.save(self.source)

    def review(self, **options):
        rules = checker.load_rules(mode="article")
        parsed, issues, stats = checker.analyze(str(self.source), "article", rules)
        output = self.root / "reviewed.docx"
        checker.annotate(parsed, issues, output, **options)
        stats["annotation_summary"] = parsed._last_annotation_summary
        result = finalize_word_review(self.source, output, issues, stats, rules)
        return issues, stats["annotation_summary"], result, output

    def test_repeats_are_highlighted_and_counted_in_the_first_comment(self):
        issues, summary, result, output = self.review()
        reviewed = Document(output)
        repeated = [i for i in issues if i.rule_id in {"plural_space", "mi_prefix"}]
        self.assertEqual(80, len(repeated))
        # Three comments per repeated mistake; other findings are unaffected.
        repeat_comments = [c for c in reviewed.comments if c.text.startswith(("plural_space", "mi_prefix"))]
        self.assertEqual(2 * FOLD_REPEATED_AFTER, len(repeat_comments))
        self.assertEqual(len(issues) - summary["folded"], len(reviewed.comments))
        self.assertEqual(80 - 2 * FOLD_REPEATED_AFTER, summary["folded"])
        self.assertEqual(len(issues), result["highlighted"])
        self.assertEqual(len(issues) - summary["folded"], result["comments"])
        # Every repeat, folded or not, is yellow at its own position.
        yellow = [run.text for p in reviewed.paragraphs for run in p.runs
                  if run.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        self.assertEqual(40, yellow.count("داده ها"))
        self.assertEqual(40, yellow.count("میده"))
        totals = [c.text for c in reviewed.comments if "بار در سند آمده" in c.text]
        self.assertEqual(2, len(totals))
        self.assertTrue(all("40 بار" in text and "Ctrl+H" in text for text in totals))
        records = summary["records"]
        for record in records:
            if record["status"] == "folded":
                self.assertEqual("applied", records[record["folded_into"]]["status"])

    def test_fold_can_be_disabled(self):
        issues, summary, result, output = self.review(fold_after=0)
        self.assertEqual(0, summary["folded"])
        self.assertEqual(len(issues), len(Document(output).comments))

    def test_every_package_part_is_compressed(self):
        _, _, _, output = self.review(fold_after=0)
        with zipfile.ZipFile(output) as package:
            stored = [item.filename for item in package.infolist()
                      if item.compress_type == zipfile.ZIP_STORED and item.file_size > 1024]
        self.assertEqual([], stored)


if __name__ == "__main__":
    unittest.main()
