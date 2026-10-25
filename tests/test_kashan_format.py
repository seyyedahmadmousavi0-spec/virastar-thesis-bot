# -*- coding: utf-8 -*-
"""Clauses of the Kashan three-page guide on the official template.

tests/kashan_template.py rebuilds the official Bahman 1398 template (images for
the basmala and logo, Nazanin 12, Titr headings, 3 cm margins, abjad numbering
from the table of contents). The clean build must raise nothing; the seeded
build breaks one clause per «SEED» line and each must be reported in Word.
"""
import os
os.environ["BALE_SKIP_DOTENV"] = "1"
from collections import Counter
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from docx import Document

import checker
from safe_annotator import annotate_document
from unified_word_review import finalize_word_review

BUILDER = Path(__file__).with_name("kashan_template.py")
SEEDED = {"KASHAN_COVER_FONT", "KASHAN_DEGREE_BOLD", "KASHAN_COVER_BY", "KASHAN_COVER_DATE",
          "KASHAN_EN_COVER_ITALIC", "KASHAN_EN_THE", "KASHAN_ACK_NAMES", "KASHAN_KEYWORDS",
          "KASHAN_LIST_TITLE", "KASHAN_HEADING_FONT", "KASHAN_BODY_FONT", "KASHAN_MARGIN"}


class KashanTemplateTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="kashan_format_")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def review(self, degree, variant, scope="full"):
        source = self.root / f"{degree}_{variant}.docx"
        subprocess.run([sys.executable, "-B", str(BUILDER), str(source), degree, variant], check=True)
        rules = checker.load_rules(mode=f"kashan_{degree}")
        rules["document_scope"] = scope
        document, issues, stats = checker.analyze(str(source), f"kashan_{degree}", rules)
        output = self.root / f"{degree}_{variant}_reviewed.docx"
        annotate_document(document, issues, output)
        stats["annotation_summary"] = document._last_annotation_summary
        result = finalize_word_review(source, output, issues, stats, rules)
        self.assertEqual(len(issues), len(Document(output).comments))
        self.assertEqual(len(issues), result["comments"])
        return issues, stats

    def test_official_template_has_no_findings(self):
        for degree in ("master", "phd"):
            with self.subTest(degree=degree):
                issues, stats = self.review(degree, "clean")
                self.assertEqual([], [(i.rule_id, i.message[:60]) for i in issues])
                rows = {row["id"]: row["status"] for row in stats["university_checklist"]}
                for page in ("basmala", "persian_cover", "defense_minutes", "declaration", "acknowledgements",
                             "abstract_fa", "toc", "references", "abstract_en", "english_cover"):
                    self.assertEqual("detected", rows[page], page)

    def test_every_seeded_clause_is_reported_in_word(self):
        for degree in ("master", "phd"):
            with self.subTest(degree=degree):
                issues, _ = self.review(degree, "seeded")
                found = Counter(issue.rule_id for issue in issues)
                self.assertEqual(SEEDED, SEEDED & set(found))
                self.assertEqual(2, found["KASHAN_COVER_DATE"])      # Solar and Gregorian dates
                self.assertEqual(2, found["KASHAN_DEGREE_BOLD"])     # Persian and English covers
                self.assertTrue(all(issue.confidence == "B" for issue in issues if issue.rule_id.startswith("KASHAN_")))
                self.assertTrue(all("مستند:" in issue.message for issue in issues if issue.rule_id.startswith("KASHAN_")))

    def test_fragment_scope_skips_whole_document_claims(self):
        issues, _ = self.review("master", "seeded", scope="fragment")
        ids = {issue.rule_id for issue in issues}
        self.assertFalse(ids & {"KASHAN_SECTION_REVIEW", "KASHAN_PAGE_ORDER", "KASHAN_PAGE_NUMBERS"})
        self.assertIn("KASHAN_MARGIN", ids)

    def test_checklist_lists_what_a_file_cannot_prove(self):
        _, stats = self.review("master", "clean")
        rows = {row["id"]: row for row in stats["university_checklist"]}
        for item in ("cd", "irandoc", "signatures", "chapter_page_numbers", "cover_layout"):
            self.assertEqual("manual", rows[item]["status"])
        self.assertIn("sabt.irandoc.ac.ir", rows["irandoc"]["detail"])


if __name__ == "__main__":
    unittest.main()
