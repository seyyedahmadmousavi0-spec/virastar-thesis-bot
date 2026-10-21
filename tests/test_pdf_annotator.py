"""PDF feasibility checks: real shaped RTL glyphs, coordinates and preservation.

These assertions check MuPDF extraction/serialization, not how every external
PDF reader displays Persian annotation popups. General activation needs that
additional manual reader assessment.
"""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pymupdf

from analysis_model import Issue
from doc_guard import DocumentIntegrityError, sha256_file
from pdf_handler import extract_pdf_content
import pdf_annotator


class PdfAnnotationTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "source.pdf"
        self.output = self.root / "review.pdf"

    def create_ascii(self, existing_note=False):
        with pymupdf.open() as document:
            page = document.new_page()
            page.insert_text((50, 70), "target unchanged target 123.45 scientific content")
            page.insert_text((50, 140), "target in another paragraph")
            if existing_note:
                page.add_text_annot((30, 30), "Existing author comment")
            document.save(self.source)

    def issue(self, block, target, *, occurrence=0, confidence="A", rule="test_rule"):
        offset = -1
        for _ in range(occurrence + 1):
            offset = block["text"].find(target, offset + 1)
            self.assertGreaterEqual(offset, 0)
        issue = Issue(rule, message="این مورد نیازمند بررسی است.",
                      suggestion="صورت پیشنهادی را بررسی کنید.", confidence=confidence,
                      target_text=target, char_start=offset, char_end=offset + len(target),
                      block_id=block["block_id"], page_index=block["page_index"],
                      source_part=f"pdf/page/{block['page_index'] + 1}", story="pdf")
        self.assertTrue(issue.validate_evidence(block["text"]))
        return issue

    def blocks(self, page=0):
        extracted = extract_pdf_content(self.source)
        self.assertTrue(extracted["ok"])
        return extracted["pages"][page]["blocks"]

    def assert_exact_vertices(self, issue, annotation):
        with pymupdf.open(self.source) as original:
            page = original[issue.page_index]
            block_number = int(issue.block_id.rsplit("b", 1)[1])
            block = pdf_annotator._page_blocks(page)[block_number]
            expected = []
            for position in block[1][issue.char_start:issue.char_end]:
                if position is not None:
                    direction, span, char, _ = position
                    quad = pymupdf.recover_char_quad(direction, span, char)
                    if not quad.rect.is_empty:
                        expected.extend(tuple(point) for point in quad)
            self.assertEqual(len(expected), len(annotation.vertices))
            for expected_point, actual_point in zip(expected, annotation.vertices):
                for left, right in zip(expected_point, actual_point):
                    self.assertAlmostEqual(left, right, places=3)

    def test_exact_offsets_choose_second_repeat_and_preserve_original_streams_and_notes(self):
        self.create_ascii(existing_note=True)
        original_hash = sha256_file(self.source)
        issue = self.issue(self.blocks()[0], "target", occurrence=1)
        with pymupdf.open(self.source) as original:
            streams = pdf_annotator._streams(original)
            payload = pdf_annotator._payload(original)
            notes = pdf_annotator._existing_notes(original)
        result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertEqual((1, 1, 0, True), tuple(result[key] for key in ("requested", "applied", "unapplied", "output_created")))
        self.assertEqual(original_hash, sha256_file(self.source))
        with pymupdf.open(self.output) as document:
            self.assertEqual(payload, pdf_annotator._payload(document))
            self.assertEqual(streams, pdf_annotator._streams(document, streams))
            for xref, value in notes.items():
                self.assertEqual(value, document.xref_object(xref))
            page = document[0]
            annotation = page.load_annot(result["records"][0]["annotation_xref"])
            self.assert_exact_vertices(issue, annotation)
            self.assertIn(issue.message, annotation.info["content"])
            self.assertGreater(annotation.rect.x0, 100)  # second occurrence

    def test_real_checker_findings_preserve_pdf_block_evidence(self):
        import checker
        with pymupdf.open() as document:
            page = document.new_page()
            page.insert_text((50, 70), "This is a experimental study and researchers are interested on this method.")
            document.save(self.source)
        _, issues, stats = checker.analyze_pdf(self.source, mode="english")
        self.assertEqual(sha256_file(self.source), stats["file_sha256"])
        selected = [issue for issue in issues if issue.rule_id in {"EN_ARTICLE", "EN_PREP"}]
        self.assertTrue(selected)
        self.assertTrue(all(issue.block_id.startswith("pdf:p1:b") for issue in selected))
        result = pdf_annotator.annotate_pdf(self.source, selected, self.output)
        self.assertEqual(len(selected), result["applied"], result)
        with pymupdf.open(self.output) as document:
            page = document[0]
            for issue, record in zip(selected, result["records"]):
                self.assert_exact_vertices(issue, page.load_annot(record["annotation_xref"]))

    def test_mixed_shaped_persian_english_and_numbers_map_to_exact_raw_glyphs(self):
        if not hasattr(pymupdf.Page, "insert_htmlbox"):
            self.skipTest("RTL fixture rendering requires PyMuPDF insert_htmlbox")
        with pymupdf.open() as document:
            page = document.new_page()
            font = Path("C:/Windows/Fonts/arial.ttf")
            options = {}
            css = "div {font-family:sans-serif; font-size:16pt;}"
            if font.is_file():
                css = "@font-face {font-family:test; src:url(arial.ttf);} div {font-family:test; font-size:16pt;}"
                options["archive"] = pymupdf.Archive(str(font.parent))
            page.insert_htmlbox(pymupdf.Rect(30, 40, 570, 250),
                                '<div dir="rtl">پژوهش نمونه ها با ۱۲۳ و ABC بررسی شد.</div>',
                                css=css, **options)
            document.save(self.source)
        blocks = self.blocks()
        text = blocks[0]["text"]
        self.assertTrue(any("\u0600" <= char <= "\u06ff" or "\ufb50" <= char <= "\ufeff" for char in text))
        # Use exactly the raw presentation forms seen by the checker, not a
        # normalized string which could have different offsets or glyph counts.
        targets = [text.split("\n")[0].replace("\xa0", " ").split()[0], "۱۲۳", "ABC"]
        issues = [self.issue(next(block for block in blocks if target in block["text"]), target,
                             rule=f"mixed_{index}") for index, target in enumerate(targets)]
        before = sha256_file(self.source)
        result = pdf_annotator.annotate_pdf(self.source, issues, self.output)
        self.assertEqual(3, result["applied"], result)
        self.assertEqual(before, sha256_file(self.source))
        with pymupdf.open(self.output) as document:
            page = document[0]
            for issue, record in zip(issues, result["records"]):
                annotation = page.load_annot(record["annotation_xref"])
                self.assert_exact_vertices(issue, annotation)
                self.assertIn("دلیل:", annotation.info["content"])

    def test_persian_zwnj_diacritics_and_multiline_targets_preserve_evidence(self):
        if not hasattr(pymupdf.Page, "insert_htmlbox"):
            self.skipTest("RTL fixture rendering requires PyMuPDF insert_htmlbox")
        with pymupdf.open() as document:
            page = document.new_page()
            page.insert_htmlbox(pymupdf.Rect(30, 40, 570, 250),
                                '<div dir="rtl">می‌رود؛ عَلَم و ۱۲۳ آزمایش<br>پژوهش ادامه دارد.</div>',
                                css="div {font-family:sans-serif; font-size:16pt;}")
            document.save(self.source)
        blocks = self.blocks()
        # Extraction may not retain invisible ZWNJ or may split diacritic glyphs.
        # The annotation must follow that exact evidence or safely decline it.
        target_block = blocks[0]
        target = target_block["text"].strip("\n")
        issue = self.issue(target_block, target)
        original_hash = sha256_file(self.source)
        result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertEqual(original_hash, sha256_file(self.source))
        if result["output_created"]:
            with pymupdf.open(self.output) as document:
                page = document[0]
                self.assert_exact_vertices(issue, page.load_annot(result["records"][0]["annotation_xref"]))
        else:
            self.assertIn(result["records"][0]["reason"],
                          {"ambiguous_glyph", "invalid_geometry", "extraction_mapping_mismatch", "grapheme_boundary"})
            self.assertFalse(self.output.exists())

    def test_overlapping_range_is_accounted_instead_of_double_highlight(self):
        self.create_ascii()
        block = self.blocks()[0]
        broad = self.issue(block, "target unchanged")
        overlap = self.issue(block, "unchanged", confidence="B")
        result = pdf_annotator.annotate_pdf(self.source, [overlap, broad], self.output)
        self.assertEqual(1, result["applied"])
        self.assertEqual({"overlapping_target": 1}, result["reasons"])

    def test_rotated_text_rotated_page_and_columns_keep_independent_coordinates(self):
        with pymupdf.open() as document:
            page = document.new_page()
            page.insert_text((80, 100), "left column target")
            page.insert_text((350, 100), "right column target")
            page.insert_text((250, 300), "rotated target", rotate=90)
            page.set_rotation(90)
            document.save(self.source)
        blocks = self.blocks()
        self.assertEqual(3, sum(block["text"].count("target") for block in blocks))
        issues = [self.issue(block, "target", occurrence=occurrence, rule=f"column_{n}_{occurrence}")
                  for n, block in enumerate(blocks)
                  for occurrence in range(block["text"].count("target"))]
        result = pdf_annotator.annotate_pdf(self.source, issues, self.output)
        self.assertEqual(3, result["applied"], result)
        with pymupdf.open(self.output) as document:
            page = document[0]
            for issue, record in zip(issues, result["records"]):
                self.assert_exact_vertices(issue, page.load_annot(record["annotation_xref"]))

    def test_priority_cap_and_complete_skip_accounting(self):
        self.create_ascii()
        block = self.blocks()[0]
        issues = [self.issue(block, "unchanged", confidence="B", rule="review"),
                  self.issue(block, "target", confidence="A", rule="certain"),
                  self.issue(block, "scientific", confidence="C", rule="ambiguous")]
        result = pdf_annotator.annotate_pdf(self.source, issues, self.output, max_comments=1)
        self.assertEqual(["comment_limit", "", "confidence_filtered"], [record["reason"] for record in result["records"]])
        self.assertEqual(3, result["requested"])
        self.assertEqual(2, result["unapplied"])
        self.assertEqual({"comment_limit": 1, "confidence_filtered": 1}, result["reasons"])

    def test_wrong_evidence_block_page_protected_and_duplicates_are_never_guessed(self):
        self.create_ascii()
        block = self.blocks()[0]
        good = self.issue(block, "target")
        bad_offset = self.issue(block, "target")
        bad_offset.char_start += 1
        bad_page = self.issue(block, "target")
        bad_page.page_index = 1
        bad_block = self.issue(block, "target")
        bad_block.block_id = "pdf:p1:b20"
        protected = self.issue(block, "scientific")
        protected.protected_overlap = True
        duplicate = self.issue(block, "target", rule="another_rule")
        result = pdf_annotator.annotate_pdf(self.source, [bad_offset, bad_page, bad_block, protected, good, duplicate], self.output)
        self.assertEqual(1, result["applied"], result)
        self.assertEqual(["evidence_mismatch", "missing_exact_location", "extraction_mapping_mismatch", "protected_content", "", "duplicate_target"],
                         [record["reason"] for record in result["records"]])

    def test_image_only_and_structural_findings_create_no_misleading_copy(self):
        with pymupdf.open() as document:
            page = document.new_page()
            pixel = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 4, 4), False)
            pixel.clear_with(255)
            page.insert_image(pymupdf.Rect(40, 40, 300, 300), pixmap=pixel)
            document.save(self.source)
        issue = Issue("pdf_page_unreadable", confidence="B", page_index=0, story="pdf")
        result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertEqual(0, result["applied"])
        self.assertFalse(result["output_created"])
        self.assertFalse(self.output.exists())
        self.assertEqual("missing_exact_location", result["records"][0]["reason"])

    def test_zero_cap_and_no_findings_publish_no_copy(self):
        self.create_ascii()
        issue = self.issue(self.blocks()[0], "target")
        for issues, cap in (([issue], 0), ([], 250)):
            result = pdf_annotator.annotate_pdf(self.source, issues, self.output, max_comments=cap)
            self.assertFalse(result["output_created"])
            self.assertFalse(self.output.exists())

    def test_encrypted_pdf_is_closed_without_publishing(self):
        with pymupdf.open() as document:
            document.new_page().insert_text((30, 30), "Private example")
            document.save(self.source, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                          owner_pw="owner", user_pw="reader")
        result = pdf_annotator.annotate_pdf(self.source, [Issue("x")], self.output)
        self.assertEqual({"encrypted": 1}, result["reasons"])
        self.assertFalse(self.output.exists())

    def test_same_path_hardlink_and_symlink_aliases_are_rejected(self):
        self.create_ascii()
        before = sha256_file(self.source)
        with self.assertRaises(DocumentIntegrityError):
            pdf_annotator.annotate_pdf(self.source, [], self.source)
        import os
        alias = self.root / "alias.pdf"
        os.link(self.source, alias)
        with self.assertRaises(DocumentIntegrityError):
            pdf_annotator.annotate_pdf(self.source, [], alias)
        self.assertEqual(before, sha256_file(self.source))

    def test_changed_saved_content_is_rejected_and_temporary_copy_removed(self):
        self.create_ascii()
        issue = self.issue(self.blocks()[0], "target")
        real_payload = pdf_annotator._payload
        calls = []
        def inconsistent_payload(document):
            value = real_payload(document)
            calls.append(1)
            return value if len(calls) == 1 else value + [("changed",)]
        with patch.object(pdf_annotator, "_payload", side_effect=inconsistent_payload):
            with self.assertRaises(DocumentIntegrityError):
                pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertFalse(self.output.exists())
        self.assertFalse(list(self.root.glob(".review_pdf_*")))

    def test_extraction_mapping_mismatch_and_invalid_geometry_fail_closed(self):
        self.create_ascii()
        issue = self.issue(self.blocks()[0], "target")
        with patch.object(pdf_annotator, "_page_blocks", return_value={}):
            result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertEqual({"extraction_mapping_mismatch": 1}, result["reasons"])
        with patch.object(pymupdf, "recover_char_quad", return_value=pymupdf.Quad((0, 0), (0, 0), (0, 0), (0, 0))):
            result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertEqual({"invalid_geometry": 1}, result["reasons"])
        self.assertFalse(self.output.exists())

    def test_limits_and_serialization_failure_leave_no_partial_artifact(self):
        self.create_ascii()
        issue = self.issue(self.blocks()[0], "target")
        for cap in (-1, True, 5001, 1.2):
            with self.assertRaises(ValueError):
                pdf_annotator.annotate_pdf(self.source, [issue], self.output, max_comments=cap)
        with patch.object(pymupdf.Document, "save", side_effect=OSError("synthetic write failure")):
            result = pdf_annotator.annotate_pdf(self.source, [issue], self.output)
        self.assertFalse(result["output_created"])
        self.assertEqual({"annotation_failed": 1}, result["reasons"])
        self.assertFalse(self.output.exists())
        self.assertFalse(list(self.root.glob(".review_pdf_*")))


if __name__ == "__main__":
    unittest.main()
