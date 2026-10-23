"""Report audits distinguish verified source anchors from an independent review Word."""
import json
import unittest

from analysis_model import Issue
from report_artifacts import (annotation_summary_lines, html_report, report_data,
                              safe_annotation_summary, safe_review_document_summary)


class AnnotationCompletenessReportTests(unittest.TestCase):
    def findings(self):
        return [Issue("EXACT", message="first", confidence="A", block_id="b1"),
                Issue("PARAGRAPH", message="second", confidence="B", block_id="b2"),
                Issue("MANUAL", message="third", confidence="C", story="header")]

    def stats(self):
        return {"ambiguous_candidates": 0,
                "annotation_summary": {
                    "requested": 3, "applied": 2, "unapplied": 1,
                    "exact_applied": 1, "paragraph_applied": 1, "highlighted": 2,
                    "max_comments": None, "output_created": True,
                    "output_path": "C:/private/source-annotated.docx",
                    "reasons": {"unsupported_story": 1},
                    "records": [
                        {"index": 0, "rule_id": "EXACT", "block_id": "b1", "status": "applied", "anchor_scope": "exact"},
                        {"index": 1, "rule_id": "PARAGRAPH", "block_id": "b2", "status": "applied", "anchor_scope": "paragraph"},
                        {"index": 2, "rule_id": "MANUAL", "status": "unapplied", "reason": "unsupported_story"}]},
                "review_document": {"created": True, "findings": 3, "comments": 3,
                                    "highlighted": 3, "filetype": "docx",
                                    "path": "C:/private/full-review.docx"}}

    def test_all_findings_and_both_word_audits_round_trip_without_paths(self):
        data = report_data(self.findings(), self.stats(), {"strict": False})
        exported = json.loads(json.dumps(data, ensure_ascii=False))
        self.assertEqual(3, exported["summary"]["findings"])
        self.assertEqual(3, len(exported["findings"]))
        self.assertIsNone(exported["stats"]["annotation_summary"]["max_comments"])
        self.assertEqual(1, exported["stats"]["annotation_summary"]["exact_applied"])
        self.assertEqual(1, exported["stats"]["annotation_summary"]["paragraph_applied"])
        self.assertEqual(3, exported["stats"]["review_document"]["comments"])
        self.assertEqual("paragraph", exported["findings"][1]["annotation"]["anchor_scope"])
        self.assertEqual("unsupported_story", exported["findings"][2]["annotation"]["reason"])
        self.assertNotIn("C:/private", json.dumps(exported))

    def test_html_does_not_present_paragraph_or_companion_anchors_as_exact(self):
        html = html_report(report_data(self.findings(), self.stats(), {"strict": False}))
        self.assertIn("1 کامنت در محدودهٔ دقیق و 1 کامنت در سطح پاراگراف", html)
        self.assertIn("در Word در سطح پاراگراف کامنت شد؛ محل دقیق عبارت نیست", html)
        self.assertIn("Word گزارش کامل: 3 مورد، 3 کامنت و 3 مورد هایلایت زرد", html)
        self.assertIn("سند مستقلی از یافته‌هاست", html)
        self.assertIn("محل کامنت‌های آن، محل دقیق ایراد در سند اصلی نیست", html)
        self.assertIn("بدون سقف تعداد", html)
        self.assertIn("موارد سطح C فقط برای بررسی دستی‌اند", html)
        self.assertIn("محل دقیقِ تأییدنشده در سند اصلی ساخته نمی‌شود", html)
        self.assertIn('data-confidence="C"', html)
        self.assertIn("درج امن کامنت در این بخش پشتیبانی نمی‌شود", html)

    def test_incomplete_companion_word_remains_explicit_when_original_is_complete(self):
        stats = self.stats()
        stats["review_document"].update(comments=2, highlighted=1)
        lines = "\n".join(annotation_summary_lines(stats))
        self.assertIn("3 مورد، 2 کامنت و 1 مورد هایلایت زرد", lines)
        self.assertIn("همهٔ موارد در Word گزارش کامل تکمیل نشده", lines)
        stats["review_document"]["created"] = False
        html = html_report(report_data(self.findings(), stats, {"strict": False}))
        self.assertIn("تولید Word گزارش کامل انجام نشد", html)
        self.assertNotIn("Word گزارش کامل: 3 مورد", html)

    def test_untrusted_annotation_metadata_cannot_invent_anchor_scope_or_leak_paths(self):
        safe = safe_annotation_summary({"applied": True, "exact_applied": -1,
                                        "paragraph_applied": 2, "highlighted": "3", "max_comments": None,
                                        "records": [{"index": 0, "anchor_scope": ["exact"], "path": "secret"}]})
        self.assertNotIn("applied", safe)
        self.assertEqual(0, safe["exact_applied"])
        self.assertNotIn("highlighted", safe)
        self.assertNotIn("anchor_scope", safe["records"][0])
        self.assertNotIn("path", safe["records"][0])
        review = safe_review_document_summary({"created": True, "findings": 4, "comments": True,
                                               "highlighted": -3, "filetype": "<script>", "path": "secret"})
        self.assertEqual({"created": True, "findings": 4, "highlighted": 0}, review)

    def test_shared_summary_preserves_pdf_experimental_note_and_all_skip_totals(self):
        stats = {"annotation_format": "pdf", "annotation_summary": {
            "requested": 4, "applied": 1, "unapplied": 3,
            "output_created": True, "experimental": True,
            "reasons": {"protected_content": 1, "uncertain_location": 2}}}
        lines = "\n".join(annotation_summary_lines(stats))
        self.assertIn("از 4 مورد، 1 مورد در PDF کامنت شد", lines)
        self.assertIn("نشانه‌گذاری PDF آزمایشی است", lines)
        self.assertIn("تداخل با محتوای حفاظت‌شده: 1", lines)
        self.assertIn("محل دقیق قابل تأیید نیست: 2", lines)
        html = html_report(report_data([], stats, {}))
        self.assertEqual(1, html.count("تداخل با محتوای حفاظت‌شده: 1"))
        self.assertEqual(1, html.count("محل دقیق قابل تأیید نیست: 2"))


if __name__ == "__main__":
    unittest.main()
