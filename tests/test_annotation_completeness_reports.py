"""Report audits distinguish exact locations from qualified existing-text anchors."""
import json
import unittest

from analysis_model import Issue
from report_artifacts import (annotation_summary_lines, html_report, report_data,
                              safe_annotation_summary, safe_review_document_summary,
                              safe_unified_word_review)


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

    def unified_stats(self):
        stats = self.stats()
        stats.pop("review_document")
        stats["unified_word_review"] = {
            "output_created": True, "requested": 3, "source_comments": 2,
            "context_comments": 1, "appendix_comments": 0, "comments": 3, "highlighted": 3,
            "exact_comments": 1, "paragraph_comments": 1,
            "context_records": [{"index": 2, "anchor_par_index": 0,
                                 "anchor_xml_path": "/w:document/w:body/w:p[1]",
                                 "scope": "story", "source_story": "header", "reason": "unsupported_story",
                                 "path": "C:/private/source.docx"}],
            "path": "C:/private/review.docx", "source_content": "private manuscript"}
        return stats

    def test_unified_word_counts_round_trip_without_changing_original_anchor_audit(self):
        data = report_data(self.findings(), self.unified_stats(), {"strict": False})
        exported = json.loads(json.dumps(data, ensure_ascii=False))
        unified = exported["stats"]["unified_word_review"]
        self.assertEqual(3, unified["comments"])
        self.assertEqual(2, unified["source_comments"])
        self.assertEqual(1, unified["context_comments"])
        self.assertEqual(0, unified["appendix_comments"])
        self.assertEqual(2, exported["stats"]["annotation_summary"]["applied"])
        self.assertEqual("unapplied", exported["findings"][2]["annotation"]["status"])
        self.assertNotIn("anchor_scope", exported["findings"][2]["annotation"])
        self.assertEqual("story", exported["findings"][2]["context_annotation"]["scope"])
        self.assertEqual(0, exported["findings"][2]["context_annotation"]["anchor_par_index"])
        self.assertNotIn("C:/private", json.dumps(exported))
        self.assertNotIn("private manuscript", json.dumps(exported))

    def test_same_word_explains_context_comments_without_claiming_exact_source_anchors(self):
        data = report_data(self.findings(), self.unified_stats(), {"strict": False})
        lines = "\n".join(annotation_summary_lines(data["stats"]))
        self.assertIn("همان Word بازبینی: از 3 مورد، 3 مورد کامنت و 3 مورد هایلایت زرد دارد", lines)
        self.assertIn("2 مورد در محل قابل تأیید و 1 مورد با کامنت کلی/زمینه‌ای روی متن موجود", lines)
        self.assertIn("ممیزی محل‌های قابل تأیید: از 3 مورد، 2 مورد در محل قابل تأیید کامنت شد", lines)
        self.assertIn("عبارت زردشده لزوماً خطا نیست", lines)
        self.assertIn("سقف کامنت در محل‌های قابل تأیید: بدون سقف تعداد", lines)
        self.assertNotIn("از 3 مورد، 2 مورد در Word کامنت شد", lines)
        html = html_report(data)
        self.assertIn("کامنت کلی/زمینه‌ای روی متن موجود؛ عبارت زردشده لزوماً خطا نیست", html)
        self.assertIn("محل کامنت زمینه‌ای: پاراگراف 1", html)
        self.assertIn("در Word در سطح پاراگراف کامنت شد؛ محل دقیق عبارت نیست", html)
        self.assertNotIn("سند مستقلی از یافته‌هاست", html)
        self.assertNotIn("یافته‌های تکمیلی", html)
        self.assertNotIn("پایان همان Word", html)
        self.assertIn("درج امن کامنت در این بخش پشتیبانی نمی‌شود", html)

    def test_incomplete_unified_counts_do_not_invent_per_finding_completion(self):
        stats = self.unified_stats()
        stats["unified_word_review"]["comments"] = 2
        html = html_report(report_data(self.findings(), stats, {"strict": False}))
        self.assertIn("ثبت کامنت و هایلایت همهٔ موارد در همین Word تکمیل نشده", html)
        self.assertNotIn("کامنت کلی/زمینه‌ای روی متن موجود؛ عبارت زردشده لزوماً خطا نیست", html)
        stats["unified_word_review"]["output_created"] = False
        html = html_report(report_data(self.findings(), stats, {"strict": False}))
        self.assertIn("تولید Word یکپارچه تکمیل نشد", html)
        self.assertNotIn("همان Word بازبینی: از 3 مورد", html)

    def test_unified_summary_rejects_metadata_types_and_html_escapes_source_content(self):
        safe = safe_unified_word_review({
            "output_created": "yes", "requested": -2, "source_comments": True,
            "appendix_comments": "1", "comments": 5, "highlighted": 5,
            "exact_comments": 2, "paragraph_comments": 2,
            "path": "secret", "html": "<script>alert(1)</script>"})
        self.assertEqual({"output_created": False, "requested": 0, "comments": 5,
                          "highlighted": 5, "exact_comments": 2, "paragraph_comments": 2}, safe)
        issues = self.findings()
        issues[2].message = '<script>alert("source")</script>'
        html = html_report(report_data(issues, self.unified_stats(), {"strict": False}))
        self.assertIn('&lt;script&gt;alert(&quot;source&quot;)&lt;/script&gt;', html)
        self.assertNotIn('<script>alert("source")</script>', html)

    def test_context_anchor_metadata_is_whitelisted_and_duplicates_do_not_choose_a_location(self):
        stats = self.unified_stats()
        stats["unified_word_review"]["context_records"] += [
            {"index": 4, "anchor_par_index": True, "anchor_xml_path": "C:/private/secret.docx",
             "scope": ["exact"], "source_story": "unknown", "reason": "<script>"},
            {"index": False, "scope": "paragraph"},
            {"index": 2, "anchor_par_index": 99, "scope": "document"},
        ]
        data = report_data(self.findings(), stats, {"strict": False})
        safe = data["stats"]["unified_word_review"]
        self.assertEqual({"index": 4}, safe["context_records"][1])
        self.assertNotIn("context_annotation", data["findings"][2])
        html = html_report(data)
        self.assertNotIn("محل کامنت زمینه‌ای: پاراگراف 100", html)
        self.assertNotIn("C:/private", json.dumps(data))

    def test_historical_appendix_counts_do_not_certify_existing_text_context_comments(self):
        stats = self.unified_stats()
        stats["unified_word_review"].pop("context_comments")
        stats["unified_word_review"]["appendix_comments"] = 1
        data = report_data(self.findings(), stats, {"strict": False})
        html = html_report(data)
        self.assertIn("ثبت کامنت و هایلایت همهٔ موارد در همین Word تکمیل نشده", html)
        self.assertNotIn("کامنت کلی/زمینه‌ای روی متن موجود؛ عبارت زردشده لزوماً خطا نیست", html)
        self.assertNotIn("context_annotation", data["findings"][2])


if __name__ == "__main__":
    unittest.main()
