"""University output describes observable evidence without certifying a thesis."""
import copy
import json
import unittest

from report_artifacts import (PRODUCT_VERSION, html_report, report_data,
                              safe_university_checklist, safe_university_profile,
                              university_profile_lines)


class UniversityReportTests(unittest.TestCase):
    def setUp(self):
        self.stats = {
            "university_profile": {
                "id": "kashan", "title": "دانشگاه کاشان", "degree": "master",
                "degree_label": "کارشناسی ارشد", "source_status": "user_provided_summary",
                "source_label": "خلاصهٔ ارسالی کاربر", "reference_style": "APA؛ ویرایش تعیین نشده",
                "document_scope": "fragment", "applied_sources": ["B"],
                "excluded_sources": [{"id": "A", "reason": "تعلق راهنما به دانشگاه تأیید نشده است"}],
                "limitations": ["امضا و مهر به بررسی انسانی نیاز دارد"],
                "sources": [{"id": "B", "title": "راهنمای کارشناسی ارشد دانشگاه کاشان",
                             "date": "بهمن ۱۳۹۸", "section": "بخش ۵.۱ متن ارسالی",
                             "kind": "user_provided_summary", "verified": False,
                             "applies": True, "applicable_degrees": ["master"]}],
            },
            "university_checklist": [
                {"id": "FA_ABSTRACT", "label": "چکیدهٔ فارسی", "status": "detected",
                 "detail": "فقط عنوان در متن استخراج‌شده یافت شد", "evidence": "چکیده",
                 "location": "متن اصلی / پاراگراف ۱", "source": "راهنمای کارشناسی ارشد؛ بخش ۵.۱ متن ارسالی"},
                {"id": "SIGNATURE", "label": "امضای تعهدنامه", "status": "manual",
                 "detail": "از روی متن قابل تأیید نیست", "source": "بخش ۵.۵ متن ارسالی"},
            ],
        }

    def test_profile_is_additive_and_not_present_for_general_review(self):
        ordinary = report_data([], {}, {})
        self.assertNotIn("university_profile", ordinary["stats"])
        self.assertNotIn("university_checklist", ordinary["stats"])
        self.assertEqual([], university_profile_lines({}))
        self.assertNotIn('class="notes university-profile"', html_report(ordinary))
        result = report_data([], self.stats, {})
        self.assertEqual("2.0", result["schema_version"])
        self.assertEqual("2.3.0", PRODUCT_VERSION)
        self.assertEqual("master", result["stats"]["university_profile"]["degree"])

    def test_nested_private_configuration_is_never_exported(self):
        profile = self.stats["university_profile"]
        profile["token"] = "PRIVATE_TOKEN"
        profile["local_path"] = "C:/private/policy.pdf"
        profile["sources"][0].update(private_path="C:/private/master.pdf", authorization="secret")
        profile["excluded_sources"][0]["cache_path"] = "C:/private/unverified.pdf"
        self.stats["university_checklist"][0]["internal_metadata"] = {"token": "SECRET"}
        encoded = json.dumps(report_data([], self.stats, {}), ensure_ascii=False)
        for value in ("PRIVATE_TOKEN", "C:/private/", "authorization", "internal_metadata", "cache_path"):
            self.assertNotIn(value, encoded)
        self.assertIn("بهمن ۱۳۹۸", encoded)
        self.assertIn("بخش ۵.۱ متن ارسالی", encoded)
        self.assertNotIn('"page"', encoded)

    def test_html_escapes_every_profile_and_checklist_text_node(self):
        profile = self.stats["university_profile"]
        profile["title"] = '<img src=x onerror="alert(1)">'
        profile["sources"][0]["title"] = "<script>SOURCE_SECRET</script>"
        self.stats["university_checklist"][0].update(label="<svg/onload=alert(1)>",
                                                   evidence="</td><script>EVIDENCE</script>")
        html = html_report(report_data([], self.stats, {}))
        for unsafe in ("<img src=x", "<script>SOURCE_SECRET", "<svg/onload", "<script>EVIDENCE"):
            self.assertNotIn(unsafe, html)
        for escaped in ("&lt;img", "&lt;script&gt;SOURCE_SECRET", "&lt;svg/onload", "&lt;script&gt;EVIDENCE"):
            self.assertIn(escaped, html)
        self.assertIn("connect-src 'none'", html)

    def test_checklist_does_not_turn_detected_text_into_approval(self):
        data = report_data([], self.stats, {})
        html = html_report(data)
        text = "\n".join(university_profile_lines(self.stats))
        for output in (html, text):
            self.assertIn("نشانهٔ متنی یافت شد؛ تأیید قالب نیست", output)
            self.assertIn("اعتبار رسمی و به‌روز بودن مستقلاً تأیید نشده", output)
            self.assertIn("با اطلاعات موجود قابل تأیید نیست", output)
            self.assertIn("شناسه‌های منبع با سطح اطمینان A/B/C ایرادها متفاوت‌اند", output)
            self.assertIn("APA؛ ویرایش تعیین نشده", output)
            self.assertNotIn("APA 7", output)
        self.assertEqual({}, data["summary"]["by_category"])
        self.assertIsNone(data["summary"]["by_confidence"]["C"])

    def test_unknown_or_malformed_status_is_manual_and_input_is_unchanged(self):
        self.stats["university_profile"].update(degree={"secret": 1}, document_scope=[],
                                                applied_sources=["B", "C:/secret"],
                                                source_status="official_approval")
        self.stats["university_profile"]["sources"][0]["applicable_degrees"] = ["master", {}, "other"]
        self.stats["university_checklist"][0]["status"] = {"approved": True}
        before = copy.deepcopy(self.stats)
        data = report_data([], self.stats, {})
        self.assertEqual(before, self.stats)
        safe = data["stats"]["university_profile"]
        self.assertEqual("unknown", safe["degree"])
        self.assertEqual("fragment", safe["document_scope"])
        self.assertEqual(["B"], safe["applied_sources"])
        self.assertEqual(["master"], safe["sources"][0]["applicable_degrees"])
        self.assertEqual("user_provided_summary", safe["source_status"])
        self.assertEqual("manual", data["stats"]["university_checklist"][0]["status"])
        self.assertIsNone(safe_university_profile({"id": "other"}))
        self.assertEqual([], safe_university_checklist({"bad": True}))

    def test_uncertain_bibliography_result_survives_export_and_html_escaping(self):
        self.stats["university_reference_checks"] = {
            "status": "partial", "recognized_entries": 1, "alphabetical_checked": False,
            "style": "APA", "edition": None, "online_verified": False,
            "detail": "نام ناشناخته یا مدخل چندپاراگرافی <script>raw</script>؛ ترتیب قابل تأیید نیست.",
            "private_path": "C:/private/references.txt", "token": "REFERENCE_SECRET",
            "parser_state": {"auth": "PRIVATE_AUTH"},
        }
        data = report_data([], self.stats, {})
        checks = data["stats"]["university_reference_checks"]
        self.assertEqual(1, checks["recognized_entries"])
        self.assertFalse(checks["alphabetical_checked"])
        self.assertIsNone(checks["edition"])
        encoded = json.dumps(data, ensure_ascii=False)
        for secret in ("C:/private/", "REFERENCE_SECRET", "PRIVATE_AUTH", "parser_state"):
            self.assertNotIn(secret, encoded)
        text = "\n".join(university_profile_lines(self.stats))
        html = html_report(data)
        self.assertIn("نام ناشناخته یا مدخل چندپاراگرافی", text)
        self.assertIn("&lt;script&gt;raw&lt;/script&gt;", html)
        self.assertNotIn("<script>raw</script>", html)
        for output in (text, html):
            self.assertIn("ترتیب الفبایی: قابل بررسی مطمئن نبود؛ تأیید نشده است", output)
            self.assertIn("انطباق کامل با APA یا صحت اطلاعات منابع را تأیید نمی‌کند", output)
            self.assertIn("راستی‌آزمایی آنلاین منابع انجام نشده است", output)

    def test_unrecognized_completion_claim_does_not_certify_apa(self):
        self.stats["university_reference_checks"] = {
            "status": "complete", "recognized_entries": True,
            "detail": "دادهٔ کافی برای تشخیص مدخل‌ها نیست.",
        }
        data = report_data([], self.stats, {})
        self.assertEqual("manual", data["stats"]["university_reference_checks"]["status"])
        self.assertNotIn("recognized_entries", data["stats"]["university_reference_checks"])
        for output in (html_report(data), "\n".join(university_profile_lines(self.stats))):
            self.assertIn("وضعیت بررسی منابع: با اطلاعات موجود قابل تأیید نیست", output)
            self.assertIn("ویرایش APA تعیین نشده است", output)


if __name__ == "__main__":
    unittest.main()
