# -*- coding: utf-8 -*-
"""
تست لایه حفاظت از محتوای علمی (Protected Scientific Content Tests).
راستی‌آزمایی مصونیت فرمول‌های ریاضی و شیمیایی، یکاهای اندازه‌گیری (SI)،
اصطلاحات لاتین علمی، ارقام آماری، شناسه‌های DOI، پیوندها و ارجاعات متنی.
"""
import unittest
from protected_content import ProtectedContentDetector
import checker


class TestProtectedContent(unittest.TestCase):

    def test_scientific_units_protection(self):
        """آزمون یکاهای اندازه‌گیری و عدم گزارش ارقام لاتین یا فاصله‌گذاری."""
        sample_texts = [
            "غلظت نهایی پروتئین 15.5 mg/dL و دمای واکنش 25 °C تنظیم گردید.",
            "محلول دارای 100 µL آب مقطر و 50 mM نمک طعام بود.",
            "فرکانس نمونه‌برداری 4.2 GHz و ولتاژ اعمالی 220 V است.",
            "وزن نمونه 70 kg و طول آن 50 µm اندازه‌گیری شد.",
            "میزان بازدهی واکنش 95% و با خطای 1.2% گزارش شد.",
            "میزان کربن دی‌اکسید 400 ppm اندازه گرفته شد."
        ]
        for t in sample_texts:
            issues = checker.check_text(0, t, "تست", {})
            # نباید هیچ خطایی برای رقم لاتین (latin_digits) در این عبارات صادر شود
            latin_digit_issues = [i for i in issues if i.rule_id == "latin_digits"]
            self.assertEqual(len(latin_digit_issues), 0, f"خطای کاذب رقم لاتین روی واحد علمی: {t} -> {latin_digit_issues}")

    def test_chemical_formulas_protection(self):
        """آزمون فرمول‌های شیمیایی و زیست‌مولکولی."""
        sample = "ترکیب H2O و NaCl در کنار CO2 و اسید H2SO4 در بافر سنتز شد."
        issues = checker.check_text(0, sample, "تست", {})
        latin_digit_issues = [i for i in issues if i.rule_id == "latin_digits"]
        self.assertEqual(len(latin_digit_issues), 0, f"خطای کاذب رقم لاتین روی فرمول شیمیایی: {latin_digit_issues}")

    def test_statistical_expressions_protection(self):
        """آزمون شاخص‌های آماری مانند p-value و فاصله‌های اطمینان."""
        sample = "تفاوت معنادار آماری مشاهده شد (p < 0.05 و 95% CI [1.2, 3.4])."
        issues = checker.check_text(0, sample, "تست", {})
        # علائم < و اعداد آماری نباید خطا تولید کنند
        punct_issues = [i for i in issues if i.rule_id in ("latin_digits", "space_before_punct")]
        self.assertEqual(len(punct_issues), 0, f"خطای کاذب روی شاخص آماری: {punct_issues}")

    def test_academic_latin_and_citations(self):
        """آزمون اصطلاحات لاتین علمی و ارجاعات متنی."""
        sample = "بر اساس یافته‌های محققان قبلی (Smith et al., 2020) آزمایش‌ها به صورت in vitro انجام شد [1-3]."
        issues = checker.check_text(0, sample, "تست", {})
        punct_issues = [i for i in issues if i.rule_id in ("no_space_after_punct", "latin_punct")]
        self.assertEqual(len(punct_issues), 0, f"خطای کاذب روی ارجاع و عبارات لاتین: {punct_issues}")

    def test_doi_and_urls(self):
        """آزمون آدرس‌های وب و DOI."""
        sample = "اطلاعات بیشتر در https://doi.org/10.1016/j.cell.2021.05.001 یا www.nature.com در دسترس است."
        issues = checker.check_text(0, sample, "تست", {})
        punct_issues = [i for i in issues if i.rule_id in ("no_space_after_punct", "latin_digits")]
        self.assertEqual(len(punct_issues), 0, f"خطای کاذب روی DOI و URL: {punct_issues}")


if __name__ == "__main__":
    unittest.main()
