# -*- coding: utf-8 -*-
"""
آزمون‌های جامع دقت و پوشش قواعد (Positive, Negative, Boundary Rule Tests).
شامل:
- Positive Cases: خطاهای واقعی که قطعا باید کشف شوند.
- Negative Cases: متون معتبر علمی که هرگز نباید اخطار کاذب ایجاد کنند.
- Boundary Cases: موارد مرزی، تهی، رشته‌های بسیار طولانی، نمادها.
"""
import unittest
import checker
from analysis_model import ConfidenceLevel


class TestRulesPrecision(unittest.TestCase):

    def setUp(self):
        self.rules = checker.load_rules(mode="thesis")

    # ------------------ ۱. موارد مثبت (True Positives) ------------------
    def test_positive_mi_prefix_with_space(self):
        """کشف پیشوند «می» جدا شده با فاصله معمولی."""
        text = "او به کتابخانه می رود تا کتاب بخواند."
        issues = checker.check_text(0, text, "تست", self.rules)
        mi_issues = [i for i in issues if i.rule_id == "mi_prefix"]
        self.assertGreater(len(mi_issues), 0, "قاعده mi_prefix موفق به کشف «می رود» نشد.")

    def test_positive_arabic_letters(self):
        """کشف ي و ك عربی به جای ی و ک فارسی."""
        text = "اين كتاب بسيار مفيد است."
        issues = checker.check_text(0, text, "تست", self.rules)
        ar_issues = [i for i in issues if i.rule_id == "arabic_letters"]
        self.assertGreater(len(ar_issues), 0, "قاعده arabic_letters حروف ي و ك را تشخیص نداد.")

    def test_positive_common_mistakes(self):
        """کشف غلط‌های املایی رایج مانند «بخاطر» و «بدلیل»."""
        text = "پژوهشگر بخاطر کمبود داده نتوانست تحلیل کند."
        issues = checker.check_text(0, text, "تست", self.rules)
        mistake_issues = [i for i in issues if i.rule_id == "common_mistake"]
        self.assertGreater(len(mistake_issues), 0, "غلط رایج «بخاطر» کشف نشد.")
        self.assertEqual(mistake_issues[0].suggestion, "به خاطر")

    def test_positive_plural_space(self):
        """کشف فاصله معمولی به جای نیم‌فاصله در «ها»."""
        text = "تمام داده ها به دقت جمع‌آوری شدند."
        issues = checker.check_text(0, text, "تست", self.rules)
        plural_issues = [i for i in issues if i.rule_id == "plural_space"]
        self.assertGreater(len(plural_issues), 0, "علامت جمع «ها» با فاصله عادی کشف نشد.")

    def test_positive_english_article_and_preposition(self):
        """کشف خطای حروف تعریف و حروف اضافه در متن انگلیسی."""
        text = "This is a experimental study and researchers are interested on this method."
        issues = checker.check_text(0, text, "تست", self.rules)
        article_issues = [i for i in issues if i.rule_id == "EN_ARTICLE"]
        prep_issues = [i for i in issues if i.rule_id == "EN_PREP"]
        self.assertGreater(len(article_issues), 0, "خطای 'a experimental' کشف نشد.")
        self.assertGreater(len(prep_issues), 0, "خطای 'interested on' کشف نشد.")

    # ------------------ ۲. موارد منفی (True Negatives - بدون هشدار کاذب) ------------------
    def test_negative_valid_persian_scientific_text(self):
        """متن کاملاً صحیح فارسی با نیم‌فاصله‌های استاندارد نباید هیچ خطای قطعی تولید کند."""
        text = "پژوهشگر در این بررسی، داده‌ها را به دقت تحلیل می‌کند و نتایج به‌دست‌آمده نشان‌دهنده اثربخشی روش است."
        issues = checker.check_text(0, text, "تست", self.rules)
        errors = [i for i in issues if i.severity == "error" and i.confidence == ConfidenceLevel.HIGH_A.value]
        self.assertEqual(len(errors), 0, f"هشدار کاذب در متن صحیح: {errors}")

    def test_negative_scientific_units_no_false_latin_digits(self):
        """متن حاوی اعداد و یکاهای علمی نباید هشدار رقم لاتین تولید کند."""
        text = "در این آزمایش، 25 mg از ماده در 100 mL آب مقطر حل شد و در دمای 37 °C به مدت 2 hr نگهداری گردید."
        issues = checker.check_text(0, text, "تست", self.rules)
        latin_digit_issues = [i for i in issues if i.rule_id == "latin_digits"]
        self.assertEqual(len(latin_digit_issues), 0, f"هشدار کاذب روی یکاهای علمی: {latin_digit_issues}")

    def test_negative_reduplication_allowed(self):
        """واژگان ترکیبی مکرر مجاز (کم‌کم، ذره‌ذره) نباید خطای تکرار ناخواسته بگیرند."""
        text = "غلظت محلول کم کم افزایش یافت و نمونه ها ذره ذره رسوب کردند."
        issues = checker.check_text(0, text, "تست", self.rules)
        # فقط باید به شکل info با پیشنهاد نیم‌فاصله ظاهر شود، نه اخطار تکرار ناخواسته
        warn_repeated = [i for i in issues if i.rule_id == "repeated_word" and i.severity == "warn"]
        self.assertEqual(len(warn_repeated), 0, f"هشدار اشتباه روی تکرار مجاز: {warn_repeated}")

    # ------------------ ۳. موارد مرزی (Boundary Cases) ------------------
    def test_boundary_empty_and_whitespace(self):
        """رشته خالی یا فقط فاصله نباید هیچ مشکلی ایجاد کند یا کرش کند."""
        self.assertEqual(checker.check_text(0, "", "تست", self.rules), [])
        self.assertEqual(checker.check_text(0, "   \n\t  ", "تست", self.rules), [])

    def test_boundary_extreme_lengths(self):
        """پاراگراف بسیار طولانی باید بدون وقفه یا کرش پردازش شود."""
        long_text = "این یک جمله تستی استاندارد است. " * 150
        issues = checker.check_text(0, long_text, "تست", self.rules)
        # باید پاراگراف طولانی را به درستی تشخیص دهد
        long_par_issues = [i for i in issues if i.rule_id == "long_paragraph"]
        self.assertGreater(len(long_par_issues), 0)


if __name__ == "__main__":
    unittest.main()
