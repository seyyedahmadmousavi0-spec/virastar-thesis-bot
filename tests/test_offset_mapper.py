# -*- coding: utf-8 -*-
"""
تست نگاشتار آفست و نمایش تحلیلی متن (Offset Mapper Tests).
راستی‌آزمایی انطباق دقیق موقعیت کاراکترهای متن نرمال‌شده با متن اصلی دست‌نخورده.
"""
import unittest
from analysis_model import OffsetMapper


class TestOffsetMapper(unittest.TestCase):

    def test_identity_mapping(self):
        """تست نگاشت متن استاندارد بدون تغییر طول."""
        text = "این یک متن فارسی استاندارد است."
        mapper = OffsetMapper.create_normalized(text)
        self.assertEqual(mapper.normalized_text, text)
        s, e, ok = mapper.map_slice_to_original(4, 6)
        self.assertTrue(ok)
        self.assertEqual(text[s:e], "یک")

    def test_arabic_characters_normalization_mapping(self):
        """تست نگاشت هنگام تبدیل ي و ك به ی و ک بدون جابجایی ایندکس."""
        orig = "تحقيق و بررسي كتاب‌هاي علمي"
        mapper = OffsetMapper.create_normalized(orig)
        self.assertIn("تحقیق", mapper.normalized_text)
        self.assertIn("بررسی", mapper.normalized_text)
        self.assertIn("کتاب", mapper.normalized_text)

        # استخراج کلمه بررسی از روی آفست متن نرمال‌شده
        pos_norm = mapper.normalized_text.find("بررسی")
        s, e, ok = mapper.map_slice_to_original(pos_norm, pos_norm + len("بررسی"))
        self.assertTrue(ok)
        self.assertEqual(orig[s:e], "بررسي")


if __name__ == "__main__":
    unittest.main()
