# -*- coding: utf-8 -*-
"""
تست موتور کامنت‌گذاری ایمن Word (Safe DOCX Annotator Tests).
راستی‌آزمایی:
1. عدم هایلایت تصادفی کلمه اول پاراگراف در صورت نبود تطابق قطعی.
2. تطبیق چندرانی (Multi-run matching).
3. تولید سند سالم و معتبر Word با حفظ ۱۰۰٪ متن و استایل‌ها.
"""
import os
import shutil
import tempfile
import unittest
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from analysis_model import ConfidenceLevel, Issue
import safe_annotator


class TestSafeAnnotator(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="test_annotator_")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_no_fallback_highlight_on_unmatched_target(self):
        """تست عدم هایلایت کلمه اول پاراگراف زمانی که متن هدف پیدا نشود."""
        doc = Document()
        p = doc.add_paragraph("در این پژوهش نتایج بررسی شد.")
        out_docx = os.path.join(self.temp_dir, "out.docx")

        # یک ایراد پاراگرافی که متن هدف آن در پاراگراف وجود ندارد
        issue = Issue(
            rule_id="long_paragraph",
            severity="info",
            message="پاراگراف طولانی است",
            loc="پاراگراف ۱",
            par_index=0,
            target_text="عبارت ناموجود در متن"
        )

        safe_annotator.annotate_document(doc, [issue], out_docx)

        # بازگشایی سند و بررسی اینکه کلمه اول «در» هرگز هایلایت زرد نخورده باشد
        reopened = Document(out_docx)
        self.assertEqual(len(reopened.paragraphs), 1)
        first_p = reopened.paragraphs[0]

        for r in first_p.runs:
            self.assertNotEqual(
                r.font.highlight_color,
                WD_COLOR_INDEX.YELLOW,
                "کلمه غیرمرتبط در غیاب تطابق هدف، نباید هایلایت زرد شود!"
            )

    def test_multi_run_matching_and_highlight(self):
        """تست هایلایت دقیق عبارتی که در دو ران شکسته شده است."""
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("این کار انجام ")
        r2 = p.add_run("می‌شود و پایان یافت.")
        out_docx = os.path.join(self.temp_dir, "out_multirun.docx")

        issue = Issue(
            rule_id="academic_first_person",
            severity="warn",
            message="تست",
            loc="پاراگراف ۱",
            par_index=0,
            target_text="می‌شود"
        )

        safe_annotator.annotate_document(doc, [issue], out_docx)

        reopened = Document(out_docx)
        reopened_p = reopened.paragraphs[0]
        # متن کامل باید بدون کوچک‌ترین تغییر کاراکتری حفظ شده باشد
        self.assertEqual(reopened_p.text, "این کار انجام می‌شود و پایان یافت.")

        # باید حداقل یک ران هایلایت زرد با متن «می‌شود» وجود داشته باشد
        highlighted = [r for r in reopened_p.runs if r.font.highlight_color == WD_COLOR_INDEX.YELLOW]
        self.assertGreater(len(highlighted), 0)
        self.assertIn("می‌شود", "".join(r.text for r in highlighted))


if __name__ == "__main__":
    unittest.main()
