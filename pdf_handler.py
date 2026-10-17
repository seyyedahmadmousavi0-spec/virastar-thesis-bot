# -*- coding: utf-8 -*-
"""
ماژول پردازش فایل‌های PDF (PDF Text Extractor & Analyzer).
پشتیبانی از استخراج ساختاریافته متن از فایل‌های PDF، تشخیص فایل‌های اسکن‌شده (OCR-needed)،
و اجرای تحلیل کامل نگارشی، علمی و ساختاری روی محتوای متنی PDF.
"""
import os
from typing import Dict, List, Tuple

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None


def extract_pdf_content(pdf_path: str) -> Dict:
    """استخراج متن و صفحات از فایل PDF با تحلیل خوانایی."""
    if pymupdf is None:
        return {
            "ok": False,
            "error": "کتابخانه PyMuPDF نصب نیست.",
            "is_scanned": False,
            "pages": [],
            "total_pages": 0
        }

    try:
        doc = pymupdf.open(pdf_path)
    except Exception as e:
        return {
            "ok": False,
            "error": f"عدم امکان باز کردن فایل PDF: {e}",
            "is_scanned": False,
            "pages": [],
            "total_pages": 0
        }

    import unicodedata

    pages_text = []
    total_chars = 0

    for i, page in enumerate(doc):
        raw_text = page.get_text("text") or ""
        # نرمال‌سازی یونی‌کد (تبدیل Presentation Forms به حروف استاندارد فارسی)
        text = unicodedata.normalize("NFKC", raw_text)
        total_chars += len(text.strip())
        
        # استخراج پاراگراف‌ها با حفظ پیوستگی جملات
        raw_paras = [p.strip() for p in text.split("\n\n") if p.strip()]
        cleaned_paras = []
        for p in raw_paras:
            # یکپارچه‌سازی شکستگی خطوط درون یک پاراگراف
            lines = [ln.strip() for ln in p.splitlines() if ln.strip()]
            p_unified = " ".join(lines)
            if p_unified:
                cleaned_paras.append(p_unified)

        pages_text.append({
            "page_num": i + 1,
            "text": text,
            "paragraphs": cleaned_paras
        })

    # بررسی اسکن‌شده بودن فایل (اگر میانگین حروف هر صفحه کمتر از ۵۰ حرف باشد احتمالاً اسکن/تصویر است)
    is_scanned = (total_chars / max(1, len(doc))) < 50

    return {
        "ok": True,
        "total_pages": len(doc),
        "total_chars": total_chars,
        "is_scanned": is_scanned,
        "pages": pages_text
    }

