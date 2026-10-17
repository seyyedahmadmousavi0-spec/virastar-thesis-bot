# -*- coding: utf-8 -*-
"""
ماژول ارزیابی تخصصی بر اساس شیوه نامه دانشگاه علوم پزشکی تهران (TUMS Thesis Guideline Checker).
منطبق بر فایل ۴۹ صفحه‌ای «دستورالعمل نگارش پایان‌نامه مقاطع تحصیلات تکمیلی دانشگاه علوم پزشکی تهران».
"""
import json
import os
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple

HERE = os.path.dirname(os.path.abspath(__file__))

@dataclass
class TUMSIssue:
    rule: str
    severity: str  # error, warn, info
    message: str
    loc: str = ""
    snippet: str = ""
    ref_page: str = ""


def load_tums_profile():
    path = os.path.join(HERE, "profiles", "tums.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def check_tums_checklist(doc_paragraphs) -> Tuple[List[Dict], List[TUMSIssue]]:
    """بررسی ۲۷ جزء الزامی و پیشنهادی طبق چک‌لیست صفحات ۴۵ تا ۴۹ دستورالعمل دانشگاه علوم پزشکی تهران."""
    profile = load_tums_profile()
    checklist = profile.get("checklist_items", [])
    issues = []
    status_list = []

    full_text_lower = " ".join(p.text.lower() for p in doc_paragraphs)

    for item in checklist:
        name = item["name"]
        item_id = item["id"]
        patterns = item.get("patterns", [])
        is_required = item.get("required", True)
        manual = item.get("manual_check", False)

        if manual:
            status = "نیازمند بررسی انسانی (نسخه چاپی/فیزیکی)"
            status_list.append({"name": name, "status": status, "severity": "info"})
            continue

        found = False
        for pat in patterns:
            if pat.lower() in full_text_lower:
                found = True
                break

        if found:
            status = "موجود ✅"
            status_list.append({"name": name, "status": status, "severity": "ok"})
        else:
            if is_required:
                status = "یافت نشد ❌"
                issues.append(TUMSIssue(
                    rule="tums_checklist_missing",
                    severity="error",
                    message=f"بخش الزامی «{name}» در پایان‌نامه یافت نشد (طبق پیوست‌های شیوه نامه دانشگاه علوم پزشکی تهران).",
                    ref_page="صفحات ۸ و ۴۵ تا ۴۹ دستورالعمل"
                ))
            else:
                status = "اختیاری / یافت نشد ℹ️"
            status_list.append({"name": name, "status": status, "severity": "warn" if is_required else "info"})

    return status_list, issues


def check_tums_chapter_proportions(chapters_words: Dict[int, int], total_words: int) -> List[TUMSIssue]:
    """
    بررسی قاعده صریح صفحه ۷ دستورالعمل:
    «فصول ۱ و ۲ باید حداکثر ۳۰٪ از کل متن پایان‌نامه باشد.»
    """
    issues = []
    if total_words < 1000:
        return issues

    ch1_words = chapters_words.get(1, 0)
    ch2_words = chapters_words.get(2, 0)
    intro_review_words = ch1_words + ch2_words

    share = intro_review_words / total_words
    if share > 0.30:
        pct = int(share * 100)
        issues.append(TUMSIssue(
            rule="tums_chapter_proportions",
            severity="warn",
            message=f"حجم فصول ۱ و ۲ برابر با {pct}٪ کل پایان‌نامه است؛ طبق صفحه ۷ دستورالعمل دانشگاه علوم پزشکی تهران، فصول ۱ و ۲ (مقدمه و بررسی متون) حداکثر باید ۳۰٪ از کل متن باشند.",
            ref_page="صفحه ۷ دستورالعمل"
        ))

    return issues


def check_tums_references(refs: List[str]) -> List[TUMSIssue]:
    """
    بررسی قواعد استناد دانشگاه علوم پزشکی تهران (صفحات ۲۲ و ۲۳ دستورالعمل):
    - ممنوعیت استفاده از «همکاران» و «et al» در بخش نویسندگان فهرست منابع
    - درج «ص» در منابع فارسی و «pp.» در منابع انگلیسی
    """
    issues = []
    for i, ref in enumerate(refs, 1):
        # بررسی ممنوعیت واژه‌های همکاران و et al در نویسندگان
        if "همکاران" in ref or re.search(r"\bet\s+al\b", ref, re.IGNORECASE):
            issues.append(TUMSIssue(
                rule="tums_et_al_forbidden",
                severity="error",
                message=f"منبع {i}: استفاده از واژه‌های «همکاران» و «et al» در فهرست نویسندگان مراجع طبق صفحه ۲۳ شیوه نامه مجاز نیست؛ اسامی نویسندگان باید کامل درج شود.",
                loc=f"منبع شماره {i}",
                snippet=ref[:70] + "…",
                ref_page="صفحه ۲۳ دستورالعمل"
            ))

        # بررسی ذکر ص یا pp برای شماره صفحات
        has_page = bool(re.search(r"(\bص\s*[۰-۹0-9]+|\bpp?\.\s*[0-9]+)", ref, re.IGNORECASE))
        if not has_page and len(ref.split()) > 7:
            issues.append(TUMSIssue(
                rule="tums_ref_page_format",
                severity="info",
                message=f"منبع {i}: شماره صفحات منبع طبق الگو (پس از «ص» در فارسی یا «pp.» در انگلیسی) مشخص نشده است.",
                loc=f"منبع شماره {i}",
                snippet=ref[:60] + "…",
                ref_page="صفحه ۲۳ دستورالعمل"
            ))

    return issues


def check_tums_margins(sections) -> List[TUMSIssue]:
    """
    بررسی ضوابط چاپ دورو و حاشیه‌های دانشگاه علوم پزشکی تهران (صفحه ۹ دستورالعمل):
    - حاشیه صفحات فرد: راست ۳.۵cm، چپ ۲.۵cm، بالا ۲.۵cm، پایین ۲.۵cm
    - حاشیه صفحات زوج: راست ۲.۵cm، چپ ۳.۵cm، بالا ۲.۵cm، پایین ۲.۵cm
    - حاشیه صفحات عنوان: بالا ۱۰cm، پایین ۱۰cm، راست ۳.۵cm، چپ ۲.۵cm
    """
    issues = []
    for i, s in enumerate(sections, 1):
        if not hasattr(s, "top_margin") or s.top_margin is None:
            continue
        top = s.top_margin / 360000
        bot = s.bottom_margin / 360000
        right = s.right_margin / 360000
        left = s.left_margin / 360000

        # بررسی کلی محدوده حاشیه‌ها
        if abs(top - 2.5) > 0.4 and abs(top - 10.0) > 0.5:
            issues.append(TUMSIssue(
                rule="tums_page_margin",
                severity="warn",
                message=f"بخش {i}: حاشیه بالا {top:.1f}cm است؛ طبق دستورالعمل TUMS باید ۲.۵cm (یا در صفحات عنوان فصل ۱۰cm) باشد.",
                ref_page="صفحه ۹ دستورالعمل"
            ))
        if abs(bot - 2.5) > 0.4 and abs(bot - 10.0) > 0.5:
            issues.append(TUMSIssue(
                rule="tums_page_margin",
                severity="warn",
                message=f"بخش {i}: حاشیه پایین {bot:.1f}cm است؛ طبق دستورالعمل TUMS باید ۲.۵cm باشد.",
                ref_page="صفحه ۹ دستورالعمل"
            ))

    return issues

