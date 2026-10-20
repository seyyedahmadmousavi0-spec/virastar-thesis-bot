# -*- coding: utf-8 -*-
"""
ماژول ارزیابی تخصصی بر اساس شیوه نامه دانشگاه علوم پزشکی تهران (TUMS Thesis Guideline Checker).
منطبق بر فایل ۴۹ صفحه‌ای «دستورالعمل نگارش پایان‌نامه مقاطع تحصیلات تکمیلی دانشگاه علوم پزشکی تهران».
شامل:
- چک‌لیست ۲۷ جزء الزامی پایان‌نامه (صفحات ۴۵ تا ۴۹)
- نسبت و توازن فصول ۱ و ۲ (حداکثر ۳۰٪ حجم کل)
- استانداردهای فهرست منابع و استناد (ممنوعیت et al و همکاران)
- حاشیه‌ها و تنظیمات صفحه‌آرایی تخصصی
"""
import json
import os
import re
from typing import Dict, List, Tuple

from analysis_model import ConfidenceLevel, Issue

HERE = os.path.dirname(os.path.abspath(__file__))


def load_tums_profile():
    path = os.path.join(HERE, "profiles", "tums.json")
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def check_tums_checklist(doc_paragraphs) -> Tuple[List[Dict], List[Issue]]:
    """بررسی ۲۷ جزء الزامی و پیشنهادی طبق چک‌لیست صفحات ۴۵ تا ۴۹ دستورالعمل دانشگاه علوم پزشکی تهران."""
    profile = load_tums_profile()
    checklist = profile.get("checklist_items", [])
    issues: List[Issue] = []
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
                issues.append(Issue(
                    rule_id="TUMS_CHECKLIST_MISSING",
                    category="tums_guideline",
                    severity="warn",
                    confidence=ConfidenceLevel.REVIEW_B.value,
                    message=f"بخش الزامی «{name}» در پایان‌نامه یافت نشد (طبق پیوست‌های شیوه نامه دانشگاه علوم پزشکی تهران، ص ۴۵-۴۹).",
                    loc="چک‌لیست TUMS"
                ))
            else:
                status = "اختیاری / یافت نشد ℹ️"
            status_list.append({"name": name, "status": status, "severity": "warn" if is_required else "info"})

    return status_list, issues


def check_tums_chapter_proportions(chapters_words: Dict[int, int], total_words: int) -> List[Issue]:
    """
    بررسی قاعده صریح صفحه ۷ دستورالعمل:
    «فصول ۱ و ۲ باید حداکثر ۳۰٪ از کل متن پایان‌نامه باشد.»
    """
    issues: List[Issue] = []
    if total_words < 1000:
        return issues

    ch1_words = chapters_words.get(1, 0)
    ch2_words = chapters_words.get(2, 0)
    intro_review_words = ch1_words + ch2_words

    share = intro_review_words / total_words
    maximum = load_tums_profile().get("tums_special_rules", {}).get("max_intro_review_pct", 30)
    if not isinstance(maximum, (int, float)) or not 0 < maximum <= 100:
        return issues
    if share > maximum / 100:
        pct = int(share * 100)
        issues.append(Issue(
            rule_id="TUMS_CHAPTER_PROPORTIONS",
            category="tums_guideline",
            severity="warn",
            confidence=ConfidenceLevel.REVIEW_B.value,
            message=f"حجم فصول ۱ و ۲ برابر با {pct}٪ کل پایان‌نامه است؛ طبق صفحه ۷ دستورالعمل دانشگاه علوم پزشکی تهران، مجموع فصول ۱ و ۲ حداکثر مقدار تنظیم‌شده {maximum:g}٪ است؛ نسبت بر مبنای شمارش واژه برآورد شده و نیازمند بررسی انسانی است.",
            loc="ساختار فصول"
        ))

    return issues


def check_tums_references(refs: List[str]) -> List[Issue]:
    """
    بررسی قواعد استناد دانشگاه علوم پزشکی تهران (صفحات ۲۲ و ۲۳ دستورالعمل):
    - ممنوعیت استفاده از «همکاران» و «et al» در بخش نویسندگان فهرست منابع
    - درج «ص» در منابع فارسی و «pp.» در منابع انگلیسی
    """
    issues: List[Issue] = []
    for i, ref in enumerate(refs, 1):
        # بررسی ممنوعیت واژه‌های همکاران و et al در نویسندگان
        if "همکاران" in ref or re.search(r"\bet\s+al\b", ref, re.IGNORECASE):
            issues.append(Issue(
                rule_id="TUMS_ET_AL_FORBIDDEN",
                category="tums_guideline",
                severity="warn",
                confidence=ConfidenceLevel.REVIEW_B.value,
                message=f"منبع {i}: استفاده از واژه‌های «همکاران» و «et al» در فهرست نویسندگان مراجع طبق صفحه ۲۳ شیوه نامه مجاز نیست؛ اسامی تمامی نویسندگان باید درج شود.",
                loc=f"منبع شماره {i}",
                snippet=ref[:70] + "…",
                target_text="همکاران" if "همکاران" in ref else "et al"
            ))

        # بررسی ذکر ص یا pp برای شماره صفحات
        has_page = bool(re.search(r"(\bص\s*[۰-۹0-9]+|\bpp?\.\s*[0-9]+)", ref, re.IGNORECASE))
        if not has_page and len(ref.split()) > 7:
            issues.append(Issue(
                rule_id="TUMS_REF_PAGE_FORMAT",
                category="tums_guideline",
                severity="info",
                confidence=ConfidenceLevel.REVIEW_B.value,
                message=f"منبع {i}: شماره صفحات منبع طبق الگو (پس از «ص» در فارسی یا «pp.» در انگلیسی) مشخص نشده است (صفحه ۲۳ دستورالعمل).",
                loc=f"منبع شماره {i}",
                snippet=ref[:60] + "…"
            ))

    return issues


def check_tums_margins(sections) -> List[Issue]:
    """Compare declared section margins with all compatible local profile variants."""
    profile = load_tums_profile()
    cfg = profile.get("page", {})
    variants = [cfg.get(name, {}) for name in ("margins_odd_cm", "margins_even_cm", "margins_chapter_title_cm")]
    tolerance = cfg.get("margin_tolerance_cm", 0.3)
    issues = []
    for index, section in enumerate(sections, 1):
        for side in ("top", "bottom", "left", "right"):
            value = getattr(section, side + "_margin", None)
            permitted = [v[side] for v in variants if side in v]
            if value is None or not permitted:
                continue
            measured = value / 360000
            if all(abs(measured - target) > tolerance for target in permitted):
                issues.append(Issue("TUMS_PAGE_MARGIN", "warn",
                    f"Section {index}: {side} margin is {measured:.2f} cm; local profile candidates: {permitted}. Review section/page role.",
                    f"Section {index}", category="tums_guideline", confidence="B",
                    rule_version=profile.get("profile_version", "1")))
    return issues
