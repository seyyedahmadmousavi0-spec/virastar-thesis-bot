# -*- coding: utf-8 -*-
"""
موتور جامع ارزیابی و ویرایش علمی، نگارشی و ساختاری پایان‌نامه و مقاله (فایل Word، PDF و متن).
قاعده‌محور، دقیق و قابل گسترش؛ منطبق بر شیوه نامه‌های دانشگاهی و بین‌المللی.
شامل پشتیبانی از شیوه‌نامه دانشگاه علوم پزشکی تهران (TUMS).
"""
import argparse
import copy
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

import academic_checker
import en_checker
import pdf_handler
import tums_checker

HERE = os.path.dirname(os.path.abspath(__file__))
ZWNJ = "\u200c"

# حروف فارسی/عربی
FA = "\u0621-\u063A\u0641-\u064A\u067E\u0686\u0698\u06A9\u06AF\u06CC\u06C0"
FA_RE = re.compile("[" + FA + "]")
LATIN_RE = re.compile("[A-Za-z]")
DIGIT_MAP = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

SEVERITY_ORDER = {"error": 0, "warn": 1, "info": 2}
SEVERITY_ICON = {"error": "🔴", "warn": "🟠", "info": "🔵"}

RULE_CATEGORIES = {
    # ۱. املایی و تایپوگرافی
    "arabic_letters": "۱. غلط املایی و تایپوگرافی",
    "arabic_digits": "۱. غلط املایی و تایپوگرافی",
    "latin_digits": "۱. غلط املایی و تایپوگرافی",
    "mi_prefix": "۱. غلط املایی و تایپوگرافی",
    "plural_space": "۱. غلط املایی و تایپوگرافی",
    "comparative_space": "۱. غلط املایی و تایپوگرافی",
    "ye_space": "۱. غلط املایی و تایپوگرافی",
    "common_mistake": "۱. غلط املایی و تایپوگرافی",
    "repeated_word": "۱. غلط املایی و تایپوگرافی",

    # ۲. دستور زبان و نگارش (فارسی و انگلیسی)
    "en_grammar_agreement": "۲. دستور زبان و ساختار جمله",
    "en_article": "۲. دستور زبان و ساختار جمله",
    "en_preposition": "۲. دستور زبان و ساختار جمله",

    # ۳. نشانه‌گذاری و فاصله‌گذاری
    "space_before_punct": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "no_space_after_punct": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "latin_punct": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "paren_space": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "straight_quotes": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "double_space": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "leading_space": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "blank_lines": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "en_punctuation": "۳. نشانه‌گذاری و فاصله‌گذاری",
    "en_unpaired_brackets": "۳. نشانه‌گذاری و فاصله‌گذاری",

    # ۴. سبک و لحن علمی
    "informal_word": "۴. سبک و لحن علمی",
    "first_person": "۴. سبک و لحن علمی",
    "academic_tone_emotional": "۴. سبک و لحن علمی",
    "academic_pleonasm": "۴. سبک و لحن علمی",
    "academic_first_person": "۴. سبک و لحن علمی",
    "en_academic_tone": "۴. سبک و لحن علمی",
    "en_wordiness": "۴. سبک و لحن علمی",

    # ۵. انسجام و ساختار فصول
    "long_paragraph": "۵. انسجام و ساختار فصول",
    "long_sentence": "۵. انسجام و ساختار فصول",
    "academic_cohesion": "۵. انسجام و ساختار فصول",
    "missing_section": "۵. انسجام و ساختار فصول",
    "recommended_section": "۵. انسجام و ساختار فصول",
    "chapters": "۵. انسجام و ساختار فصول",
    "heading_style": "۵. انسجام و ساختار فصول",
    "toc": "۵. انسجام و ساختار فصول",
    "abstract_length": "۵. انسجام و ساختار فصول",
    "keywords_count": "۵. انسجام و ساختار فصول",
    "tums_chapter_proportions": "۵. انسجام و ساختار فصول",

    # ۶. قالب‌بندی فایل و شیوه دانشگاه
    "page_size": "۶. قالب‌بندی و شیوه دانشگاه",
    "page_margin": "۶. قالب‌بندی و شیوه دانشگاه",
    "font_fa": "۶. قالب‌بندی و شیوه دانشگاه",
    "size_fa": "۶. قالب‌بندی و شیوه دانشگاه",
    "font_en": "۶. قالب‌بندی و شیوه دانشگاه",
    "size_en": "۶. قالب‌بندی و شیوه دانشگاه",
    "line_spacing": "۶. قالب‌بندی و شیوه دانشگاه",
    "alignment": "۶. قالب‌بندی و شیوه دانشگاه",
    "rtl": "۶. قالب‌بندی و شیوه دانشگاه",
    "page_numbers": "۶. قالب‌بندی و شیوه دانشگاه",
    "tums_page_margin": "۶. قالب‌بندی و شیوه دانشگاه",

    # ۷. جدول‌ها، شکل‌ها و فرمول‌ها
    "caption_missing": "۷. جدول‌ها، نمودارها و شکل‌ها",
    "caption_numbering": "۷. جدول‌ها، نمودارها و شکل‌ها",

    # ۸. ارجاع‌دهی و فهرست منابع
    "cite_missing_ref": "۸. ارجاع‌دهی و فهرست منابع",
    "ref_not_cited": "۸. ارجاع‌دهی و فهرست منابع",
    "numeric_cite": "۸. ارجاع‌دهی و فهرست منابع",
    "tums_et_al_forbidden": "۸. ارجاع‌دهی و فهرست منابع",
    "tums_ref_page_format": "۸. ارجاع‌دهی و فهرست منابع",

    # ۹. چک‌لیست دانشگاه علوم پزشکی تهران
    "tums_checklist_missing": "۹. چک‌لیست اختصاصی دانشگاه علوم پزشکی تهران",
}

RULE_TITLES = {
    "arabic_letters": "حرف عربی (ي / ك) به‌جای ی / ک فارسی",
    "arabic_digits": "ارقام عربی به‌جای ارقام فارسی",
    "latin_digits": "رقم لاتین در متن فارسی",
    "mi_prefix": "پیشوند «می / نمی» بدون نیم‌فاصله",
    "plural_space": "«ها / های» جدا از واژه (نیم‌فاصله لازم است)",
    "comparative_space": "«تر / ترین» جدا از واژه",
    "ye_space": "«ی» اضافه‌ی جدا از واژه",
    "space_before_punct": "فاصله قبل از علائم نگارشی (، ؛ : . ؟ !)",
    "no_space_after_punct": "نبودن فاصله بعد از علائم نگارشی",
    "latin_punct": "علائم لاتین (, ; ?) در متن فارسی",
    "paren_space": "فاصله‌گذاری نادرست پرانتز",
    "straight_quotes": "گیومه‌ی لاتین \" \" به‌جای « »",
    "double_space": "فاصله‌ی اضافی بین کلمات",
    "leading_space": "تورفتگی با Space/Tab به‌جای تنظیم پاراگراف",
    "repeated_word": "تکرار پشت‌سرهم یک واژه",
    "common_mistake": "غلط املایی/نگارشی رایج",
    "informal_word": "واژه‌ی محاوره‌ای در متن علمی",
    "first_person": "ضمیر اول‌شخص در نگارش علمی",
    "blank_lines": "پاراگراف خالی پشت‌سرهم (فاصله با Enter)",
    "long_paragraph": "پاراگراف بسیار طولانی",
    "long_sentence": "جمله‌ی بسیار طولانی",
    "page_size": "اندازه‌ی صفحه",
    "page_margin": "حاشیه‌های صفحه",
    "font_fa": "فونت متن فارسی",
    "size_fa": "اندازه‌ی قلم متن فارسی",
    "font_en": "فونت متن انگلیسی",
    "size_en": "اندازه‌ی قلم متن انگلیسی",
    "line_spacing": "فاصله‌ی خطوط",
    "alignment": "تراز پاراگراف (دوطرفه)",
    "rtl": "جهت پاراگراف (راست‌به‌چپ) تنظیم نشده",
    "missing_section": "بخش الزامی یافت نشد",
    "recommended_section": "بخش پیشنهادی یافت نشد",
    "chapters": "فصل‌ها",
    "heading_style": "عنوان فصل با استایل Heading نیست",
    "toc": "فهرست مطالب",
    "page_numbers": "شماره‌ی صفحه",
    "abstract_length": "طول چکیده",
    "keywords_count": "تعداد کلیدواژه‌ها",
    "caption_missing": "جدول/شکل بدون عنوان (کپشن)",
    "caption_numbering": "شماره‌گذاری جدول/شکل",
    "cite_missing_ref": "ارجاع درون‌متنی بدون مورد متناظر در فهرست منابع",
    "ref_not_cited": "منبع فهرست‌شده که در متن ارجاع نشده",
    "numeric_cite": "ارجاع عددی خارج از محدوده",
    # موارد افزوده شده جدید:
    "academic_tone_emotional": "لحن احساسی یا مبالغه‌آمیز در نگارش علمی",
    "academic_pleonasm": "حشو و زوائد بیانی (ساده‌نویسی علمی)",
    "academic_first_person": "کاربرد ضمیر اول‌شخص در نگارش دانشگاهی",
    "academic_cohesion": "انسجام و پیوستگی استدلال (نبود نشانگر گذار)",
    "en_punctuation": "نشانه‌گذاری نادرست در متن انگلیسی",
    "en_article": "کاربرد نادرست حروف تعریف (a / an) در انگلیسی",
    "en_academic_tone": "لحن و انتخاب واژگان آکادمیک در انگلیسی",
    "en_wordiness": "پرگویی و حشو در متن انگلیسی (Wordiness)",
    "en_preposition": "حرف اضافه نادرست در متن انگلیسی",
    "en_grammar_agreement": "تطابق دستوری فاعل/فعل یا شمارش اسامی انگلیسی",
    "en_unpaired_brackets": "عدم تقارن پرانتز یا گیومه در انگلیسی",
    "tums_checklist_missing": "بخش الزامی از چک‌لیست دانشگاه علوم پزشکی تهران",
    "tums_chapter_proportions": "توازن حجم فصول (سقف ۳۰٪ برای فصول ۱ و ۲)",
    "tums_et_al_forbidden": "ممنوعیت «همکاران» و «et al» در فهرست منابع TUMS",
    "tums_ref_page_format": "فرمت شماره صفحات در مراجع TUMS",
    "tums_page_margin": "انطباق حاشیه‌ها با ضوابط دانشگاه علوم پزشکی تهران",
}


@dataclass
class Issue:
    rule: str
    severity: str
    message: str
    loc: str = ""
    snippet: str = ""
    par_index: int = -1
    suggestion: str = ""
    target_text: str = ""


# ----------------------------------------------------------------------------
# ابزارهای کمکی
# ----------------------------------------------------------------------------
def load_rules(path=None, mode="thesis"):
    """بارگذاری قوانین پایه و ادغام با پروفایل اختصاصی در صورت انتخاب TUMS."""
    base_path = path or os.path.join(HERE, "rules.json")
    with open(base_path, encoding="utf-8") as f:
        rules = json.load(f)

    if mode == "tums":
        tums_path = os.path.join(HERE, "profiles", "tums.json")
        if os.path.exists(tums_path):
            with open(tums_path, encoding="utf-8") as tf:
                tums_data = json.load(tf)
                # ادغام مقادیر
                rules["page"] = tums_data.get("page", rules["page"])
                rules["body"] = tums_data.get("body", rules["body"])
                rules["limits"] = tums_data.get("limits", rules["limits"])
                rules["tums_special_rules"] = tums_data.get("tums_special_rules", {})
                rules["checklist_items"] = tums_data.get("checklist_items", [])
                if "tums" not in rules["modes"]:
                    rules["modes"]["tums"] = {
                        "title": "پایان‌نامه دانشگاه علوم پزشکی تهران (TUMS)",
                        "required_sections": rules["modes"]["thesis"]["required_sections"],
                        "recommended_sections": rules["modes"]["thesis"].get("recommended_sections", [])
                    }
    elif mode == "english":
        if "english" not in rules["modes"]:
            rules["modes"]["english"] = {
                "title": "متن / مقاله انگلیسی",
                "required_sections": [],
                "recommended_sections": []
            }
    return rules


def norm(s):
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("ي", "ی").replace("ك", "ک").replace(ZWNJ, " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def to_latin_digits(s):
    return s.translate(DIGIT_MAP)


def words(s):
    return re.findall(r"[\w\u200c]+", s)


def has_fa(s):
    if not s:
        return False
    s = unicodedata.normalize("NFKC", s)
    return bool(FA_RE.search(s))


def snippet_of(text, m, pad=18):
    a, b = max(0, m.start() - pad), min(len(text), m.end() + pad)
    out = text[a:b].replace("\n", " ")
    return ("…" if a else "") + out + ("…" if b < len(text) else "")


def is_heading(p):
    n = (p.style.name or "").lower() if p.style is not None else ""
    return n.startswith("heading") or n.startswith("عنوان") or n == "title"


def is_toc_style(p):
    n = (p.style.name or "").lower() if p.style is not None else ""
    return n.startswith("toc") or n.startswith("table of") or "فهرست" in n


def _attr(parent, tag, attr):
    if parent is None:
        return None
    el = parent.find(qn(tag))
    return None if el is None else el.get(qn(attr))


def _defaults(doc, kind):
    d = doc.styles.element.find(qn("w:docDefaults"))
    if d is None:
        return None
    holder = d.find(qn("w:rPrDefault" if kind == "r" else "w:pPrDefault"))
    if holder is None:
        return None
    return holder.find(qn("w:rPr" if kind == "r" else "w:pPr"))


def _levels(doc, par, kind, run=None):
    if kind == "r":
        if run is not None:
            yield run._r.rPr
        st = par.style
        while st is not None:
            yield st.element.rPr
            st = st.base_style
    else:
        yield par._p.pPr
        st = par.style
        while st is not None:
            yield st.element.pPr
            st = st.base_style
    yield _defaults(doc, kind)


def eff_rpr(doc, par, run, tag, attr):
    for lvl in _levels(doc, par, "r", run):
        v = _attr(lvl, tag, attr)
        if v is not None:
            return v
    return None


def eff_ppr(doc, par, tag, attr):
    for lvl in _levels(doc, par, "p"):
        v = _attr(lvl, tag, attr)
        if v is not None:
            return v
    return None


def eff_bidi(doc, par):
    for lvl in _levels(doc, par, "p"):
        if lvl is not None and lvl.find(qn("w:bidi")) is not None:
            el = lvl.find(qn("w:bidi"))
            return el.get(qn("w:val")) not in ("0", "false")
    return False


def font_ok(name, allowed):
    if name is None:
        return True
    n = name.lower().replace(" ", "")
    return any(n == a.lower().replace(" ", "") for a in allowed)


# ----------------------------------------------------------------------------
# قوانین متنی (تایپوگرافی، سبک علمی و انگلیسی)
# ----------------------------------------------------------------------------
MI_STEMS = ("کن|کرد|شو|شد|توان|باش|گوی|گفت|ده|داد|خواه|خواست|گیر|گرفت|بین|دید|رس|رو|"
            "نویس|پرداز|یاب|یافت|آید|آورد|ماند|ساز|شناس|پذیر")

SCIENTIFIC_UNITS_TAIL_RE = re.compile(
    r"^\s*(?:°\s*[cCfFkK]|(?:µ|u|m|c|k|d|n|p)?[gmldL]|(?:µ|u|m|k)?[vVaAwWjJ]|"
    r"sec|min|hr|hours?|days?|years?|hz|kHz|MHz|GHz|pa|kPa|MPa|GPa|mol|mmol|µmol|nmol|pmol|"
    r"rpm|bp|kb|mb|gb|tb|kda|kDa|iu|IU|cfu|CFU|ph|pH|%|‰|[xX]|fold)\b",
    re.IGNORECASE
)
SCIENTIFIC_HEAD_RE = re.compile(
    r"(?:(?:p|r|t|f|z)\s*[<>=~]|doi:|url:|https?://|www\.|n\s*=|ci\s*|ph\s*|covid-?)\s*$",
    re.IGNORECASE
)
ABBR_EXEMPT_RE = re.compile(r"(?:[هـقصدجم]\.[شقم]?)")

REDUPLICATION_WORDS = {
    "کم", "ذره", "پله", "تک", "رفته", "قطره", "دسته", "بخش", "صفحه",
    "سطر", "نکته", "لحظه", "ثانیه", "قدم", "سال", "روز", "ماه",
    "یواش", "آرام", "دور", "نرم", "تکه", "بار", "برگ", "دانه", "گروه",
    "مرحله", "سطح", "جزء", "تارتار", "یک", "دو"
}
REPEATED_WORD_RE = re.compile(r"(?<![" + FA + r"])([" + FA + r"]{2,})[ \u00a0]+\1(?![" + FA + "])")

TEXT_RULES = [
    ("arabic_letters", "error", re.compile("[يك]"),
     "حرف عربی به‌کار رفته؛ «ی» و «ک» فارسی تایپ شود (کیبورد فارسی استاندارد)."),
    ("arabic_digits", "error", re.compile("[٠-٩]"),
     "رقم عربی (٠١٢) به‌کار رفته؛ از ارقام فارسی (۰۱۲) استفاده کنید."),
    ("latin_digits", "warn",
     re.compile("[" + FA + r"]\s?[0-9]{1,4}(?![A-Za-z0-9])|(?<![A-Za-z0-9.,/-])[0-9]{1,4}\s?[" + FA + "]"),
     "رقم لاتین کنار متن فارسی؛ در متن فارسی ارقام فارسی بنویسید."),
    ("mi_prefix", "error",
     re.compile(r"(?<![" + FA + ZWNJ + r"])(ن?می)(?:" + MI_STEMS + r")"),
     "پیشوند «می/نمی» باید با نیم‌فاصله به فعل بچسبد: «می‌کند»."),
    ("mi_prefix", "error",
     re.compile(r"(?<![" + FA + ZWNJ + r"])ن?می[ \u00a0]+(?=[" + FA + r"]{2})"),
     "«می/نمی» با فاصله از فعل جدا شده؛ باید نیم‌فاصله باشد: «می‌کند»."),
    ("plural_space", "error",
     re.compile("[" + FA + r"]{2,}[ \u00a0]+(?:ها|های|هایی|هایش|هایم|هایمان|هایتان|هایشان)(?![" + FA + "])"),
     "«ها/های» با فاصله‌ی معمولی جدا شده؛ نیم‌فاصله بگذارید: «کتاب‌ها»."),
    ("comparative_space", "warn",
     re.compile("[" + FA + r"]{2,}[ \u00a0]+(?:تر|ترین)(?![" + FA + "])"),
     "«تر/ترین» باید با نیم‌فاصله به واژه بچسبد: «بزرگ‌تر»."),
    ("ye_space", "warn",
     re.compile("[" + FA + r"][ \u00a0]+ی(?![" + FA + "])"),
     "«ی» اضافه جدا از واژه است؛ «خانه‌ی» یا «خانهٔ»."),
    ("ye_space", "warn",
     re.compile("[" + FA + r"][ \u00a0]+ای(?![" + FA + "])"),
     "«ای» نکره جدا از واژه است؛ «نامه‌ای»."),
    ("space_before_punct", "warn",
     re.compile("(?<=[" + FA + r"\)])[ \u00a0]+[،؛؟!:\.](?!\.)"),
     "قبل از علائم نگارشی فاصله نیاید و بعد از آن‌ها یک فاصله بیاید."),
    ("no_space_after_punct", "warn",
     re.compile("[،؛؟](?=[" + FA + r"\w])|(?<=[" + FA + r"])\.(?=[" + FA + r"])|(?<=[" + FA + r"]):(?=[" + FA + "])"),
     "بعد از علامت نگارشی یک فاصله لازم است."),
    ("latin_punct", "warn",
     re.compile("(?<=[" + FA + r"])[,;?]"),
     "از علائم فارسی استفاده کنید: «،» «؛» «؟»."),
    ("paren_space", "warn",
     re.compile(r"\([ \u00a0]+|[ \u00a0]+\)|(?<=[" + FA + r"])\((?![A-Za-z0-9۰-۹])|\)(?=[" + FA + "])"),
     "داخل پرانتز فاصله نیاید؛ قبل از «(» و بعد از «)» یک فاصله بگذارید."),
    ("straight_quotes", "warn", re.compile(r'"[^"\n]{1,200}"'),
     "از گیومه‌ی فارسی « » استفاده کنید."),
    ("double_space", "warn", re.compile(r"(?<=\S)[ \u00a0]{2,}(?=\S)"),
     "بیش از یک فاصله بین کلمات."),
    ("leading_space", "warn", re.compile(r"^[ \t\u00a0]+(?=\S)"),
     "ابتدای پاراگراف با Space/Tab تورفتگی داده شده؛ از تنظیمات پاراگراف (First line) استفاده کنید."),
]


def _word_re(w):
    return re.compile("(?<![" + FA + ZWNJ + "])" + re.escape(w) + "(?![" + FA + ZWNJ + "])")


def check_text(par_index, text, loc, rules, in_refs=False, in_declaration=False):
    out = []
    if not text.strip():
        return out

    # تایپوگرافی عمومی
    for rule, sev, rx, msg in TEXT_RULES:
        if rule in ("arabic_digits", "arabic_letters") or has_fa(text):
            for m in rx.finditer(text):
                # فیلتر استثناهای علمی ارقام لاتین (واحدها، درصدها، شاخص‌های آماری)
                if rule == "latin_digits":
                    after_pos = m.end()
                    tail = text[after_pos:after_pos + 12]
                    head = text[max(0, m.start() - 10):m.start()]
                    if SCIENTIFIC_UNITS_TAIL_RE.match(tail) or SCIENTIFIC_HEAD_RE.search(head):
                        continue
                # فیلتر اختصارات علمی و تاریخی (مانند هـ.ش یا ص.)
                if rule == "no_space_after_punct":
                    snippet = text[max(0, m.start() - 3):min(len(text), m.end() + 3)]
                    if ABBR_EXEMPT_RE.search(snippet):
                        continue
                out.append(Issue(rule, sev, msg, loc, snippet_of(text, m), par_index, target_text=m.group(0)))

    # بررسی تکرار واژه با تفکیک واژگان تکرارشدنی مجاز (مانند کم‌کم یا ذره‌ذره)
    if has_fa(text):
        for m in REPEATED_WORD_RE.finditer(text):
            w = m.group(1)
            if w in REDUPLICATION_WORDS:
                out.append(Issue("repeated_word", "info",
                                 f"تکرار واژه «{w} {w}»؛ برای ساختار ترکیبی بهتر است با نیم‌فاصله پیوند داده شود: «{w}‌{w}».",
                                 loc, snippet_of(text, m), par_index, suggestion=f"{w}‌{w}", target_text=m.group(0)))
            else:
                out.append(Issue("repeated_word", "warn",
                                 f"تکرار ناخواسته واژه «{w}» پشت‌سرهم؛ واژه دوم اضافی است.",
                                 loc, snippet_of(text, m), par_index, suggestion=w, target_text=m.group(0)))

    # بررسی انگلیسی در صورت وجود متن انگلیسی در پاراگراف
    en_words = [w for w in text.split() if LATIN_RE.match(w)]
    if len(en_words) >= 3:
        en_issues = en_checker.check_english_text(text, loc)
        for ei in en_issues:
            out.append(Issue(ei.rule, ei.severity, ei.message, ei.loc, ei.snippet, par_index, ei.suggestion, target_text=ei.target_text))

    if not has_fa(text):
        return out

    t = text.replace("ي", "ی").replace("ك", "ک")
    for wrong, right in rules.get("common_mistakes", {}).items():
        rx = _word_re(wrong)
        for m in rx.finditer(t):
            out.append(Issue("common_mistake", "error", f"«{wrong}» ← «{right}»", loc,
                             snippet_of(t, m), par_index, right, target_text=m.group(0)))
    if in_refs:
        return out

    # واژگان محاوره‌ای
    for w in rules.get("informal_words", []):
        for m in _word_re(w).finditer(t):
            out.append(Issue("informal_word", "info",
                             f"«{w}» محاوره‌ای است؛ معادل رسمی بنویسید.", loc, snippet_of(t, m), par_index, target_text=m.group(0)))

    # بررسی سبک علمی و مبالغه (ماژول تخصصی academic_checker)
    acad_issues = academic_checker.check_academic_style(text, loc, in_declaration)
    for ai in acad_issues:
        out.append(Issue(ai.rule, ai.severity, ai.message, ai.loc, ai.snippet, par_index, ai.suggestion, target_text=ai.target_text))

    lim = rules.get("limits", {})
    n_words = len(words(text))
    if n_words > lim.get("max_paragraph_words", 300):
        out.append(Issue("long_paragraph", "info",
                         f"پاراگراف {n_words} کلمه دارد؛ بهتر است برای خوانایی شکسته شود.", loc, text[:60] + "…", par_index))
    for sent in re.split(r"[\.؟!؛\?]\s", text):
        if len(words(sent)) > lim.get("max_sentence_words", 60):
            out.append(Issue("long_sentence", "info",
                             f"جمله‌ای با {len(words(sent))} کلمه؛ برای خوانایی و درک ساده‌تر کوتاه‌تر شود.", loc,
                             sent[:60] + "…", par_index))
    return out


# ----------------------------------------------------------------------------
# بررسی صفحه‌آرایی و استایل‌ها
# ----------------------------------------------------------------------------
def check_page(doc, rules, mode="thesis"):
    out = []
    cfg = rules["page"]
    tol, mtol = cfg["size_tolerance_cm"], cfg["margin_tolerance_cm"]
    for i, s in enumerate(doc.sections, 1):
        if s.page_width is None or s.page_height is None:
            continue
        w, h = s.page_width / 360000, s.page_height / 360000
        if abs(w - cfg["width_cm"]) > tol or abs(h - cfg["height_cm"]) > tol:
            out.append(Issue("page_size", "error",
                             f"بخش {i}: اندازه‌ی صفحه {w:.1f}×{h:.1f} سانتی‌متر است؛ باید A4 "
                             f"({cfg['width_cm']}×{cfg['height_cm']}) باشد.", f"بخش {i}"))
        for side, attr in (("top", "top_margin"), ("bottom", "bottom_margin"),
                           ("right", "right_margin"), ("left", "left_margin")):
            v = getattr(s, attr)
            if v is None:
                continue
            cm, want = v / 360000, cfg["margins_cm"][side]
            if abs(cm - want) > mtol:
                nm = {"top": "بالا", "bottom": "پایین", "right": "راست", "left": "چپ"}[side]
                out.append(Issue("page_margin", "warn",
                                 f"بخش {i}: حاشیه‌ی {nm} {cm:.2f}cm است؛ باید {want}cm باشد.", f"بخش {i}"))

    # بررسی حاشیه‌های تخصصی TUMS
    if mode == "tums":
        tums_issues = tums_checker.check_tums_margins(doc.sections)
        for ti in tums_issues:
            out.append(Issue(ti.rule, ti.severity, ti.message, ti.loc, ti.snippet))

    return out


def check_body_format(doc, rules, ctx):
    out = []
    b = rules["body"]
    for i, p, heading in ctx:
        text = p.text
        if len(text.strip()) < b.get("min_chars_to_count_as_body", 40):
            continue
        if is_heading(p) or is_toc_style(p) or (p.style is not None and "caption" in (p.style.name or "").lower()):
            continue
        loc = f"زیر «{heading}»، پاراگراف {i + 1}"
        bad = defaultdict(set)
        for r in p.runs:
            if not r.text.strip():
                continue
            fa, en = has_fa(r.text), bool(LATIN_RE.search(r.text))
            if fa:
                name = eff_rpr(doc, p, r, "w:rFonts", "w:cs")
                sz = eff_rpr(doc, p, r, "w:szCs", "w:val")
                if not font_ok(name, b["fa_fonts"]):
                    bad["font_fa"].add(name)
                if sz and abs(int(sz) / 2 - b["fa_size_pt"]) > 0.1:
                    bad["size_fa"].add(f"{int(sz) / 2:g}")
            elif en:
                name = eff_rpr(doc, p, r, "w:rFonts", "w:ascii")
                sz = eff_rpr(doc, p, r, "w:sz", "w:val")
                if not font_ok(name, b["en_fonts"]):
                    bad["font_en"].add(name)
                if sz and abs(int(sz) / 2 - b["en_size_pt"]) > 0.1:
                    bad["size_en"].add(f"{int(sz) / 2:g}")
        snip = text[:50] + "…"
        for k, vals in bad.items():
            want = {"font_fa": b["fa_fonts"][0], "size_fa": b["fa_size_pt"],
                    "font_en": b["en_fonts"][0], "size_en": b["en_size_pt"]}[k]
            out.append(Issue(k, "error", f"یافت‌شده: {', '.join(map(str, vals))} — مطلوب: {want}", loc, snip, i))

        # فاصله‌ی خطوط
        line, lrule = None, "auto"
        for lvl in _levels(doc, p, "p"):
            sp = lvl.find(qn("w:spacing")) if lvl is not None else None
            if sp is not None and sp.get(qn("w:line")) is not None:
                line, lrule = int(sp.get(qn("w:line"))), sp.get(qn("w:lineRule")) or "auto"
                break
        if line is None:
            line, lrule = 240, "auto"
        if lrule == "auto":
            mult = line / 240
            target_spacing = b.get("line_spacing", 1.5)
            tol = b.get("line_spacing_tolerance", 0.2)
            if abs(mult - target_spacing) > tol:
                out.append(Issue("line_spacing", "warn",
                                 f"فاصله‌ی خطوط {mult:.2f} است؛ باید {target_spacing} باشد.", loc, snip, i))

        # تراز
        if b.get("alignment") == "justify" and has_fa(text):
            jc = eff_ppr(doc, p, "w:jc", "w:val")
            if jc not in ("both", "distribute", "thaiDistribute"):
                out.append(Issue("alignment", "warn", "متن دوطرفه (Justify) نیست.", loc, snip, i))

        # جهت
        if has_fa(text) and not eff_bidi(doc, p):
            out.append(Issue("rtl", "warn", "جهت پاراگراف راست‌به‌چپ (RTL) نیست.", loc, snip, i))
    return out


def check_fields(doc):
    out = []
    instr = " ".join(e.text or "" for e in doc.element.body.iter(qn("w:instrText")))
    has_toc = "TOC" in instr
    if not has_toc:
        out.append(Issue("toc", "info",
                         "فهرست مطالب خودکار (References ▸ Table of Contents) پیدا نشد؛ اگر فهرست را دستی "
                         "نوشته‌اید، شماره‌صفحه‌ها با ویرایش به‌هم می‌خورد."))
    has_page = False
    for s in doc.sections:
        for part in (s.footer, s.header):
            try:
                x = part._element
            except Exception:
                continue
            if any("PAGE" in (e.text or "") for e in x.iter(qn("w:instrText"))) or \
               any("PAGE" in (e.get(qn("w:instr")) or "") for e in x.iter(qn("w:fldSimple"))):
                has_page = True
    if not has_page:
        out.append(Issue("page_numbers", "warn", "شماره‌ی صفحه‌ی خودکار در سربرگ/پانویس پیدا نشد."))
    return out


# ----------------------------------------------------------------------------
# ساختار: بخش‌ها، فصل‌ها، چکیده، کپشن‌ها
# ----------------------------------------------------------------------------
ORD = {"اول": 1, "دوم": 2, "سوم": 3, "چهارم": 4, "پنجم": 5, "ششم": 6, "هفتم": 7, "هشتم": 8, "نهم": 9, "دهم": 10}
CHAPTER_RE = re.compile(r"^\s*فصل\s*(اول|دوم|سوم|چهارم|پنجم|ششم|هفتم|هشتم|نهم|دهم|[0-9۰-۹]+)")
CAPTION_RE = re.compile(r"^\s*(جدول|شکل|نمودار|تصویر)\s*(?:شماره\s*)?\(?([0-9۰-۹]+)(?:[\s\-–ـ\.]*([0-9۰-۹]+))?\)?")


def find_section(paragraphs, patterns):
    pats = [norm(x) for x in patterns]
    for i, p in enumerate(paragraphs):
        t = norm(p.text)
        if not t or (len(t.split()) > 10 and not is_heading(p)):
            continue
        if any(x in t for x in pats):
            return i
    return None


def check_structure(doc, rules, mode, stats):
    out = []
    ps = doc.paragraphs
    m = rules["modes"].get(mode, rules["modes"].get("thesis", {}))

    for sec in m.get("required_sections", []):
        if find_section(ps, sec["patterns"]) is None:
            out.append(Issue("missing_section", "error", f"بخش «{sec['name']}» پیدا نشد."))
    for sec in m.get("recommended_sections", []):
        if find_section(ps, sec["patterns"]) is None:
            out.append(Issue("recommended_section", "info", f"بخش «{sec['name']}» پیدا نشد."))

    # ارزیابی چک‌لیست ۲۷گانه دانشگاه علوم پزشکی تهران
    if mode == "tums":
        checklist_status, tums_issues = tums_checker.check_tums_checklist(ps)
        stats["tums_checklist"] = checklist_status
        for ti in tums_issues:
            out.append(Issue(ti.rule, ti.severity, ti.message, ti.loc, ti.snippet))

    # فصل‌ها و واژگان هر فصل
    nums, unstyled = [], []
    ch_indices = []
    for i, p in enumerate(ps):
        mm = CHAPTER_RE.match(norm(p.text).replace("ي", "ی"))
        if mm and len(p.text.split()) <= 12 and not is_toc_style(p) and "." * 4 not in p.text:
            v = mm.group(1)
            n = ORD.get(v) or int(to_latin_digits(v))
            nums.append(n)
            ch_indices.append((n, i))
            if not is_heading(p):
                unstyled.append(i)
    stats["chapters"] = len(nums)

    # شمارش کلمات هر فصل جهت بررسی سقف ۳۰ درصدی فصول ۱ و ۲ در TUMS
    ch_words = {}
    for idx, (ch_num, start_p) in enumerate(ch_indices):
        end_p = ch_indices[idx + 1][1] if idx + 1 < len(ch_indices) else len(ps)
        w_cnt = sum(len(words(p.text)) for p in ps[start_p:end_p])
        ch_words[ch_num] = w_cnt

    if mode == "tums":
        total_w = stats.get("words", 0)
        prop_issues = tums_checker.check_tums_chapter_proportions(ch_words, total_w)
        for pi in prop_issues:
            out.append(Issue(pi.rule, pi.severity, pi.message, pi.loc, pi.snippet))

    if mode in ("thesis", "tums"):
        min_ch = rules.get("limits", {}).get("min_chapters_thesis", 5 if mode == "tums" else 3)
        if len(nums) < min_ch:
            out.append(Issue("chapters", "warn", f"تعداد {len(nums)} عنوان «فصل …» پیدا شد؛ پایان‌نامه استاندارد باید حداقل {min_ch} فصل کامل داشته باشد."))
        uniq = sorted(set(nums))
        if uniq and uniq != list(range(1, uniq[-1] + 1)):
            out.append(Issue("chapters", "warn", f"ترتیب/شماره‌ی فصل‌ها ناقص یا تکراری است: {nums}"))
    if unstyled:
        out.append(Issue("heading_style", "warn",
                         f"{len(unstyled)} عنوان فصل با استایل Heading ساخته نشده؛ این امر مانع ایجاد فهرست خودکار می‌شود.",
                         f"پاراگراف {unstyled[0] + 1}", ps[unstyled[0]].text[:50], unstyled[0]))

    # چکیده و کلیدواژه
    a_i = find_section(ps, ["چکیده", "چکیده فارسی"])
    k_i = find_section(ps, ["کلیدواژه", "کلید واژه", "واژگان کلیدی", "واژه های کلیدی", "کلمات کلیدی"])
    if a_i is not None:
        end = k_i if (k_i is not None and k_i > a_i) else min(a_i + 12, len(ps))
        n = sum(len(words(x.text)) for x in ps[a_i + 1:end])
        lo, hi = rules.get("limits", {}).get("abstract_words", [150, 350])
        if n and not lo <= n <= hi:
            out.append(Issue("abstract_length", "warn",
                             f"چکیده‌ی فارسی حدود {n} کلمه است؛ محدوده‌ی مطلوب {lo}–{hi} کلمه.",
                             f"پاراگراف {a_i + 1}", "", a_i))
    if k_i is not None:
        body = ps[k_i].text
        body = body.split(":", 1)[1] if ":" in body else (ps[k_i + 1].text if k_i + 1 < len(ps) else "")
        cnt = len([x for x in re.split(r"[،,؛;]", body) if x.strip()])
        lo, hi = rules.get("limits", {}).get("keywords_count", [3, 7])
        if cnt and not lo <= cnt <= hi:
            out.append(Issue("keywords_count", "warn",
                             f"{cnt} کلیدواژه یافت شد؛ تعداد مطلوب {lo}–{hi}.", f"پاراگراف {k_i + 1}", body[:50], k_i))

    # کپشن‌ها
    kinds = {"جدول": [], "شکل": [], "نمودار": []}
    for i, p in enumerate(ps):
        mm = CAPTION_RE.match(p.text)
        if mm and len(p.text.split()) < 40:
            raw_kind = mm.group(1)
            k = "شکل" if raw_kind == "تصویر" else raw_kind
            a = int(to_latin_digits(mm.group(2)))
            b = int(to_latin_digits(mm.group(3))) if mm.group(3) else None
            kinds[k].append((i, a, b))
    n_tables, n_imgs = len(doc.tables), len(doc.inline_shapes)
    stats["tables"], stats["images"] = n_tables, n_imgs
    if n_tables > len(kinds["جدول"]):
        out.append(Issue("caption_missing", "warn",
                         f"{n_tables} جدول در فایل هست ولی فقط {len(kinds['جدول'])} عنوان «جدول …» پیدا شد."))
    if n_imgs > len(kinds["شکل"]) + len(kinds["نمودار"]):
        out.append(Issue("caption_missing", "warn",
                         f"{n_imgs} تصویر در فایل هست ولی فقط {len(kinds['شکل']) + len(kinds['نمودار'])} "
                         "عنوان «شکل/نمودار …» پیدا شد."))
    for kind, items in kinds.items():
        last = {}
        for i, a, b in items:
            key = a if b is not None else 0
            seq = b if b is not None else a
            exp = last.get(key, 0) + 1
            if seq != exp:
                out.append(Issue("caption_numbering", "warn",
                                 f"{kind} {a}{'-' + str(b) if b else ''}: شماره‌ی بعدی باید {exp} می‌بود.",
                                 f"پاراگراف {i + 1}", ps[i].text[:50], i))
            last[key] = seq
    return out


# ----------------------------------------------------------------------------
# ارجاع‌دهی و منابع
# ----------------------------------------------------------------------------
REF_HEAD = ["منابع", "مراجع", "فهرست منابع", "references", "bibliography"]
AFTER_REFS_HEAD = [
    "پیوست", "پیوست‌ها", "ضمیمه", "ضمائم", "ضمایم", "appendix", "appendices",
    "چکیده انگلیسی", "abstract", "اظهارنامه", "تأییدیه", "صورتجلسه",
    "declaration", "فرم تعهد", "سپاسگزاری"
]
CITE_RE = re.compile(r"\(([^()\n]{2,90}?)[،,]\s*([0-9۰-۹]{4})[a-z]?(?:\s*[،,;؛][^()\n]{0,40})?\)")
NUM_CITE_RE = re.compile(r"\[([0-9۰-۹]+(?:\s*[-–,،]\s*[0-9۰-۹]+)*)\]")


def first_surname(s):
    s = norm(s)
    s = re.split(r"\s+(?:و|and|&|et al)\b|[،,;؛]", s)[0]
    return s.strip().split(" ")[0] if s.strip() else ""


def check_citations(doc, rules, mode="thesis"):
    out = []
    ps = doc.paragraphs
    ref_i = None
    for i in range(len(ps) - 1, -1, -1):
        t = norm(ps[i].text)
        if t and len(t.split()) <= 4 and any(t == norm(h) or t.startswith(norm(h)) for h in REF_HEAD):
            ref_i = i
            break
    if ref_i is None:
        return out, None

    refs = []
    for p in ps[ref_i + 1:]:
        t_clean = norm(p.text)
        if not t_clean:
            continue
        # اگر به سرفصل پیوست‌ها یا بخش‌های پس از منابع رسیدیم، شمارش مراجع را متوقف کن
        if len(p.text.split()) <= 5 and any(t_clean == norm(h) or t_clean.startswith(norm(h)) for h in AFTER_REFS_HEAD):
            break
        if len(p.text.strip()) > 15:
            refs.append(p.text.strip())

    body_text = "\n".join(p.text for p in ps[:ref_i])
    refs_norm = [to_latin_digits(norm(r)) for r in refs]

    # ارزیابی ضوابط مراجع دانشگاه علوم پزشکی تهران
    if mode == "tums":
        tums_ref_issues = tums_checker.check_tums_references(refs)
        for tri in tums_ref_issues:
            out.append(Issue(tri.rule, tri.severity, tri.message, tri.loc, tri.snippet))

    seen = set()
    cited = set()
    for m in CITE_RE.finditer(body_text):
        sur = first_surname(m.group(1))
        year = to_latin_digits(m.group(2))
        if not sur or (sur, year) in seen:
            continue
        seen.add((sur, year))
        hit = [k for k, r in enumerate(refs_norm) if sur in r and year in r]
        if hit:
            cited.update(hit)
        else:
            out.append(Issue("cite_missing_ref", "error",
                             f"ارجاع «{m.group(0)}» در فهرست منابع پیدا نشد.", "", m.group(0)))
    if seen:
        for k, r in enumerate(refs):
            if k not in cited:
                out.append(Issue("ref_not_cited", "warn", "این منبع در متن ارجاع داده نشده است.", "فهرست منابع", r[:70]))
    # ارجاع عددی
    nums = set()
    for m in NUM_CITE_RE.finditer(body_text):
        for tok in re.split(r"[,،]", to_latin_digits(m.group(1))):
            parts = re.split(r"[-–]", tok)
            for x in parts:
                x_str = x.strip()
                if x_str.isdigit():
                    val = int(x_str)
                    if val > 0:  # مراجع از ۱ آغاز می‌شوند؛ فیلتر اعداد غیرمرتبط یا بازه‌های ریاضی
                        nums.add(val)
    if nums and refs and max(nums) > len(refs):
        out.append(Issue("numeric_cite", "error",
                         f"ارجاع [{max(nums)}] در متن هست ولی فهرست منابع فقط {len(refs)} مورد دارد."))
    return out, ref_i


# ----------------------------------------------------------------------------
# اجرای کامل روی Word (DOCX)
# ----------------------------------------------------------------------------
def analyze(path, mode="thesis", rules=None):
    rules = rules or load_rules(mode=mode)
    doc = Document(path)
    issues, stats = [], {"mode": mode}

    heading = "ابتدای سند"
    ctx = []
    paragraphs_texts = []
    for i, p in enumerate(doc.paragraphs):
        t = p.text
        if is_heading(p) and t.strip():
            heading = t.strip()[:50]
        ctx.append((i, p, heading))
        paragraphs_texts.append(t)

    stats["paragraphs"] = len(ctx)
    stats["words"] = sum(len(words(p.text)) for p in doc.paragraphs)

    cit_issues, ref_i = check_citations(doc, rules, mode=mode)

    blank_run = 0
    for i, p, h in ctx:
        text = p.text
        if is_toc_style(p):
            continue
        if not text.strip() and not p._p.xpath(".//w:drawing") and not p._p.xpath(".//w:br[@w:type='page']"):
            blank_run += 1
            if blank_run == 3:
                issues.append(Issue("blank_lines", "warn",
                                    "چند پاراگراف خالی پشت‌سرهم؛ برای فاصله از Spacing پاراگراف یا Page Break استفاده کنید.",
                                    f"زیر «{h}»، پاراگراف {i + 1}", "", i))
            continue
        blank_run = 0
        in_refs = ref_i is not None and i > ref_i
        in_decl = any(kw in h.lower() for kw in ["تعهد", "اصالت", "سپاس", "تقدیم", "declaration"])
        issues += check_text(i, text, f"زیر «{h}»، پاراگراف {i + 1}", rules, in_refs, in_declaration=in_decl)

    # ارزیابی انسجام پاراگراف‌ها
    cohesion_issues = academic_checker.evaluate_cohesion(paragraphs_texts)
    for ci in cohesion_issues:
        issues.append(Issue(ci.rule, ci.severity, ci.message, ci.loc, ci.snippet))

    issues += check_page(doc, rules, mode=mode)
    issues += check_body_format(doc, rules, ctx)
    issues += check_fields(doc)
    issues += check_structure(doc, rules, mode, stats)
    issues += cit_issues
    return doc, issues, stats


# ----------------------------------------------------------------------------
# اجرای بررسی روی فایل PDF
# ----------------------------------------------------------------------------
def analyze_pdf(pdf_path, mode="thesis", rules=None):
    """تحلیل و استخراج متنی فایل PDF متناسب با اصول بند ۲ و ۱۰."""
    rules = rules or load_rules(mode=mode)
    extracted = pdf_handler.extract_pdf_content(pdf_path)
    issues, stats = [], {"mode": mode, "is_pdf": True}

    if not extracted.get("ok"):
        issues.append(Issue("pdf_read_error", "error", extracted.get("error", "خطا در خواندن فایل PDF")))
        return None, issues, stats

    if extracted.get("is_scanned"):
        issues.append(Issue("pdf_scanned_warning", "warn",
                            "این فایل PDF اسکن‌شده یا تصویرمحور است و فاقد لایه متنی استاندارد می‌باشد. طبق اصل ۱۰، استخراج متن با اطمینان کافی نیازمند سامانه OCR است و نمی‌توان درباره بخش‌های ناخوانا نظر قطعی داد."))
        stats["is_scanned"] = True
        return None, issues, stats

    all_paragraphs = []
    total_words = 0
    par_idx = 0

    for p_info in extracted.get("pages", []):
        p_num = p_info["page_num"]
        for par_text in p_info["paragraphs"]:
            all_paragraphs.append(par_text)
            w_cnt = len(words(par_text))
            total_words += w_cnt
            loc = f"صفحه {p_num}، پاراگراف {par_idx + 1}"
            issues += check_text(par_idx, par_text, loc, rules)
            par_idx += 1

    stats["paragraphs"] = len(all_paragraphs)
    stats["words"] = total_words
    stats["total_pages"] = extracted.get("total_pages", 0)

    # ارزیابی انسجام
    cohesion_issues = academic_checker.evaluate_cohesion(all_paragraphs)
    for ci in cohesion_issues:
        issues.append(Issue(ci.rule, ci.severity, ci.message, ci.loc, ci.snippet))

    return None, issues, stats


# ----------------------------------------------------------------------------
# اجرای بررسی روی متن مستقیم ارسالی کاربر در چت
# ----------------------------------------------------------------------------
def analyze_text(raw_text, mode="thesis", rules=None):
    """بررسی متن مستقیم وارد شده در چت بله."""
    rules = rules or load_rules(mode=mode)
    issues, stats = [], {"mode": mode, "is_raw_text": True}
    paragraphs = [p.strip() for p in raw_text.split("\n\n") if p.strip()]

    stats["paragraphs"] = len(paragraphs)
    stats["words"] = len(words(raw_text))

    for i, p in enumerate(paragraphs):
        loc = f"پاراگراف {i + 1}"
        issues += check_text(i, p, loc, rules)

    cohesion_issues = academic_checker.evaluate_cohesion(paragraphs)
    for ci in cohesion_issues:
        issues.append(Issue(ci.rule, ci.severity, ci.message, ci.loc, ci.snippet))

    return None, issues, stats


# ----------------------------------------------------------------------------
# تولید گزارش کامل و ساختاریافته (منطبق بر بخش ۱۶ دستورالعمل)
# ----------------------------------------------------------------------------
def build_report(issues, stats, rules, examples=1000):
    mode_key = stats.get("mode", "thesis")
    mode_title = rules.get("modes", {}).get(mode_key, {}).get("title", mode_key)
    if mode_key == "tums":
        mode_title = "پایان‌نامه دانشگاه علوم پزشکی تهران (TUMS)"

    sev_count = Counter(it.severity for it in issues)

    lines = [
        "══════════════════════════════════════════════════════════════",
        "  سامانه جامع ویراستار پژوهش — گزارش ارزیابی علمی و نگارشی",
        "══════════════════════════════════════════════════════════════",
        "",
        "الف) خلاصه ارزیابی و مشخصات سند",
        "──────────────────────────────────────────────────────────────",
        f"• شیوه نامه / الگوی انتخابی: {mode_title}",
        f"• تعداد پاراگراف‌ها: {stats.get('paragraphs', 0)} | تعداد کلمات: {stats.get('words', 0)}",
    ]
    if "total_pages" in stats:
        lines.append(f"• تعداد صفحات PDF: {stats['total_pages']}")
    if "tables" in stats or "images" in stats:
        lines.append(f"• تعداد جدول‌ها: {stats.get('tables', 0)} | تعداد تصاویر: {stats.get('images', 0)}")
    lines += [
        f"• شاخص خطاها: {SEVERITY_ICON['error']} {sev_count['error']} خطای قطعی  |  "
        f"{SEVERITY_ICON['warn']} {sev_count['warn']} هشدار/اصلاحی  |  "
        f"{SEVERITY_ICON['info']} {sev_count['info']} پیشنهاد بهبود",
        "",
        "ب) دسته‌بندی تفکیک‌شده ایرادها",
        "──────────────────────────────────────────────────────────────",
    ]

    # گروه‌بندی بر اساس دسته‌بندی‌های استاندارد
    cat_groups = defaultdict(list)
    for it in issues:
        cat = RULE_CATEGORIES.get(it.rule, "سایر موارد")
        cat_groups[cat].append(it)

    for cat_name, cat_issues in sorted(cat_groups.items()):
        lines.append(f"📌 {cat_name}: {len(cat_issues)} مورد")

    # چک‌لیست دانشگاه علوم پزشکی تهران در صورت انتخاب TUMS
    if "tums_checklist" in stats and stats["tums_checklist"]:
        lines += [
            "",
            "ج) چک‌لیست انطباق با دستورالعمل دانشگاه علوم پزشکی تهران (TUMS)",
            "──────────────────────────────────────────────────────────────",
        ]
        for item in stats["tums_checklist"]:
            lines.append(f"  • {item['name']}: {item['status']}")

    lines += [
        "",
        "د) جزئیات ایرادها و پیشنهادهای اصلاحی",
        "──────────────────────────────────────────────────────────────",
    ]

    if not issues:
        lines.append("✅ هیچ خطای نگارشی، ساختاری یا قالب‌بندی قابل‌توجهی یافت نشد.")
    else:
        issue_counter = 1
        for cat_name, cat_issues in sorted(cat_groups.items()):
            lines.append(f"\n【 {cat_name} 】")
            for it in cat_issues[:examples]:
                rule_name = RULE_TITLES.get(it.rule, it.rule)
                row = f"[{issue_counter}] {SEVERITY_ICON[it.severity]} {rule_name}"
                if it.loc:
                    row += f"\n    📍 محل: {it.loc}"
                row += f"\n    📝 شرح: {it.message}"
                if it.snippet:
                    row += f"\n    🔍 متن: «{it.snippet}»"
                if it.suggestion:
                    row += f"\n    💡 پیشنهاد اصلاح: «{it.suggestion}»"
                lines.append(row + "\n")
                issue_counter += 1

    lines += [
        "",
        "هـ) جمع‌بندی و یادآوری‌های اخلاقی و پژوهشی",
        "──────────────────────────────────────────────────────────────",
        "۱. مطابق اصول ویرایش علمی، معنای داده‌ها، فرمول‌ها و نام‌ها بدون بازبینی تغییر نکرده است.",
        "۲. میان خطای قطعی و پیشنهادهای سبکی تفکیک شده است؛ پیشنهادهای بهبود نیازمند تأیید نهایی نویسنده هستند.",
        "۳. شیوه نامه رسمی دانشگاه یا نشریه مرجع نهایی ارزیابی است.",
        "══════════════════════════════════════════════════════════════"
    ]
    return "\n".join(lines)


def build_summary(issues, stats, rules):
    """خلاصه‌ی پیام چت در پیام‌رسان بله."""
    mode_key = stats.get("mode", "thesis")
    mode_title = rules.get("modes", {}).get(mode_key, {}).get("title", mode_key)
    if mode_key == "tums":
        mode_title = "دانشگاه علوم پزشکی تهران (TUMS)"

    sev = Counter(it.severity for it in issues)
    lines = [
        f"📄 بررسی «{mode_title}» به پایان رسید",
        f"📊 {stats.get('paragraphs', 0)} پاراگراف · {stats.get('words', 0)} کلمه",
        f"{SEVERITY_ICON['error']} {sev['error']} خطای قطعی  ·  {SEVERITY_ICON['warn']} {sev['warn']} هشدار  ·  {SEVERITY_ICON['info']} {sev['info']} پیشنهاد بهبود",
        "",
    ]
    # دسته‌بندی ایرادها
    cat_counts = Counter(RULE_CATEGORIES.get(it.rule, "سایر موارد") for it in issues)
    for cat, cnt in cat_counts.most_common(5):
        lines.append(f"• {cat}: {cnt} مورد")

    if not issues:
        lines.append("✅ هیچ موردی یافت نشد؛ متن منطبق بر اصول ارزیابی‌شده است.")
    else:
        lines.append("\n📋 گزارش کامل همراه با جزئیات دقیق ارسال شد.")
    return "\n".join(lines)


def split_message(text, limit=3800):
    parts, cur = [], ""
    for block in text.split("\n\n"):
        if len(cur) + len(block) + 2 > limit and cur:
            parts.append(cur.rstrip())
            cur = ""
        while len(block) > limit:
            parts.append(block[:limit])
            block = block[limit:]
        cur += block + "\n\n"
    if cur.strip():
        parts.append(cur.rstrip())
    return parts


def annotate(doc, issues, out_path, max_comments=250, per_par=4):
    """
    ذخیره‌ی نسخه‌ی کامنت‌گذاری‌شده‌ی فایل کاربر در Word.
    قاعده اساسی: به هیچ عنوان به متن اصلی دست زده نمی‌شود (حفظ ۱۰۰٪ نویسه‌ها و فاصله‌ها).
    کلمات یا عبارات دارای ایراد با رنگ زرد (Yellow Highlight) مشخص شده و کامنت تفصیلی به آن‌ها متصل می‌شود.
    """
    if doc is None or not hasattr(doc, "add_comment"):
        return False
    by_par = defaultdict(list)
    for it in sorted(issues, key=lambda x: SEVERITY_ORDER[x.severity]):
        if it.par_index >= 0:
            by_par[it.par_index].append(it)
    done = 0
    for idx in sorted(by_par):
        if done >= max_comments:
            break
        if idx >= len(doc.paragraphs):
            continue
        p = doc.paragraphs[idx]
        if not p.text.strip():
            continue

        seen = set()
        for it in by_par[idx]:
            if done >= max_comments:
                break
            key = (it.rule, it.message, it.target_text)
            if key in seen:
                continue
            seen.add(key)

            lines = [f"{SEVERITY_ICON[it.severity]} {RULE_TITLES.get(it.rule, it.rule)}: {it.message}"]
            if it.suggestion:
                lines.append(f"💡 پیشنهاد اصلاح: {it.suggestion}")
            comment_text = "\n".join(lines)

            target = it.target_text.strip() if it.target_text else ""
            target_run = None

            # اگر کلمه یا عبارت مشخصی هدف باشد و در متن وجود داشته باشد
            if target and target in p.text:
                for r in list(p.runs):
                    if target in r.text:
                        t = r.text
                        s_idx = t.find(target)
                        e_idx = s_idx + len(target)
                        before = t[:s_idx]
                        after = t[e_idx:]

                        # تبدیل r به بخش قبل، و اضافه کردن target هایلایت‌شده و after بدون تغییر هیچ حرفی
                        r.text = before

                        t_run = p.add_run(target)
                        if r._r.rPr is not None:
                            t_run._r.append(copy.deepcopy(r._r.rPr))
                        t_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

                        a_run = p.add_run(after)
                        if r._r.rPr is not None:
                            a_run._r.append(copy.deepcopy(r._r.rPr))

                        p_elm = p._p
                        r_elm = r._r
                        r_pos = p_elm.index(r_elm)
                        p_elm.insert(r_pos + 1, t_run._r)
                        p_elm.insert(r_pos + 2, a_run._r)

                        target_run = t_run
                        break

            if target_run is None:
                non_empty = [r for r in p.runs if r.text.strip()]
                if non_empty:
                    target_run = non_empty[0]

            if target_run is not None:
                try:
                    doc.add_comment(runs=target_run, text=comment_text, author="ویراستار پژوهش", initials="وپ")
                    done += 1
                except Exception:
                    continue

    doc.save(out_path)
    return done > 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--mode", choices=["thesis", "tums", "article", "english"], default="thesis")
    ap.add_argument("--annotate", help="مسیر خروجی docx کامنت‌گذاری‌شده")
    a = ap.parse_args()
    R = load_rules(mode=a.mode)
    if a.file.lower().endswith(".pdf"):
        d, iss, st = analyze_pdf(a.file, a.mode, R)
    else:
        d, iss, st = analyze(a.file, a.mode, R)
    print(build_report(iss, st, R))
    if a.annotate and d is not None:
        print("annotated:", annotate(d, iss, a.annotate))
