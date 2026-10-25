"""Read-only, source-labelled University of Kashan review.

Authority: the graduate office's three-page writing guide (source A) and the
official master's (B) and PhD (C) templates of Bahman 1398, supplied by the
bot owner. Word-format clauses live in kashan_format. All document findings
require human review (at most B). Textual presence is not proof of page
layout, signatures, official approval, originality or readiness.
"""
import copy
import json
from pathlib import Path
import re
import unicodedata

from analysis_model import Issue
from protected_content import ProtectedContentDetector
from rule_registry import FindingBuffer, bounded_findings

MODES = {"kashan": "unknown", "kashan_master": "master", "kashan_phd": "phd"}
_DEGREES = {"unknown": "مقطع تعیین نشده", "master": "کارشناسی ارشد", "phd": "دکتری"}
_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
_TRANSLATE = str.maketrans({"ي": "ی", "ك": "ک", "\u200c": " "})
SOURCE_LABEL = ("راهنمای سه‌صفحه‌ای و الگوهای رسمی ارشد و دکتری تحصیلات تکمیلی دانشگاه کاشان (بهمن ۱۳۹۸)، "
                "ارسالی مالک ربات؛ به‌روزبودن آن‌ها مستقلاً تأیید نشده و باید با تحصیلات تکمیلی تطبیق داده شود")
GUIDE_A = "راهنمای سه‌صفحه‌ای تحصیلات تکمیلی دانشگاه کاشان، {}"


# Ezafe and vowel marks or tatweel do not change which heading is meant:
# «تعهدنامهٔ»، «تعهدنامۀ» and «تعهدنامه‌ی اصالت» all name «تعهدنامه اصالت».
_MARKS = re.compile("[ـً-ٰٕ]")
_EZAFE_YE = re.compile(r"(?<=ه) ی(?= |$)")


def _norm(text):
    text = unicodedata.normalize("NFC", str(text)).replace("ۀ", "ه").translate(_TRANSLATE)
    text = _EZAFE_YE.sub("", _MARKS.sub("", text))
    return " ".join(text.split()).strip(" :：؛،.").casefold()


def _source(degree, section):
    """Cite the official template; ``section`` is kept for call compatibility."""
    if degree == "master":
        return "الگوی رسمی راهنمای تهیه و تدوین پایان‌نامه کارشناسی ارشد دانشگاه کاشان، بهمن ۱۳۹۸ (منبع B)"
    if degree == "phd":
        return "الگوی رسمی راهنمای تهیه و تدوین رساله دکتری دانشگاه کاشان، بهمن ۱۳۹۸ (منبع C)"
    return "الگوهای رسمی ارشد و دکتری دانشگاه کاشان، بهمن ۱۳۹۸ (منابع B و C)"


def _profile_metadata():
    path = Path(__file__).parent / "profiles" / "kashan.json"
    if path.is_file():
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    # A deterministic fallback keeps direct use safe; no university rules guessed.
    return {"id": "kashan", "title": "دانشگاه کاشان", "sources": []}


def apply_profile(rules, mode):
    """Return an isolated deep copy, preserving the caller's generic/TUMS rules.

    Numeric base values remain solely for backwards-compatible rule validation;
    profile_checks explicitly disables their execution. Generic structural rules
    are replaced by this module's source-aware checklist, not merged with it.
    """
    if mode not in MODES:
        raise ValueError("Unsupported Kashan mode")
    adjusted = copy.deepcopy(rules)
    adjusted["university_profile_id"] = "kashan"
    adjusted["university_degree"] = MODES[mode]
    adjusted["profile_checks"] = {
        "page": False, "body_format": False, "generic_structure": False,
        "generic_fields": False,
    }
    adjusted["kashan_profile"] = _profile_metadata()
    adjusted["reference_style"] = {"name": "APA", "edition": None,
                                   "source": "user_provided_summary"}
    # No degree or mandatory master sequence may leak into generic Kashan mode.
    for name, degree in MODES.items():
        adjusted.setdefault("modes", {})[name] = {
            "title": "دانشگاه کاشان — " + _DEGREES[degree],
            "required_sections": [], "recommended_sections": [],
        }
    for key in ("abstract_words", "keywords_count", "min_chapters_thesis"):
        adjusted.setdefault("limits", {}).pop(key, None)
    # This profile must not accidentally retain another university's checks.
    adjusted.pop("tums_special_rules", None)
    adjusted.pop("checklist_items", None)
    return adjusted


def _records(doc, blocks):
    """Keep original source strings and coordinates; never normalize the source."""
    if blocks is None:
        blocks = list(doc.paragraphs) if doc is not None else []
    result = []
    for index, block in enumerate(blocks):
        is_dict = isinstance(block, dict)
        getter = block.get if is_dict else lambda key, default=None: getattr(block, key, default)
        text = block if isinstance(block, str) else getter("text", "")
        if not isinstance(text, str) or not text.strip():
            continue
        paragraph = getter("paragraph", None)
        if paragraph is None and hasattr(block, "_p"):
            paragraph = block
        par_index = getter("paragraph_index", index if paragraph is not None else -1)
        story = getter("story", "body" if paragraph is not None else "text")
        style = getattr(getattr(paragraph, "style", None), "name", "") or ""
        result.append({"text": text, "paragraph": paragraph, "style": style,
                       "par_index": par_index,
                       "location": getter("location", f"پاراگراف {index + 1}"),
                       "block_id": getter("block_id", ""),
                       "source_part": getter("source_part", ""),
                       "xml_path": getter("xml_path", ""), "story": story,
                       "page_index": getter("page_index", -1),
                       "extraction_status": getter("extraction_status", "native"),
                       "index": len(result)})
    return result


def _is_toc(record):
    text = record["text"]
    style = record["style"].lower()
    toc_entry_style = style.startswith(("toc", "table of contents")) and style not in {"toc heading", "table of contents heading"}
    return (toc_entry_style
            or bool(re.search(r"[.\u2026]{3,}|\t\s*[0-9۰-۹]+\s*$", text)))


def _heading(record):
    """An isolated title, never a mention inside prose or a TOC entry."""
    if record["story"] not in {"body", "text", "pdf"} or _is_toc(record):
        return False
    if record["extraction_status"].startswith("partial:"):
        return False
    text = record["text"].strip()
    return "\n" not in text and len(text.split()) <= 12 and len(text) <= 110


# Master sequence from attachment section 5.1. Optional/conditional items never
# produce a missing-section finding. PhD intentionally does not inherit it.
_MASTER = [
    ("basmala", "صفحهٔ بسم‌الله", ("بسم الله الرحمن الرحیم", "بسم الله"), "required"),
    ("persian_cover", "صفحهٔ عنوان فارسی", (), "required"),
    ("defense_minutes", "صورت‌جلسهٔ دفاع", ("صورت جلسه دفاع*", "صورتجلسه دفاع*", "صورت جلسه", "صورتجلسه"), "required"),
    ("declaration", "تعهدنامه", ("تعهدنامه", "تعهد نامه", "تعهد اصالت اثر", "تعهدنامه اصالت اثر"), "required"),
    ("dedication", "تقدیم", ("تقدیم", "تقدیم به"), "optional"),
    ("acknowledgements", "سپاسگزاری", ("سپاسگزاری", "سپاس گزاری", "تقدیر و تشکر", "تشکر و قدردانی", "سپاس و قدردانی"), "required"),
    ("abstract_fa", "چکیدهٔ فارسی", ("چکیده", "چکیده فارسی"), "required"),
    ("toc", "فهرست مطالب", ("فهرست مطالب",), "required"),
    ("tables_list", "فهرست جدول‌ها", ("فهرست جدول ها", "فهرست جدولها", "فهرست جداول"), "conditional"),
    ("figures_list", "فهرست شکل‌ها", ("فهرست شکل ها", "فهرست شکلها", "فهرست اشکال"), "conditional"),
    ("abbreviations", "فهرست نشانه‌های اختصاری", ("فهرست نشانه های اختصاری", "فهرست اختصارات", "فهرست علائم و اختصارات"), "conditional"),
    ("chapters", "فصل‌های پژوهش؛ بدون الزام تعداد ثابت", (), "required"),
    ("references", "منابع و مراجع", ("منابع", "مراجع", "منابع و مراجع", "منابع و مآخذ", "references", "bibliography"), "required"),
    ("appendices", "پیوست‌ها، در صورت وجود", ("پیوست", "پیوست ها", "پیوستها", "ضمائم", "appendix", "appendices"), "optional"),
    ("abstract_en", "چکیدهٔ انگلیسی", ("abstract", "english abstract"), "required"),
    ("english_cover", "صفحهٔ عنوان انگلیسی", (), "required"),
]
_CHAPTER = re.compile(r"^فصل\s+(?:اول|دوم|سوم|چهارم|پنجم|ششم|هفتم|هشتم|نهم|دهم|[0-9۰-۹]+)(?:\s*[:：؛،.\-–]\s*.*)?$", re.I)


def _english_cover(records):
    """Require an isolated university marker plus nearby title-page context.

    A dissertation title in the bibliography or an English quotation is never
    treated as the user's English cover. Paragraph/page layout is unconfirmed.
    """
    for record in records:
        if not _heading(record) or not re.fullmatch(r"(?:university\s+of\s+kashan|kashan\s+university)", record["text"].strip(), re.I):
            continue
        window = records[max(0, record["index"] - 7):record["index"] + 12]
        if any(_heading(r) and re.fullmatch(r"references|bibliography", r["text"].strip(), re.I) for r in window):
            continue
        context = [r for r in window if r["story"] == record["story"] and not _is_toc(r)]
        if any(re.search(r"\b(?:submitted\s+(?:in|for|to)|supervisors?|faculty\s+of|degree\s+of|a\s+thesis|ph\.?\s*d\.?)\b", r["text"], re.I)
               for r in context):
            return record, context
    # Official template: «Faculty of …», then degree, supervisor and «By:».
    for record in records:
        if not (_heading(record) and re.match(r"\s*faculty\s+of\b", record["text"], re.I)):
            continue
        window = []
        for item in records[record["index"]:record["index"] + 18]:
            if window and _heading(item) and re.fullmatch(r"\s*(?:abstract|references|bibliography)\s*:?\s*", item["text"], re.I):
                break
            if item["story"] == record["story"] and not _is_toc(item):
                window.append(item)
        joined = " ".join(item["text"] for item in window)
        if (re.search(r"\b(?:thesis|dissertation)\b", joined, re.I)
                and re.search(r"\b(?:supervisors?|by)\s*(?:\(s\))?\s*:", joined, re.I)):
            return record, window
    return None, []


def _persian_cover(records):
    for record in records:
        if _heading(record) and _norm(record["text"]) == "دانشگاه کاشان":
            nearby = records[max(0, record["index"] - 3):record["index"] + 10]
            if any(re.search(r"(?:پایان[\s\u200c-]*نامه|رساله|استاد راهنما|دانشکده)", r["text"]) for r in nearby if not _is_toc(r)):
                return record
    # The official template prints «دانشگاه کاشان» inside the logo image, so
    # the page is also recognised by its own lines: faculty, degree,
    # supervisor and «توسط».
    for record in records[:60]:
        if not (_heading(record) and _norm(record["text"]).startswith("دانشکده")):
            continue
        nearby = " ".join(_norm(r["text"]) for r in records[record["index"]:record["index"] + 16] if not _is_toc(r))
        if (re.search(r"پایان ?نامه|رساله", nearby) and re.search(r"استاد(?:\(ان\)|ان)? ?راهنما", nearby)
                and re.search(r"توسط|به ?وسیله|نگارش", nearby)):
            return record
    return None


_STOP_HEADINGS = ("صورتجلسه دفاع*", "صورت جلسه دفاع*", "تعهدنامه", "تعهد نامه", "تقدیم", "سپاسگزاری",
                  "سپاس گزاری", "چکیده", "فهرست مطالب", "abstract", "فصل اول")


def persian_cover_window(records):
    """Records of the Persian title page, ending at the next page heading."""
    start = _persian_cover(records)
    if start is None:
        return []
    window = []
    for record in records[start["index"]:start["index"] + 20]:
        if window and _find_heading([record], _STOP_HEADINGS):
            break
        window.append(record)
    return window


def _find_heading(records, patterns):
    """Patterns ending in «*» match a heading prefix: the official minutes page
    is titled «صورتجلسه دفاع از پایان‌نامه کارشناسی ارشد»."""
    exact = {_norm(pattern) for pattern in patterns if not pattern.endswith("*")}
    prefixes = tuple(_norm(pattern[:-1]) for pattern in patterns if pattern.endswith("*"))
    for record in records:
        if not _heading(record):
            continue
        text = _norm(record["text"])
        if text in exact or (prefixes and text.startswith(prefixes)):
            return record
    return None


def _make_issue(rule_id, message, source, record=None, span=None, suggestion=""):
    issue = Issue(rule_id, "warn", message + "\nمستند: " + source,
                  loc=record["location"] if record else "ساختار سند",
                  confidence="B", category="university_profile", suggestion=suggestion)
    if record:
        issue.par_index = record["par_index"]
        issue.paragraph_index = record["par_index"]
        issue.block_id = record["block_id"]
        issue.source_part = record["source_part"]
        issue.xml_path = record["xml_path"]
        issue.story = record["story"]
        issue.page_index = record["page_index"]
        issue.extraction_status = record["extraction_status"]
        if span:
            issue.char_start, issue.char_end = span
            issue.target_text = record["text"][span[0]:span[1]]
            issue.validate_evidence(record["text"])
        else:
            issue.snippet = record["text"][:160]
    return issue


def _leading_image(doc, records):
    """The official basmala page is calligraphy: an image before any text."""
    if doc is None or not records or records[0]["par_index"] < 0:
        return None
    for index, paragraph in enumerate(doc.paragraphs[:records[0]["par_index"]]):
        if paragraph._p.xpath(".//w:drawing|.//w:pict"):
            return {"text": "[تصویر آغاز سند]", "location": f"پاراگراف {index + 1}", "par_index": index}
    return None


def _checklist(records, degree, scope, issues, doc=None):
    english_cover, cover_context = _english_cover(records)
    persian_cover = _persian_cover(records)
    if degree in {"master", "phd"}:
        # The official PhD template (Bahman 1398) has the master's sequence.
        definitions = _MASTER
    else:
        definitions = [_MASTER[6], _MASTER[12], _MASTER[14]]
    rows = []
    all_text_readable = bool(records) and not any(r["extraction_status"].startswith(("partial:", "no_reliable", "analysis_limited")) for r in records)
    for item_id, label, patterns, requirement in definitions:
        if item_id == "english_cover":
            found = english_cover
        elif item_id == "persian_cover":
            found = persian_cover
        elif item_id == "basmala":
            found = _find_heading(records, patterns) or _leading_image(doc, records)
        elif item_id == "chapters":
            found = next((r for r in records if _heading(r) and _CHAPTER.fullmatch(_norm(r["text"]))), None)
        else:
            found = _find_heading(records, patterns)
        source = _source(degree, "۵.۱" if degree == "master" else "۵.۲" if item_id.endswith("cover") else "۵.۴")
        row = {"id": item_id, "label": label, "source": source,
               "status": "manual", "detail": "با دادهٔ موجود قابل تأیید نیست؛ بازبینی انسانی لازم است."}
        if found:
            row.update(status="detected", evidence=found["text"][:160], location=found["location"],
                       par_index=found.get("par_index", -1),
                       detail="نشانهٔ متنی یافت شد؛ ترتیب صفحات، کامل‌بودن محتوا و صفحه‌آرایی تأیید نشده است.")
        elif requirement == "optional":
            row.update(status="optional", detail="نبودن این بخش به‌تنهایی مغایرت نیست؛ تقدیم اختیاری و پیوست وابسته به وجود محتوای مربوط است.")
        elif scope == "fragment" or degree == "unknown" or not all_text_readable:
            row["detail"] = "مقطع نامشخص، محدودهٔ جزئی یا متن ناکافی است؛ دربارهٔ نبودن این بخش در کل سند نمی‌توان داوری کرد."
        elif requirement == "conditional":
            row["detail"] = "لزوم این فهرست و کامل‌بودن آن با توجه به محتوای واقعی سند باید دستی بررسی شود."
        else:
            row.update(status="not_detected", detail="عنوان یا نشانهٔ مستقل در متن بررسی‌شده تشخیص داده نشد؛ این نتیجه اثبات نبودن صفحه نیست.")
            issues.append(_make_issue("KASHAN_SECTION_REVIEW", f"نشانهٔ مستقل بخش «{label}» در متن استخراج‌شده تشخیص داده نشد؛ وجود آن و شمول الگو را بررسی کنید.", source,
                                      suggestion="وجود این بخش را در فایل اصلی و الگوی معتبر مقطع بررسی کنید؛ بخش جدید خودکار ساخته نمی‌شود."))
        rows.append(row)
    # Clauses that a Word file cannot prove are listed for the student to check.
    manual = [
        ("cover_layout", "جای آرم، وسط صفحه بودن عنوان و درج عمودی نام استادان راهنما و مشاور",
         "قلم، Bold، تاریخ و «توسط» خودکار بررسی می‌شوند؛ چیدمان دیداری صفحهٔ عنوان را در Word یا PDF نگاه کنید.", GUIDE_A.format("بندهای ۲ تا ۵")),
        ("chapter_page_numbers", "شمارهٔ صفحهٔ ابتدای فصل‌ها",
         "در ابتدای هر فصل شمارهٔ صفحه تایپ نشود ولی شمرده شود (Different First Page). صفحه‌های Abstract و عنوان انگلیسی نیز شماره ندارند.",
         GUIDE_A.format("بند ۱۸ و تذکر صفحهٔ ۲")),
        ("lines_per_page", "۲۴ تا ۲۶ سطر در هر صفحه",
         "با فاصلهٔ خطوط Exactly یا At least خودکار سنجیده می‌شود؛ با Multiple یا Single تعداد سطر یک صفحهٔ کامل را در Word بشمارید.", GUIDE_A.format("بند ۲۹")),
        ("dedication_note", "صفحهٔ تقدیم",
         "راهنمای سه‌صفحه‌ای ایجاد آن را می‌خواهد ولی الگوهای رسمی آن را «اختیاری» نوشته‌اند؛ نبودنش ایراد گرفته نمی‌شود. با گروه آموزشی تطبیق دهید.",
         GUIDE_A.format("بند ۱۰ و الگوهای رسمی")),
        ("reference_style_note", "شیوهٔ فهرست منابع",
         "الگوهای رسمی منابع را به ترتیب الفبا و به روش APA می‌خواهند؛ راهنمای سه‌صفحه‌ای شماره‌گذاری در کروشه را آورده است. اگر منابع شماره دارند، کروشه و رقم فارسی/انگلیسی خودکار بررسی می‌شود.",
         GUIDE_A.format("بند ۱۹ و الگوهای رسمی")),
        ("signatures", "صورت‌جلسهٔ دفاع و تعهدنامهٔ امضاشده",
         "صورت‌جلسه پس از عنوان فارسی و تعهدنامهٔ امضاشده (بدون تاریخ و شماره) پس از آن؛ امضا و اعتبار فرم از روی فایل قابل تأیید نیست.",
         GUIDE_A.format("بندهای ۸ و ۹")),
        ("chapter_count", "تعداد و ترتیب فصل‌ها", "نمونهٔ سه‌فصلی الزام تعداد ثابت نیست؛ تأیید ساختار مصوب رشته و ترتیب واقعی با نویسنده است.",
         _source(degree, "۵.۲ و ۵.۳")),
        ("toc_pages", "تطبیق فهرست‌ها با عنوان‌ها و شمارهٔ صفحات", "تشخیص عنوان از متن، تطبیق صفحه‌بندی Word/PDF و به‌روزبودن فهرست را ثابت نمی‌کند.",
         _source(degree, "۵.۳")),
        ("references_integrity", "تطبیق ارجاعات و کامل‌بودن منابع", "فقط الگوهای روشن محلی بررسی می‌شوند؛ هیچ نویسنده، سال، DOI یا اطلاعات منبعی از روی حدس تکمیل نمی‌شود.",
         _source(degree, "۵.۴")),
        ("authenticity", "اصالت پژوهش، صحت داده‌ها و امضاها", "از خواندن فایل نمی‌توان اصالت، نبود سرقت علمی، صحت داده‌ها یا امضاهای معتبر را تضمین کرد.",
         _source(degree, "۵.۵")),
        ("before_cd", "پیش از تهیهٔ CD و ثبت در ایرانداک",
         "از کارشناس تحصیلات تکمیلی نامهٔ انجام نظرات اصلاحی داوران، صورت‌جلسهٔ دفاع و فرم تعهدنامه را دریافت کنید.", GUIDE_A.format("مقدمه")),
        ("cd", "تحویل CD",
         "دو CD قاب‌دار (برای دو استاد راهنما سه CD): یکی کتابخانهٔ دانشکده، یکی استاد راهنما. عنوان فارسی روی جلد و انگلیسی پشت جلد. "
         "داخل CD: فایل Word و PDF یکپارچه از بسم‌الله تا عنوان انگلیسی، اسکن JPG صورت‌جلسه و اسکن تعهدنامهٔ امضاشده.", GUIDE_A.format("نکات تهیهٔ CD")),
        ("irandoc", "ثبت در ایرانداک",
         "ثبت و بارگذاری Word و PDF پروپوزال و پایان‌نامه در sabt.irandoc.ac.ir بدون پرداخت هزینه و دریافت کد رهگیری ۷ رقمی "
         "(پروپوزال با ۱ و پایان‌نامه با ۲ آغاز می‌شود). تسویه‌حساب منوط به تأیید این کدهاست.", GUIDE_A.format("نکات ثبت در ایرانداک")),
    ]
    if degree == "phd":
        manual.append(("phd_defense_panel", "فرم دفاع دکتری و داور خارج دانشگاه", "الگوی رسمی دکتری داور خارج دانشگاه دارد؛ نقش، هویت و امضاها باید با فرم معتبر تطبیق داده شود.",
                       _source(degree, "۵.۲")))
    if degree == "unknown":
        manual.insert(0, ("degree", "انتخاب مقطع تحصیلی", "ابتدا کارشناسی ارشد یا دکتری را انتخاب کنید؛ الگو و عنوان مقطع حدس زده نشده است.",
                          _source(degree, "۳ و ۵")))
    for item_id, label, detail, source in manual:
        rows.append({"id": item_id, "label": label, "status": "manual", "detail": detail, "source": source})
    return rows, cover_context


def _check_cover(degree, context, issues):
    if degree == "unknown" or not context:
        return
    wrong = r"\bThesis\b" if degree == "phd" else r"\bDissertation\b"
    wanted = "Ph.D. Dissertation" if degree == "phd" else "Thesis"
    for record in context:
        text = record["text"]
        if len(text.split()) > 25 or not _heading(record):
            continue
        match = re.search(wrong, text, re.I)
        if match:
            issues.append(_make_issue("KASHAN_ENGLISH_DEGREE", f"در محدودهٔ احتمالی صفحهٔ عنوان انگلیسی، عبارت «{match.group()}» با مقطع انتخاب‌شده هماهنگ نیست؛ الگوی {wanted} را بازبینی کنید.",
                                      _source(degree, "۵.۲" if degree == "phd" else "۵.۱"), record, match.span(),
                                      f"فقط در صفحهٔ عنوان انگلیسی و پس از تأیید مقطع، عبارت «{wanted}» را با الگوی معتبر تطبیق دهید."))
            break


def _check_persian_cover(degree, records, issues):
    """Flag only an explicit opposite degree on a plausible Persian cover.

    A university marker alone is insufficient: require a nearby standalone
    faculty/supervisor label and an unquoted degree title line. Neither a cover
    mention in bibliography nor prose discussing another thesis is evidence.
    """
    if degree not in {"master", "phd"}:
        return
    wrong = (r"پایان[\s\u200c-]*نامه\s+کارشناس[یي]\s+ارشد"
             if degree == "phd" else r"رساله\s+د[کك]تر[یي]")
    wanted = "رساله دکتری" if degree == "phd" else "پایان‌نامه کارشناسی ارشد"
    for marker in records:
        if not _heading(marker) or _norm(marker["text"]) != "دانشگاه کاشان":
            continue
        window = records[max(0, marker["index"] - 3):marker["index"] + 10]
        context = [record for record in window
                   if record["story"] == marker["story"] and not _is_toc(record)]
        if _find_heading(context, ("منابع", "مراجع", "منابع و مراجع", "منابع و مآخذ", "references", "bibliography")):
            continue
        positive_context = any(
            _heading(record) and re.match(r"^(?:دانشکده\s+|استاد(?:ان)?\s+(?:راهنما|مشاور)(?:\s|:|$))", _norm(record["text"]))
            for record in context)
        if not positive_context:
            continue
        for record in context:
            if not _heading(record):
                continue
            match = re.fullmatch(r"\s*(?P<degree>" + wrong + r")(?:\s+(?:در\s+رشته|رشته|گرایش)\s+[^«»\"\n]+)?\s*", record["text"])
            if not match:
                continue
            issues.append(_make_issue("KASHAN_PERSIAN_DEGREE",
                f"در محدودهٔ احتمالی صفحهٔ عنوان فارسی، عبارت «{match.group('degree')}» با مقطع انتخاب‌شده هماهنگ نیست؛ عنوان {wanted} را بازبینی کنید.",
                _source(degree, "۵.۲" if degree == "phd" else "۵.۱"), record, match.span("degree"),
                f"پس از تأیید مقطع، فقط عنوان مقطع صفحهٔ فارسی را با عبارت «{wanted}» و الگوی معتبر تطبیق دهید."))
            return


_YEAR = re.compile(r"\(\s*(?:(?:1[3-9]|20)[0-9۰-۹]{2}|[۱۲۱][۰-۹]{3}|n\.?\s*d\.?|بی[\s\u200c-]*تا|بدون تاریخ)(?:[a-zالفبج]*)\s*\)", re.I)
_LATIN_AUTHOR = re.compile(r"^\s*(?P<author>[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{1,65}),\s*[A-ZÀ-ÖØ-Þ]\.")
_FA_AUTHOR = re.compile(r"^\s*(?P<author>[\u0621-\u063a\u0641-\u064a\u067e\u0686\u0698\u06a9\u06af\u06cc\u200c -]{2,65})[،,]\s*[\u0621-\u06ff]")


def _author_sort_key(name, script):
    if script == "latin":
        return "".join(ch for ch in unicodedata.normalize("NFD", _norm(name))
                       if not unicodedata.combining(ch))
    # Unicode codepoint order puts Persian پ/چ/ژ/گ after ی. Use the actual
    # Persian alphabet, and leave unsupported characters unassessed.
    alphabet = "ابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی"
    aliases = str.maketrans({"آ": "ا", "أ": "ا", "إ": "ا", "ؤ": "و", "ئ": "ی", "ة": "ه"})
    chars = [ch for ch in _norm(name).translate(aliases)
             if ch not in " -" and not unicodedata.combining(ch)]
    if any(ch not in alphabet for ch in chars):
        return None
    return tuple(alphabet.index(ch) for ch in chars)


def _references(records, degree, issues):
    start = _find_heading(records, ("منابع", "مراجع", "منابع و مراجع", "منابع و مآخذ", "references", "bibliography"))
    if start is None:
        return {"status": "manual", "detail": "عنوان مستقل منابع تشخیص داده نشد؛ ارزیابی منابع کامل نیست."}
    entries = []
    for record in records[start["index"] + 1:]:
        if (_find_heading([record], ("پیوست", "پیوست ها", "ضمائم", "appendix", "appendices", "abstract", "چکیده انگلیسی"))
                or (_heading(record) and _CHAPTER.fullmatch(_norm(record["text"])))
                or (_heading(record) and re.fullmatch(r"university\s+of\s+kashan|kashan\s+university", record["text"].strip(), re.I))):
            break
        if not _is_toc(record) and record["story"] in {"body", "text", "pdf"}:
            entries.append(record)
    source = _source(degree, "۵.۴")
    seen = {}
    parsed = []
    uncertain = False
    for record in entries:
        text = record["text"]
        author = _LATIN_AUTHOR.match(text) or _FA_AUTHOR.match(text)
        year = _YEAR.search(text.translate(_DIGITS))
        if not author:
            uncertain = True
            continue
        if not year:
            uncertain = True
            if len(text.split()) >= 6 and len(text) >= 35:
                issues.append(_make_issue("KASHAN_REFERENCE_YEAR", "در این مدخل با الگوی روشن نام نویسنده، سال داخل پرانتز یا نشانگر بی‌تاریخ شناخته نشد؛ کامل‌بودن منبع را بررسی کنید.", source, record,
                                          author.span("author"), "سال یا بی‌تاریخ‌بودن را از خود منبع بررسی کنید؛ اطلاعات از روی حدس افزوده نشود."))
            continue
        normalized = _norm(text)
        if len(text) >= 30 and normalized in seen:
            issues.append(_make_issue("KASHAN_REFERENCE_DUPLICATE", "این مدخل با مدخل پیشین در همین بخش از نظر متن تکراری است؛ تکرار احتمالی را بررسی کنید.", source, record,
                                      (0, len(text)), "با منبع اصلی و مدخل قبلی تطبیق دهید؛ حذف خودکار انجام نمی‌شود."))
        else:
            seen[normalized] = record
        name = author.group("author")
        script = "latin" if re.search(r"[A-Za-z]", name) else "persian"
        sort_key = _author_sort_key(name, script)
        if sort_key is None:
            uncertain = True
        parsed.append((script, sort_key, record, author.span("author")))
    # Sorting can only be assessed if every paragraph is a complete recognizable
    # entry. Wrapped paragraphs, mixed conventions and unknown names stay manual.
    if len(parsed) >= 2 and not uncertain:
        previous = {}
        for script, name, record, span in parsed:
            if script in previous and name < previous[script]:
                issues.append(_make_issue("KASHAN_REFERENCE_ALPHABETICAL", "ترتیب الفبایی نام خانوادگی در دو مدخل پیاپیِ هم‌زبان نیازمند بازبینی است؛ راهنمای کاشان APA و ترتیب الفبایی را ذکر می‌کند، بدون تعیین ویرایش APA.", source, record, span,
                                          "ترتیب نام نویسندگان را با شیوه‌نامهٔ معتبر و نویسه‌های زبان منبع بررسی کنید."))
            previous[script] = name
    return {"status": "partial" if parsed else "manual", "recognized_entries": len(parsed),
            "alphabetical_checked": len(parsed) >= 2 and not uncertain,
            "style": "APA", "edition": None, "online_verified": False,
            "detail": "بررسی محلی و محتاطانهٔ نام/سال، تکرار و ترتیب الفبایی؛ صحت یا کامل‌بودن اطلاعات کتاب‌شناختی تأیید نشده است."}


@bounded_findings
def check_document(doc, blocks, stats, rules):
    """Return source-labelled B findings; add auditable metadata to stats only."""
    if rules.get("university_profile_id") != "kashan":
        return []
    degree = rules.get("university_degree", "unknown")
    if degree not in _DEGREES:
        degree = "unknown"
    scope = "fragment" if rules.get("document_scope") == "fragment" or stats.get("is_raw_text") else "full"
    records = _records(doc, blocks)
    if not records or stats.get("analysis_status") in {"failed", "partial"}:
        scope = "fragment"
    metadata = rules.get("kashan_profile") or _profile_metadata()
    stats["university_profile"] = {
        "id": "kashan", "title": "دانشگاه کاشان", "degree": degree,
        "degree_label": _DEGREES[degree], "source_status": "owner_provided_originals",
        "source_label": SOURCE_LABEL,
        "applied_sources": ["A", "B"] if degree == "master" else ["A", "C"] if degree == "phd" else ["A", "B", "C"],
        "excluded_sources": [],
        "sources": copy.deepcopy(metadata.get("sources", [])),
        "reference_style": "APA؛ ویرایش تعیین نشده", "document_scope": scope,
        "limitations": [SOURCE_LABEL,
                        "چیدمان دیداری صفحه‌ها، امضاها، CD و ایرانداک از روی فایل قابل تأیید نیستند و در چک‌لیست برای بررسی دستی آمده‌اند.",
                        "نشانه‌های متنی، تأیید صفحه‌آرایی، امضا، ترتیب صفحات، اصالت پژوهش یا آمادگی نهایی نیستند.",
                        "پیشنهادها خودکار اعمال نمی‌شوند؛ سند اصلی و داده‌های علمی تغییر نمی‌کنند."],
    }
    issues = FindingBuffer()
    # A PDF block can combine a heading and prose, and extracted reading order
    # may be uncertain. It can establish textual presence, never absence of an
    # entire required page. Only native Word full-document review emits those
    # bounded, explicitly qualified missing-heading suggestions.
    checklist_scope = scope if doc is not None and not stats.get("is_pdf") else "fragment"
    checklist, cover_context = _checklist(records, degree, checklist_scope, issues, doc)
    stats["university_checklist"] = checklist
    _check_cover(degree, cover_context, issues)
    _check_persian_cover(degree, records, issues)
    if not stats.get("is_pdf"):
        import kashan_format
        kashan_format.check_format(doc, records, degree, checklist_scope, checklist, cover_context, issues)
    stats["university_reference_checks"] = _references(records, degree, issues)
    by_id = {record["block_id"]: record for record in records if record["block_id"]}
    for issue in issues:
        if issue.location_status != "exact":
            continue
        record = by_id.get(issue.block_id)
        if record is None:
            continue
        spans = ProtectedContentDetector.find_spans(record["text"])
        configured_overlap = any(
            term and issue.char_start < match.end() and match.start() < issue.char_end
            for term in rules.get("protected_terms", [])
            for match in re.finditer(re.escape(term), record["text"]))
        issue.protected_overlap = (configured_overlap or
            ProtectedContentDetector.is_protected(issue.char_start, issue.char_end, spans))
    return issues
