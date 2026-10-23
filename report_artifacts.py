"""Portable, offline review artifacts. Source content is always treated as data."""
from collections import Counter
import csv
from datetime import datetime, timezone
import hashlib
from html import escape
import io
import json
from pathlib import Path
import re

from doc_guard import atomic_write

SCHEMA_VERSION = "2.0"
PRODUCT_VERSION = "2.3.2"
CONFIDENCE_LABELS = {"A": "اطمینان بالا", "B": "نیازمند بازبینی", "C": "مبهم"}
CATEGORY_LABELS = {
    "general": "بازبینی عمومی", "english_grammar": "دستور زبان انگلیسی",
    "english_style": "سبک نگارش انگلیسی", "academic_style": "سبک نگارش دانشگاهی",
    "university_profile": "شیوه‌نامهٔ دانشگاه", "citation_consistency": "ارجاع‌دهی و فهرست منابع",
    "persian_orthography": "املاء و فاصله‌گذاری فارسی", "document_review": "ساختار و قالب سند",
}
STORY_LABELS = {"body": "متن اصلی", "table": "جدول‌ها", "header": "سرصفحه",
                "footer": "پاصفحه", "footnote": "پاورقی", "endnote": "یادداشت پایانی",
                "textbox": "جعبه‌متن", "pdf": "صفحه‌های PDF", "text": "متن ارسالی"}
COVERAGE_LABELS = {"analyzed": "بررسی شد", "partial": "بخشی بررسی شد",
                   "present": "وجود دارد؛ بررسی نشد", "absent": "در سند وجود نداشت"}
UNIVERSITY_CHECKLIST_LABELS = {
    "detected": "نشانهٔ متنی یافت شد؛ تأیید قالب نیست",
    "not_detected": "در متن استخراج‌شده یافت نشد؛ نیازمند بررسی",
    "manual": "با اطلاعات موجود قابل تأیید نیست؛ بررسی دستی لازم است",
    "optional": "در صورت شمول؛ بررسی نویسنده لازم است",
}
ANNOTATION_REASONS = {
    "ambiguous_excluded": "مورد مبهم؛ فقط برای بررسی دستی در گزارش",
    "protected_content": "تداخل با محتوای حفاظت‌شده", "unsupported_story": "درج امن کامنت در این بخش پشتیبانی نمی‌شود",
    "location_unavailable": "محل دقیق در دسترس نیست", "unsafe_structure": "ساختار این محل برای کامنت‌گذاری ایمن نیست",
    "uncertain_location": "محل دقیق قابل تأیید نیست", "overlapping_target": "هم‌پوشانی با کامنت دیگر",
    "duplicate": "تکراری", "comment_limit": "رسیدن به سقف کامنت‌های سند",
    "paragraph_limit": "رسیدن به سقف کامنت‌های پاراگراف", "annotation_failed": "تولید امن نسخهٔ بازبینی ممکن نشد",
    "dependency_unavailable": "ابزار نشانه‌گذاری PDF در دسترس نیست", "encrypted": "PDF رمزگذاری شده است",
    "repaired_pdf": "ساختار PDF نیازمند ترمیم است", "signed_pdf": "PDF امضای دیجیتال دارد",
    "page_budget": "تعداد صفحه‌ها بیش از سقف نشانه‌گذاری است", "character_budget": "حجم متن بیش از سقف نشانه‌گذاری است",
    "confidence_filtered": "مورد مبهم؛ فقط برای بررسی دستی در گزارش",
    "missing_exact_location": "محل دقیق در صفحه قابل تأیید نیست", "source_mismatch": "محل با صفحهٔ منبع تطبیق ندارد",
    "extraction_mapping_mismatch": "نگاشت متن به نویسه‌های صفحه قابل تأیید نیست",
    "evidence_mismatch": "شاهد با متن صفحه تطبیق ندارد", "grapheme_boundary": "مرز نویسهٔ کامل قابل تأیید نیست",
    "ambiguous_glyph": "نویسهٔ نمایشی با چند حرف متناظر است", "invalid_geometry": "مختصات امن در صفحه در دسترس نیست",
    "no_visible_target": "برای این شاهد محل دیداری قابل تأییدی یافت نشد", "duplicate_target": "نشانه‌گذاری تکراری",
    "not_processed": "نشانه‌گذاری این مورد تکمیل نشد",
}


def category_label(value):
    return CATEGORY_LABELS.get(value, value or "بازبینی عمومی")


def annotation_format_label(stats):
    return "PDF" if stats.get("annotation_format") == "pdf" else "Word"


def location_label(value):
    """Translate display paths only; source XML identities remain unchanged."""
    text = str(value or "بررسی در سطح سند")
    if not re.match(r"^(?:body|table|header|footer|footnote|endnote|textbox|Paragraph|Page)(?:\b| /)", text):
        return text
    tokens = {**STORY_LABELS, "table": "جدول", "paragraph": "پاراگراف", "row": "سطر", "cell": "سلول",
              "Paragraph": "پاراگراف", "Page": "صفحه"}
    return re.sub(r"\b(?:" + "|".join(map(re.escape, tokens)) + r")\b",
                  lambda match: tokens[match.group()], text)


def ambiguous_count(issues, stats, rules):
    visible = sum(issue.confidence == "C" for issue in issues)
    if "ambiguous_candidates" in stats:
        return visible + max(0, int(stats["ambiguous_candidates"]))
    return None if rules.get("strict", True) else visible


def safe_annotation_summary(summary):
    """Export audit fields only, excluding caller paths and arbitrary metadata."""
    if not isinstance(summary, dict):
        return None
    safe = {key: max(0, value) for key in ("requested", "applied", "unapplied", "max_comments",
                                         "exact_applied", "paragraph_applied", "highlighted")
            if isinstance((value := summary.get(key)), int) and not isinstance(value, bool)}
    if "max_comments" in summary and summary["max_comments"] is None:
        safe["max_comments"] = None
    safe["output_created"] = bool(summary.get("output_created", False))
    if isinstance(summary.get("experimental"), bool):
        safe["experimental"] = summary["experimental"]
    reasons = summary.get("reasons", {})
    safe["reasons"] = {key: max(0, value) for key, value in (reasons.items() if isinstance(reasons, dict) else [])
                       if key in ANNOTATION_REASONS and isinstance(value, int) and not isinstance(value, bool)}
    safe["records"] = []
    numeric = {"index", "par_index", "page_index", "char_start", "char_end"}
    textual = {"rule_id", "confidence", "block_id", "source_part", "xml_path", "story", "status", "reason"}
    records = summary.get("records", [])
    for record in records if isinstance(records, list) else []:
        if not isinstance(record, dict):
            continue
        row = {key: value for key, value in record.items()
               if (key in numeric and isinstance(value, int) and not isinstance(value, bool))
               or (key in textual and isinstance(value, str))}
        if isinstance(record.get("anchor_scope"), str) and record["anchor_scope"] in {"exact", "paragraph"}:
            row["anchor_scope"] = record["anchor_scope"]
        safe["records"].append(row)
    return safe


def safe_review_document_summary(summary):
    """Export companion Word counts, never its temporary path or source content."""
    if not isinstance(summary, dict):
        return None
    safe = {"created": bool(summary.get("created", False))}
    safe.update({key: max(0, value) for key in ("findings", "comments", "highlighted")
                 if isinstance((value := summary.get(key)), int) and not isinstance(value, bool)})
    if summary.get("filetype") == "docx":
        safe["filetype"] = "docx"
    return safe


def safe_unified_word_review(summary):
    """Export existing-text annotation counts and known context-anchor identities."""
    if not isinstance(summary, dict):
        return None
    safe = {"output_created": summary.get("output_created") is True}
    safe.update({key: max(0, value) for key in (
        "requested", "source_comments", "context_comments", "appendix_comments", "comments", "highlighted",
        "exact_comments", "paragraph_comments")
        if isinstance((value := summary.get(key)), int) and not isinstance(value, bool)})
    if isinstance(summary.get("context_records"), list):
        safe["context_records"] = []
        for record in summary["context_records"]:
            if not isinstance(record, dict):
                continue
            index = record.get("index")
            if not isinstance(index, int) or isinstance(index, bool) or index < 0:
                continue
            row = {"index": index}
            paragraph = record.get("anchor_par_index")
            if isinstance(paragraph, int) and not isinstance(paragraph, bool) and paragraph >= -1:
                row["anchor_par_index"] = paragraph
            path = record.get("anchor_xml_path")
            if isinstance(path, str) and re.fullmatch(
                    r"/w:document/w:body(?:/w:[A-Za-z][A-Za-z0-9]*(?:\[[1-9][0-9]*\])?)+", path):
                row["anchor_xml_path"] = path
            if record.get("scope") in ("paragraph", "document", "story"):
                row["scope"] = record["scope"]
            story = record.get("source_story")
            if isinstance(story, str) and story in {"", *STORY_LABELS}:
                row["source_story"] = story
            if isinstance(record.get("reason"), str) and record["reason"] in ANNOTATION_REASONS:
                row["reason"] = record["reason"]
            safe["context_records"].append(row)
    return safe


def _unified_word_review_complete(summary):
    return (summary is not None and summary.get("output_created")
            and all(key in summary for key in ("requested", "comments", "highlighted",
                                               "source_comments", "context_comments"))
            and summary["comments"] == summary["requested"] == summary["highlighted"]
            and summary["source_comments"] + summary["context_comments"] == summary["requested"]
            and summary.get("appendix_comments", 0) == 0)


def annotation_summary_lines(stats):
    """Separate certified source locations from qualified existing-text anchors."""
    lines = []
    summary = safe_annotation_summary(stats.get("annotation_summary"))
    unified = safe_unified_word_review(stats.get("unified_word_review"))
    if unified is not None:
        if unified["output_created"]:
            lines.append("همان Word بازبینی: از " + str(unified.get("requested", 0)) + " مورد، "
                         + str(unified.get("comments", 0)) + " مورد کامنت و "
                         + str(unified.get("highlighted", 0)) + " مورد هایلایت زرد دارد.")
            lines.append("محل ثبت در همان متن موجود: " + str(unified.get("source_comments", 0))
                         + " مورد در محل قابل تأیید و " + str(unified.get("context_comments", 0))
                         + " مورد با کامنت کلی/زمینه‌ای روی متن موجود.")
            if not _unified_word_review_complete(unified):
                lines.append("ثبت کامنت و هایلایت همهٔ موارد در همین Word تکمیل نشده است؛ تعدادهای بالا را بررسی کنید.")
            if unified.get("context_comments", 0):
                lines.append("کامنت کلی/زمینه‌ای در سطح پاراگراف، بخش یا سند است؛ عبارت زردشده لزوماً خطا نیست "
                             "و محل دقیق ایراد را تأیید نمی‌کند. نوع ارجاع داخل کامنت مشخص شده است.")
        else:
            lines.append("تولید Word یکپارچه تکمیل نشد؛ ثبت همهٔ موارد در همین فایل تأیید نمی‌شود و جزئیات در گزارش باقی است.")
    if summary is not None:
        label = annotation_format_label(stats)
        if unified is not None and label == "Word":
            lines.append("ممیزی محل‌های قابل تأیید: از " + str(summary.get("requested", 0)) + " مورد، "
                         + str(summary.get("applied", 0)) + " مورد در محل قابل تأیید کامنت شد.")
        else:
            lines.append("از " + str(summary.get("requested", 0)) + " مورد، "
                         + str(summary.get("applied", 0)) + " مورد در " + label + " کامنت شد.")
        if label == "Word" and any(key in summary for key in ("exact_applied", "paragraph_applied")):
            lines.append("در نسخهٔ بازبینیِ سند اصلی: " + str(summary.get("exact_applied", 0))
                         + " کامنت در محدودهٔ دقیق و " + str(summary.get("paragraph_applied", 0))
                         + " کامنت در سطح پاراگراف ثبت شد.")
            if summary.get("paragraph_applied", 0):
                lines.append("کامنت سطح پاراگراف به کل پاراگراف مربوط است؛ محل دقیق یک عبارت را نشان نمی‌دهد.")
        if "highlighted" in summary and unified is None:
            lines.append("تعداد مواردِ دارای کامنت و هایلایت زرد: " + str(summary["highlighted"]) + ".")
        if "max_comments" in summary:
            limit_label = "سقف کامنت در محل‌های قابل تأیید: " if unified is not None else "سقف تعداد کامنت: "
            lines.append(limit_label + ("بدون سقف تعداد" if summary["max_comments"] is None
                                        else str(summary["max_comments"])) + ".")
        if not summary.get("output_created"):
            lines.append("تولید نسخهٔ بازبینی تکمیل نشد؛ جزئیات در گزارش باقی است.")
        if summary.get("experimental"):
            lines.append("نشانه‌گذاری PDF آزمایشی است؛ نمایش هایلایت و یادداشت فارسی در PDFخوان شما باید بررسی شود.")
        if unified is not None and summary.get("reasons"):
            lines.append("علت انتخاب کامنت کلی/زمینه‌ای به‌جای محل دقیق برای بعضی موارد:")
        for reason, count in summary.get("reasons", {}).items():
            lines.append(ANNOTATION_REASONS[reason] + ": " + str(count))
    review = safe_review_document_summary(stats.get("review_document"))
    if review is not None:
        if review["created"]:
            lines.append("Word گزارش کامل: " + str(review.get("findings", 0)) + " مورد، "
                         + str(review.get("comments", 0)) + " کامنت و "
                         + str(review.get("highlighted", 0)) + " مورد هایلایت زرد.")
            if any(review.get(key, 0) < review.get("findings", 0) for key in ("comments", "highlighted")):
                lines.append("کامنت و هایلایت همهٔ موارد در Word گزارش کامل تکمیل نشده است؛ تعدادهای بالا را بررسی کنید.")
            lines.append("Word گزارش کامل، سند مستقلی از یافته‌هاست؛ محل کامنت‌های آن، محل دقیق ایراد در سند اصلی نیست.")
        else:
            lines.append("تولید Word گزارش کامل انجام نشد؛ همهٔ یافته‌های گزارش‌شده در گزارش متنی باقی‌اند.")
    if summary is not None or review is not None or unified is not None:
        lines.append("موارد سطح C فقط برای بررسی دستی‌اند؛ کامنت یا هایلایت، اطمینان آن‌ها را افزایش نمی‌دهد.")
        lines.append("برای محتوای حفاظت‌شده یا محل نامطمئن، محل دقیقِ تأییدنشده در سند اصلی ساخته نمی‌شود.")
    return lines


def safe_coverage(coverage):
    if not isinstance(coverage, dict):
        return {}
    result = {}
    for story, info in coverage.items():
        if story not in STORY_LABELS or not isinstance(info, dict):
            continue
        safe = {key: max(0, value) for key in ("present_blocks", "analyzed_blocks", "partial_blocks")
                if isinstance((value := info.get(key)), int) and not isinstance(value, bool)}
        safe["status"] = info.get("status") if info.get("status") in COVERAGE_LABELS else "partial"
        result[story] = safe
    return result


def safe_university_profile(profile):
    """Keep the documented profile contract, never arbitrary nested configuration."""
    if not isinstance(profile, dict) or profile.get("id") != "kashan":
        return None
    textual = {"id", "title", "degree_label", "source_label", "reference_style"}
    safe = {key: value for key, value in profile.items() if key in textual and isinstance(value, str)}
    degree = profile.get("degree")
    scope = profile.get("document_scope")
    safe["degree"] = degree if isinstance(degree, str) and degree in {"master", "phd", "unknown"} else "unknown"
    safe["document_scope"] = scope if isinstance(scope, str) and scope in {"full", "fragment"} else "fragment"
    # This release implements a supplied summary, not independently authenticated guides.
    safe["source_status"] = "user_provided_summary"
    for key in ("applied_sources", "limitations"):
        values = profile.get(key, [])
        safe[key] = [value for value in values if isinstance(value, str)] if isinstance(values, list) else []
    safe["applied_sources"] = [value for value in safe["applied_sources"] if value in {"A", "B", "C"}]
    safe["excluded_sources"] = []
    for source in profile.get("excluded_sources", []) if isinstance(profile.get("excluded_sources"), list) else []:
        if isinstance(source, dict):
            safe["excluded_sources"].append({key: value for key, value in source.items()
                                              if key in {"id", "reason"} and isinstance(value, str)})
    safe["sources"] = []
    source_keys = {"id", "title", "date", "section", "status", "scope", "note", "kind", "reason"}
    for source in profile.get("sources", []) if isinstance(profile.get("sources"), list) else []:
        if isinstance(source, dict):
            item = {key: value for key, value in source.items()
                    if key in source_keys and isinstance(value, str)}
            item.update({key: value for key in ("verified", "applies")
                         if isinstance((value := source.get(key)), bool)})
            if isinstance(source.get("applicable_degrees"), list):
                item["applicable_degrees"] = [degree for degree in source["applicable_degrees"]
                                               if isinstance(degree, str) and degree in {"master", "phd"}]
            safe["sources"].append(item)
    return safe


def safe_university_checklist(checklist):
    """Checklist observations describe evidence, never university approval."""
    safe = []
    for item in checklist if isinstance(checklist, list) else []:
        if not isinstance(item, dict):
            continue
        row = {key: value for key, value in item.items()
               if key in {"id", "label", "detail", "source", "evidence", "location"} and isinstance(value, str)}
        status = item.get("status")
        row["status"] = status if isinstance(status, str) and status in UNIVERSITY_CHECKLIST_LABELS else "manual"
        safe.append(row)
    return safe


def safe_university_reference_checks(checks):
    """Export only the implemented local bibliography checks, not private parsing state."""
    if not isinstance(checks, dict):
        return None
    status = checks.get("status")
    safe = {"status": status if isinstance(status, str) and status in {"partial", "manual"} else "manual"}
    if isinstance(checks.get("detail"), str):
        safe["detail"] = checks["detail"]
    count = checks.get("recognized_entries")
    if isinstance(count, int) and not isinstance(count, bool):
        safe["recognized_entries"] = max(0, count)
    for key in ("alphabetical_checked", "online_verified"):
        if isinstance(checks.get(key), bool):
            safe[key] = checks[key]
    if checks.get("style") == "APA":
        safe["style"] = "APA"
    if "edition" in checks and checks["edition"] is None:
        safe["edition"] = None
    return safe


def university_profile_lines(stats):
    """Plain-text profile section shared with the TXT report."""
    profile = safe_university_profile(stats.get("university_profile"))
    if profile is None:
        return []
    lines = ["", "پروفایل دانشگاهی: " + profile.get("title", "دانشگاه کاشان"),
             "مقطع: " + profile.get("degree_label", "مشخص نشده"),
             "مرجع: خلاصهٔ راهنماهای ارسالی کاربر؛ اعتبار رسمی و به‌روز بودن مستقلاً تأیید نشده است.",
             "انتخاب این پروفایل به معنی تأیید دانشگاه، آمادگی نهایی یا صحت علمی سند نیست."]
    if profile.get("reference_style"):
        lines.append("سبک ارجاع: " + profile["reference_style"])
    lines.append("دامنهٔ سند: " + ("فایل کامل" if profile["document_scope"] == "full" else "بخش یا متن محدود؛ نبود یک عنوان الزاماً نقص سند کامل نیست"))
    lines.extend("محدودیت: " + value for value in profile["limitations"])
    lines.append("منابع شیوه‌نامه؛ شناسه‌های منبع با سطح اطمینان A/B/C ایرادها متفاوت‌اند:")
    for source in profile["sources"]:
        identity = "منبع " + source.get("id", "نامشخص") + ": " + source.get("title", "عنوان ارائه نشده")
        details = [source[key] for key in ("date", "section", "status", "note", "reason") if source.get(key)]
        lines.append(identity + (" — " + "؛ ".join(details) if details else ""))
    if profile["applied_sources"]:
        lines.append("منابع منتخب این مقطع: " + "، ".join(profile["applied_sources"]))
    for source in profile["excluded_sources"]:
        lines.append("منبع اجرا‌نشده " + source.get("id", "نامشخص") + ": " + source.get("reason", "دامنهٔ کاربرد تأیید نشده"))
    reference_checks = safe_university_reference_checks(stats.get("university_reference_checks"))
    if reference_checks is not None:
        status = "بررسی محدود و محلی" if reference_checks["status"] == "partial" else "با اطلاعات موجود قابل تأیید نیست؛ بررسی دستی لازم است"
        lines.append("وضعیت بررسی منابع: " + status)
        if reference_checks.get("detail"):
            lines.append("جزئیات بررسی منابع: " + reference_checks["detail"])
        if "recognized_entries" in reference_checks:
            lines.append("مدخل‌های دارای الگوی قابل تشخیص نام/سال: " + str(reference_checks["recognized_entries"]))
        if "alphabetical_checked" in reference_checks:
            lines.append("ترتیب الفبایی: " + ("فقط برای مدخل‌های قابل تشخیص، محلی بررسی شد" if reference_checks["alphabetical_checked"] else "قابل بررسی مطمئن نبود؛ تأیید نشده است"))
        lines.append("این بررسی، انطباق کامل با APA یا صحت اطلاعات منابع را تأیید نمی‌کند؛ ویرایش APA تعیین نشده است.")
        if reference_checks.get("online_verified") is False:
            lines.append("راستی‌آزمایی آنلاین منابع انجام نشده است.")
    checklist = safe_university_checklist(stats.get("university_checklist"))
    if checklist:
        lines.append("چک‌لیست شواهد و بررسی دستی؛ یافتن متن، امضا، قالب یا صحت محتوا را تأیید نمی‌کند:")
    for item in checklist:
        lines.append("• " + item.get("label", item.get("id", "مورد چک‌لیست")) + ": " + UNIVERSITY_CHECKLIST_LABELS[item["status"]])
        for key, label in (("detail", "توضیح"), ("source", "مستند"), ("location", "محل"), ("evidence", "شاهد متنی")):
            if item.get(key):
                lines.append("  " + label + ": " + item[key])
    return lines


def _university_profile_html(stats):
    profile = safe_university_profile(stats.get("university_profile"))
    if profile is None:
        return ""
    # Use escaped data in text nodes only; no profile data becomes HTML or a URL.
    lines = university_profile_lines(stats)
    lead = ''.join('<p>' + escape(line) + '</p>' for line in lines[2:]
                   if not line.startswith(("• ", "  ")) and line and not line.startswith("چک‌لیست"))
    checklist = safe_university_checklist(stats.get("university_checklist"))
    rows = []
    for item in checklist:
        details = ''.join('<p><strong>' + label + ': </strong><span dir="auto">' + escape(item[key]) + '</span></p>'
                          for key, label in (("detail", "توضیح"), ("source", "مستند"), ("location", "محل"), ("evidence", "شاهد متنی")) if item.get(key))
        rows.append('<tr><th scope="row">' + escape(item.get("label", item.get("id", "مورد چک‌لیست")))
                    + '</th><td>' + escape(UNIVERSITY_CHECKLIST_LABELS[item["status"]]) + details + '</td></tr>')
    table = ('<h3>چک‌لیست شواهد و بررسی دستی</h3><p class="muted">یافتن متن، امضا، قالب یا صحت محتوا را تأیید نمی‌کند.</p>'
             '<table><thead><tr><th>موضوع</th><th>نتیجه و مستند</th></tr></thead><tbody>' + ''.join(rows) + '</tbody></table>') if rows else ''
    return '<section class="notes university-profile"><h2>پروفایل دانشگاهی: ' + escape(profile.get("title", "دانشگاه کاشان")) + '</h2>' + lead + table + '</section>'


def report_data(issues, stats, rules):
    """Export findings and provenance without local filesystem paths or configuration text."""
    encoded_rules = json.dumps(rules, sort_keys=True, ensure_ascii=False).encode("utf-8")
    rows = []
    annotation = safe_annotation_summary(stats.get("annotation_summary"))
    records = {record.get("index"): record for record in (annotation or {}).get("records", [])}
    unified_word_review = safe_unified_word_review(stats.get("unified_word_review"))
    context_records = (unified_word_review or {}).get("context_records", [])
    context_indices = Counter(record["index"] for record in context_records)
    context_by_index = {record["index"]: record for record in context_records
                        if context_indices[record["index"]] == 1}
    for index, issue in enumerate(issues):
        row = issue.to_dict()
        identity = [stats.get("file_sha256") or stats.get("analysis_text_sha256")]
        identity += [row.get(key) for key in ("rule_id", "block_id", "source_part", "xml_path", "par_index", "char_start", "char_end", "message")]
        row["finding_id"] = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode("utf-8")).hexdigest()[:20]
        row["category_label"] = category_label(row.get("category"))
        row["location_label"] = location_label(row.get("loc"))
        record = records.get(index)
        if record is not None and all(record.get(key, row.get(key)) == row.get(key)
                                      for key in ("rule_id", "block_id", "char_start", "char_end")):
            row["annotation"] = {key: record[key] for key in ("status", "reason", "anchor_scope") if key in record}
            if (record.get("status") == "unapplied" and _unified_word_review_complete(unified_word_review)
                    and index in context_by_index):
                row["context_annotation"] = {key: value for key, value in context_by_index[index].items()
                                             if key != "index"}
        rows.append(row)
    # Only known analysis statistics are exported; arbitrary caller metadata stays private.
    keys = {"mode", "words", "body_words", "paragraphs", "tables", "images", "chapters", "total_pages",
            "table_cells_analyzed", "coverage_limits", "file_sha256", "is_pdf", "is_scanned",
            "is_raw_text", "block_count", "story_counts", "partial_blocks", "protected_content_audit",
            "extraction_pages", "analysis_duration_seconds", "skipped_blocks", "source_chars", "analysis_text_sha256",
            "analysis_status", "analysis_error_code", "ambiguous_candidates", "coverage",
            "report_available", "report_retention_minutes", "annotation_format"}
    safe_stats = {key: value for key, value in stats.items() if key in keys}
    if "coverage" in safe_stats:
        safe_stats["coverage"] = safe_coverage(safe_stats["coverage"])
    if annotation is not None:
        safe_stats["annotation_summary"] = annotation
    review_document = safe_review_document_summary(stats.get("review_document"))
    if review_document is not None:
        safe_stats["review_document"] = review_document
    if unified_word_review is not None:
        safe_stats["unified_word_review"] = unified_word_review
    profile = safe_university_profile(stats.get("university_profile"))
    if profile is not None:
        safe_stats["university_profile"] = profile
        safe_stats["university_checklist"] = safe_university_checklist(stats.get("university_checklist"))
        reference_checks = safe_university_reference_checks(stats.get("university_reference_checks"))
        if reference_checks is not None:
            safe_stats["university_reference_checks"] = reference_checks
    counts = dict(Counter(i.confidence for i in issues))
    counts["C"] = ambiguous_count(issues, stats, rules)
    return {
        "schema_version": SCHEMA_VERSION,
        "product_version": PRODUCT_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "configuration_sha256": hashlib.sha256(encoded_rules).hexdigest(),
        "strict": bool(rules.get("strict", True)),
        "summary": {"findings": len(rows), "by_confidence": counts,
                    "by_category": dict(Counter(i.category for i in issues))},
        "stats": safe_stats,
        "limitations": ["این بررسی قاعده‌محور است و صحت محتوای علمی را تأیید نمی‌کند.",
                        "سطح اطمینان احتمال آماری یا درصد دقت نیست.",
                        *(["بررسی سند تکمیل نشد؛ متن قابل‌اعتماد برای ارزیابی در دسترس نبود."]
                          if stats.get("analysis_status") == "failed" else []),
                        *(["بررسی بخشی از سند انجام شد؛ بخش‌های بررسی‌نشده و محدودیت‌ها را در گزارش ببینید."]
                          if stats.get("analysis_status") == "partial" else []),
                        *(["نشانه‌گذاری PDF آزمایشی است؛ نمایش یادداشت فارسی در PDFخوان شما باید بررسی شود."]
                          if annotation and annotation.get("experimental") else []),
                        *stats.get("coverage_limits", [])],
        "findings": rows,
    }


def json_report(data):
    return json.dumps(data, ensure_ascii=False, indent=2) + "\n"


def _csv_cell(value):
    text = "" if value is None else str(value)
    # A spreadsheet must not execute a formula from an author's quoted source text.
    return "'" + text if text.lstrip().startswith(("=", "+", "-", "@")) or text.startswith(("\t", "\r")) else text


def csv_report(data):
    stream = io.StringIO(newline="")
    writer = csv.writer(stream)
    fields = ("finding_id", "rule_id", "rule_version", "category", "confidence", "loc", "story",
              "source_part", "block_id", "char_start", "char_end", "original_text", "message", "suggestion")
    writer.writerow(fields)
    for row in data["findings"]:
        writer.writerow([_csv_cell(row.get(key)) for key in fields])
    return stream.getvalue()


def html_report(data):
    """A self-contained RTL report with no fonts, scripts, analytics or network dependencies."""
    cards = []
    output_label = annotation_format_label(data["stats"])
    unified = safe_unified_word_review(data["stats"].get("unified_word_review"))
    unified_complete = _unified_word_review_complete(unified)
    for row in data["findings"]:
        confidence = row["confidence"]
        original = row.get("original_text")
        evidence = original if original is not None else row.get("snippet", "")
        suggestion = row.get("suggestion")
        annotation = row.get("annotation", {})
        annotation_text = ((f"در {output_label} در سطح پاراگراف کامنت شد؛ محل دقیق عبارت نیست"
                            if annotation.get("anchor_scope") == "paragraph" else f"در {output_label} کامنت شد")
                           if annotation.get("status") == "applied" else
                           ANNOTATION_REASONS.get(annotation.get("reason"), f"در {output_label} کامنت نشد")) if annotation else ""
        if unified_complete and annotation.get("status") == "unapplied":
            annotation_text += "؛ کامنت کلی/زمینه‌ای روی متن موجود؛ عبارت زردشده لزوماً خطا نیست"
            context = row.get("context_annotation", {})
            paragraph_index = context.get("anchor_par_index", -1)
            if isinstance(paragraph_index, int) and not isinstance(paragraph_index, bool) and paragraph_index >= 0:
                annotation_text += "؛ محل کامنت زمینه‌ای: پاراگراف " + str(paragraph_index + 1)
            elif "/w:tbl" in context.get("anchor_xml_path", ""):
                annotation_text += "؛ محل کامنت زمینه‌ای: پاراگرافی از جدول در همین سند"
        cards.append(
            '<article class="finding" data-confidence="' + escape(confidence, quote=True) + '">'
            '<div class="finding-top"><span class="badge ' + escape(confidence, quote=True) + '">'
            + escape(CONFIDENCE_LABELS.get(confidence, confidence)) + '</span><code>'
            + escape(row["rule_id"]) + '</code><span class="category">' + escape(row.get("category_label", category_label(row.get("category")))) + '</span></div>'
            '<h3>' + escape(row.get("message", "")) + '</h3>'
            '<p class="location">' + escape(row.get("location_label", location_label(row.get("loc")))) + '</p>'
            + ('<p class="muted">وضعیت ' + output_label + ': ' + escape(annotation_text) + '</p>' if annotation_text else '')
            + ('<p class="evidence-label">' + ("شاهد دقیق" if original is not None else "متن پیرامونی")
               + '</p><blockquote dir="auto">' + escape(evidence) + '</blockquote>' if evidence else '')
            + ('<p class="suggestion"><strong>پیشنهاد برای بازبینی: </strong><span dir="auto">'
               + escape(suggestion) + '</span></p>' if suggestion else '')
            + '<details><summary>جزئیات ممیزی</summary><dl><dt>شناسهٔ نتیجه</dt><dd><code>'
            + escape(row["finding_id"]) + '</code></dd><dt>اعتبار مکان</dt><dd>'
            + escape(row.get("location_status", "unverified")) + '</dd><dt>اعتبار استخراج</dt><dd>'
            + escape(row.get("extraction_status", "native")) + '</dd><dt>نسخهٔ قاعده</dt><dd>'
            + escape(row.get("rule_version", "1")) + '</dd></dl></details></article>'
        )
    counts = data["summary"]["by_confidence"]
    stat_cards = ''.join('<div class="stat"><b>' + (str(counts.get(level, 0)) if counts.get(level, 0) is not None else "نامشخص") + '</b><span>' + level + ' · ' + label + '</span></div>'
                         for level, label in CONFIDENCE_LABELS.items())
    limits = ''.join('<li>' + escape(str(value)) + '</li>' for value in data["limitations"])
    empty = '<p class="empty" id="empty">در محدودهٔ قواعد اجراشده موردی یافت نشد.</p>' if not cards else '<p class="empty" id="empty" hidden>موردی با این فیلتر یافت نشد.</p>'
    source_hash = data["stats"].get("file_sha256")
    hash_row = '<p class="hash">SHA-256 ورودی: <code>' + escape(source_hash) + '</code></p>' if source_hash else ''
    summary = data["stats"].get("annotation_summary")
    annotation_section = ''
    if summary is not None or data["stats"].get("review_document") is not None or unified is not None:
        summary = summary or {}
        annotation_section = ('<section class="notes"><h2>کامنت‌ها و هایلایت‌های گزارش و نسخهٔ بازبینی</h2>'
                              + ''.join('<p>' + escape(line) + '</p>' for line in annotation_summary_lines(data["stats"]))
                              + '</section>')
    coverage = data["stats"].get("coverage", {})
    coverage_rows = ''.join('<tr><th>' + escape(STORY_LABELS.get(story, story)) + '</th><td>'
                            + escape(COVERAGE_LABELS.get(info.get("status"), "نامشخص")) + '</td><td>'
                            + str(info.get("analyzed_blocks", 0)) + ' از ' + str(info.get("present_blocks", 0)) + '</td></tr>'
                            for story, info in coverage.items() if isinstance(info, dict))
    coverage_section = ('<section class="notes"><h2>پوشش بخش‌های سند</h2><table><thead><tr><th>بخش</th><th>وضعیت</th><th>بلوک‌های متن</th></tr></thead><tbody>'
                        + coverage_rows + '</tbody></table><p class="muted">استخراج و بررسی متن، امکان درج کامنت در آن بخش را تضمین نمی‌کند.</p></section>') if coverage_rows else ''
    manual_count = counts.get("C")
    manual_note = ('<p class="muted">تعداد موارد مبهمِ کنارگذاشته‌شده در حالت سخت‌گیرانه ثبت نشده است؛ صفر فرض نشده‌اند.</p>'
                   if manual_count is None else
                   '<p class="muted">موارد سطح C فقط برای بررسی دستی‌اند. '
                   + str(data["stats"].get("ambiguous_candidates", 0)) + ' مورد مبهم در حالت سخت‌گیرانه وارد فهرست ایرادها نشده است.</p>')
    return ('''<!doctype html>
<html lang="fa" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; connect-src 'none'; base-uri 'none'; form-action 'none'">
<title>گزارش بازبینی ویراستار پژوهش</title>
<style>
:root{color-scheme:light;--ink:#182d38;--muted:#536b76;--line:#dce5e8;--accent:#0c6d69}
*{box-sizing:border-box}body{margin:0;background:#f3f6f6;color:var(--ink);font:16px/1.9 Tahoma,Arial,sans-serif}
.wrap{max-width:1080px;margin:auto;padding:40px 24px}header{border-top:6px solid var(--accent);padding:28px;background:white;border-radius:14px}
.eyebrow{color:var(--accent);font-size:13px}h1{font-size:30px;line-height:1.6;margin:8px 0}h2{font-size:22px}h3{font-size:16px;margin:14px 0 5px}
p{margin:8px 0}.muted,.location,.category{color:var(--muted);font-size:13px}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:20px 0}
.stat{background:white;border:1px solid var(--line);border-radius:12px;padding:16px;display:flex;gap:14px;align-items:center}.stat b{font-size:30px}
.toolbar{position:sticky;top:0;z-index:1;background:#f3f6f6;padding:12px 0;display:flex;gap:10px;flex-wrap:wrap}
input,select,button{font:inherit;padding:9px 12px;border:1px solid #aabcc2;border-radius:8px;background:white;color:var(--ink)}input{flex:1;min-width:190px}button{cursor:pointer}
.finding{background:white;border:1px solid var(--line);border-radius:12px;margin:16px 0;padding:24px}.finding-top{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.badge{font-size:12px;padding:3px 10px;border-radius:20px;background:#e4efed}.badge.B{background:#fff0cf}.badge.C{background:#ececf4}
code{direction:ltr;unicode-bidi:isolate;overflow-wrap:anywhere;font-size:12px}blockquote{white-space:pre-wrap;margin:6px 0;padding:14px;background:#f6f8fa;border-right:3px solid #94b4b7;overflow-wrap:anywhere}
.evidence-label{color:var(--muted);font-size:12px}.suggestion{padding:10px 14px;border:1px dashed #8fb5b2;border-radius:8px}details{font-size:13px;margin-top:12px}summary{cursor:pointer;color:var(--accent)}dl{display:grid;grid-template-columns:130px 1fr;gap:5px}dd{margin:0;overflow-wrap:anywhere}
.notes{margin-top:32px;padding:22px;border:1px solid var(--line);border-radius:12px}.hash{font-size:12px;color:var(--muted)}.empty{text-align:center;padding:24px}footer{font-size:12px;color:var(--muted);margin:20px 0}[hidden]{display:none!important}
table{width:100%;border-collapse:collapse}th,td{text-align:right;border-bottom:1px solid var(--line);padding:8px}
@media(max-width:650px){.wrap{padding:20px 14px}.stats{grid-template-columns:1fr}.finding{padding:16px}h1{font-size:24px}.toolbar{position:static}}
@media print{body{background:white}.wrap{padding:0}.toolbar,details{display:none}.finding{break-inside:avoid}.notes,header{border-color:#ccc}}
</style></head><body><main class="wrap"><header><div class="eyebrow">ویراستار پژوهش / گزارش بازبینی</div>
<h1>پیشنهادهای بازبینی، تصمیم با نویسنده</h1><p class="muted">محل هر مورد تا حد قابل تأیید مشخص شده است. هیچ پیشنهادی به‌صورت خودکار اعمال نشده است.</p>'''
    + '<p class="muted">نسخهٔ موتور: ' + escape(data["product_version"]) + ' · زمان UTC: <time dir="ltr">' + escape(data["generated_at"]) + '</time></p></header>'
    + _university_profile_html(data["stats"])
    + '<section class="stats" aria-label="تعداد نتایج">' + stat_cards + '</section>'
    + manual_note + annotation_section + coverage_section
    + '<div class="toolbar"><input id="query" type="search" placeholder="جست‌وجو در شاهد، قاعده و پیشنهاد" aria-label="جست‌وجوی نتایج">'
    + '<select id="level" aria-label="فیلتر سطح اطمینان"><option value="default">نتایج قابل بازبینی</option><option value="all">همه، همراه موارد مبهم</option><option value="A">اطمینان بالا</option><option value="B">نیازمند بازبینی</option><option value="C">مبهم</option></select>'
    + '<button type="button" id="print">چاپ / ذخیره PDF</button></div><p id="count" class="muted" aria-live="polite"></p>'
    + '<section aria-label="نتایج بازبینی">' + ''.join(cards) + empty + '</section>'
    + '<section class="notes"><h2>دامنه و محدودیت بررسی</h2><ul>' + limits + '</ul>' + hash_row + '</section>'
    + '<footer>گزارش مستقل و آفلاین است؛ هیچ داده‌ای از این صفحه به شبکه ارسال نمی‌شود.</footer></main>'
    + '''<script>
const query=document.getElementById('query'),level=document.getElementById('level'),cards=Array.from(document.querySelectorAll('.finding'));
function filter(){let visible=0;const needle=query.value.toLocaleLowerCase();for(const card of cards){const confidence=card.dataset.confidence;const selected=level.value;const matchLevel=selected==='all'||(selected==='default'?confidence!=='C':confidence===selected);card.hidden=!(matchLevel&&card.textContent.toLocaleLowerCase().includes(needle));if(!card.hidden)visible++;}document.getElementById('count').textContent=visible+' نتیجه نمایش داده می‌شود';document.getElementById('empty').hidden=visible!==0;}
query.addEventListener('input',filter);level.addEventListener('change',filter);document.getElementById('print').addEventListener('click',()=>window.print());filter();
</script></body></html>''')


def write_artifacts(output_dir, issues, stats, rules, *, source_path=None, text_report=""):
    """Write complete, portable artifacts with stable filenames in an explicit destination."""
    directory = Path(output_dir)
    directory.mkdir(parents=True, exist_ok=True)
    data = report_data(issues, stats, rules)
    payloads = {"report.json": json_report(data).encode("utf-8"),
                "report.html": html_report(data).encode("utf-8"),
                "report.csv": csv_report(data).encode("utf-8-sig")}
    if text_report:
        payloads["report.txt"] = text_report.encode("utf-8")
    # Validate every destination first, avoiding partial publication on a source collision.
    if source_path:
        from doc_guard import ensure_distinct_output
        for name in payloads:
            ensure_distinct_output(source_path, directory / name)
    for name, payload in payloads.items():
        atomic_write(directory / name, payload, source_path=source_path)
    return {name.rsplit('.', 1)[1]: directory / name for name in payloads}
