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
PRODUCT_VERSION = "2.1.0"
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
    safe = {key: max(0, value) for key in ("requested", "applied", "unapplied", "max_comments")
            if isinstance((value := summary.get(key)), int) and not isinstance(value, bool)}
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
        safe["records"].append({key: value for key, value in record.items()
                                if (key in numeric and isinstance(value, int) and not isinstance(value, bool))
                                or (key in textual and isinstance(value, str))})
    return safe


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


def report_data(issues, stats, rules):
    """Export findings and provenance without local filesystem paths or configuration text."""
    encoded_rules = json.dumps(rules, sort_keys=True, ensure_ascii=False).encode("utf-8")
    rows = []
    annotation = safe_annotation_summary(stats.get("annotation_summary"))
    records = {record.get("index"): record for record in (annotation or {}).get("records", [])}
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
            row["annotation"] = {key: record[key] for key in ("status", "reason") if key in record}
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
    for row in data["findings"]:
        confidence = row["confidence"]
        original = row.get("original_text")
        evidence = original if original is not None else row.get("snippet", "")
        suggestion = row.get("suggestion")
        annotation = row.get("annotation", {})
        annotation_text = (f"در {output_label} کامنت شد" if annotation.get("status") == "applied" else
                           ANNOTATION_REASONS.get(annotation.get("reason"), f"در {output_label} کامنت نشد")) if annotation else ""
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
    if summary:
        reasons = ''.join('<li>' + escape(ANNOTATION_REASONS.get(reason, reason)) + ': ' + str(count) + '</li>'
                          for reason, count in summary.get("reasons", {}).items())
        annotation_section = ('<section class="notes"><h2>کامنت‌های نسخهٔ بازبینی</h2><p>از '
                              + str(summary.get("requested", len(cards))) + ' مورد، '
                              + str(summary.get("applied", 0)) + ' مورد در ' + output_label + ' کامنت شد.</p>'
                              + ('<p>تولید نسخهٔ بازبینی تکمیل نشد؛ جزئیات در گزارش باقی است.</p>'
                                 if not summary.get("output_created") else '')
                              + ('<ul>' + reasons + '</ul>' if reasons else '') + '</section>')
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
