"""Portable, offline review artifacts. Source content is always treated as data."""
from collections import Counter
import csv
from datetime import datetime, timezone
import hashlib
from html import escape
import io
import json
from pathlib import Path

from doc_guard import atomic_write

SCHEMA_VERSION = "2.0"
PRODUCT_VERSION = "2.0.1"
CONFIDENCE_LABELS = {"A": "اطمینان بالا", "B": "نیازمند بازبینی", "C": "مبهم"}


def report_data(issues, stats, rules):
    """Export findings and provenance without local filesystem paths or configuration text."""
    encoded_rules = json.dumps(rules, sort_keys=True, ensure_ascii=False).encode("utf-8")
    rows = []
    for issue in issues:
        row = issue.to_dict()
        identity = [stats.get("file_sha256") or stats.get("analysis_text_sha256")]
        identity += [row.get(key) for key in ("rule_id", "block_id", "source_part", "xml_path", "par_index", "char_start", "char_end", "message")]
        row["finding_id"] = hashlib.sha256(json.dumps(identity, ensure_ascii=False).encode("utf-8")).hexdigest()[:20]
        rows.append(row)
    # Only known analysis statistics are exported; arbitrary caller metadata stays private.
    keys = {"mode", "words", "body_words", "paragraphs", "tables", "images", "chapters", "total_pages",
            "table_cells_analyzed", "coverage_limits", "file_sha256", "is_pdf", "is_scanned",
            "is_raw_text", "block_count", "story_counts", "partial_blocks", "protected_content_audit",
            "extraction_pages", "analysis_duration_seconds", "skipped_blocks", "source_chars", "analysis_text_sha256",
            "analysis_status", "analysis_error_code"}
    safe_stats = {key: value for key, value in stats.items() if key in keys}
    return {
        "schema_version": SCHEMA_VERSION,
        "product_version": PRODUCT_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "configuration_sha256": hashlib.sha256(encoded_rules).hexdigest(),
        "strict": bool(rules.get("strict", True)),
        "summary": {"findings": len(rows), "by_confidence": dict(Counter(i.confidence for i in issues)),
                    "by_category": dict(Counter(i.category for i in issues))},
        "stats": safe_stats,
        "limitations": ["این بررسی قاعده‌محور است و صحت محتوای علمی را تأیید نمی‌کند.",
                        "سطح اطمینان احتمال آماری یا درصد دقت نیست.",
                        *(["بررسی سند تکمیل نشد؛ متن قابل‌اعتماد برای ارزیابی در دسترس نبود."]
                          if stats.get("analysis_status") == "failed" else []),
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
    for row in data["findings"]:
        confidence = row["confidence"]
        original = row.get("original_text")
        evidence = original if original is not None else row.get("snippet", "")
        suggestion = row.get("suggestion")
        cards.append(
            '<article class="finding" data-confidence="' + escape(confidence, quote=True) + '">'
            '<div class="finding-top"><span class="badge ' + escape(confidence, quote=True) + '">'
            + escape(CONFIDENCE_LABELS.get(confidence, confidence)) + '</span><code>'
            + escape(row["rule_id"]) + '</code><span class="category">' + escape(row.get("category", "")) + '</span></div>'
            '<h3>' + escape(row.get("message", "")) + '</h3>'
            '<p class="location">' + escape(row.get("loc") or "بررسی در سطح سند") + '</p>'
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
    stat_cards = ''.join('<div class="stat"><b>' + str(counts.get(level, 0)) + '</b><span>' + label + '</span></div>'
                         for level, label in CONFIDENCE_LABELS.items())
    limits = ''.join('<li>' + escape(str(value)) + '</li>' for value in data["limitations"])
    empty = '<p class="empty" id="empty">در محدودهٔ قواعد اجراشده موردی یافت نشد.</p>' if not cards else '<p class="empty" id="empty" hidden>موردی با این فیلتر یافت نشد.</p>'
    source_hash = data["stats"].get("file_sha256")
    hash_row = '<p class="hash">SHA-256 ورودی: <code>' + escape(source_hash) + '</code></p>' if source_hash else ''
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
@media(max-width:650px){.wrap{padding:20px 14px}.stats{grid-template-columns:1fr}.finding{padding:16px}h1{font-size:24px}.toolbar{position:static}}
@media print{body{background:white}.wrap{padding:0}.toolbar,details{display:none}.finding{break-inside:avoid}.notes,header{border-color:#ccc}}
</style></head><body><main class="wrap"><header><div class="eyebrow">ویراستار پژوهش / گزارش بازبینی</div>
<h1>بررسی دقیق، تصمیم با نویسنده</h1><p class="muted">نتایج به متن اصلی متصل‌اند. هیچ پیشنهادی به‌صورت خودکار اعمال نشده است.</p>'''
    + '<p class="muted">نسخهٔ موتور: ' + escape(data["product_version"]) + ' · زمان UTC: <time dir="ltr">' + escape(data["generated_at"]) + '</time></p></header>'
    + '<section class="stats" aria-label="تعداد نتایج">' + stat_cards + '</section>'
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
