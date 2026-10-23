"""Command-line interface for reproducible document review and portable reports."""
import argparse
import json
from pathlib import Path
import sys
import time

from doc_guard import DocumentIntegrityError, ensure_distinct_output, sha256_file
from report_artifacts import PRODUCT_VERSION, csv_report, html_report, json_report, report_data, write_artifacts


def main(argv=None, *, engine=None):
    if engine is None:
        import checker as engine
    parser = argparse.ArgumentParser(description="بررسی غیرمخرب اسناد پژوهشی")
    parser.add_argument("file", nargs="?")
    parser.add_argument("--mode", choices=("thesis", "tums", "article", "english", "kashan", "kashan_master", "kashan_phd"), default="thesis")
    parser.add_argument("--scope", choices=("full", "fragment"), help="سند کامل یا بخشی از سند؛ پیش‌فرض کاشان بخشی از سند است")
    parser.add_argument("--rules", help="فایل JSON قواعد سفارشی")
    parser.add_argument("--text", action="store_true", help="خواندن ورودی به‌عنوان متن UTF-8")
    parser.add_argument("--format", choices=("text", "json", "html", "csv"), default="text")
    parser.add_argument("--output-dir", help="ذخیرهٔ گزارش‌های HTML، JSON، CSV و متن")
    parser.add_argument("--annotate", help="مسیر نسخهٔ بازبینی Word؛ همهٔ یافته‌ها در همین فایل با زرد و کامنت")
    parser.add_argument("--word-report", help="مسیر گزارش مستقل Word؛ همهٔ یافته‌ها با عنوان زرد و کامنت جداگانه")
    parser.add_argument("--annotate-pdf", help="آزمایشی: مسیر کپی PDF با یادداشت در محل‌های قابل تطبیق")
    parser.add_argument("--max-comments", type=int, default=None,
                        help="سقف اختیاری کامنت دقیق/پاراگرافی (۰ تا ۱۰۰۰۰)؛ باقی موارد با ارجاع صریح در متن موجود همان Word کامنت می‌گیرند")
    parser.add_argument("--include-ambiguous", action="store_true")
    parser.add_argument("--fail-on", choices=("A", "B", "C"), help="بازگرداندن کد ۱ در صورت یافتن این سطح")
    parser.add_argument("--list-rules", action="store_true")
    parser.add_argument("--version", action="version", version=PRODUCT_VERSION)
    args = parser.parse_args(argv)
    if args.max_comments is not None and not 0 <= args.max_comments <= 10000:
        parser.error("--max-comments باید بین ۰ و ۱۰۰۰۰ باشد")
    if args.annotate and args.annotate_pdf:
        parser.error("فقط یکی از --annotate و --annotate-pdf را انتخاب کنید")
    if args.list_rules:
        from rule_registry import rule_catalog
        extras = [row[0] for row in engine.TEXT_RULES]
        print(json.dumps(rule_catalog(extras), ensure_ascii=False, indent=2))
        return 0
    if not args.file:
        parser.error("مسیر فایل ورودی لازم است")
    try:
        source = Path(args.file)
        outputs = [Path(path) for path in (args.annotate, args.annotate_pdf, args.word_report) if path]
        if args.output_dir:
            outputs.extend(Path(args.output_dir) / f"report.{suffix}" for suffix in ("html", "json", "csv", "txt"))
        for index, destination in enumerate(outputs):
            ensure_distinct_output(source, destination)
            for other in outputs[:index]:
                # The alias helper expects an existing source when the target
                # exists. A previous report may already exist while a new
                # annotation path does not, so prefer the existing path first.
                pair = (destination, other) if destination.exists() else (other, destination)
                ensure_distinct_output(*pair)
        rules = engine.load_rules(args.rules, args.mode)
        rules["document_scope"] = args.scope or ("fragment" if args.mode.startswith("kashan") else "full")
        if args.include_ambiguous:
            rules["strict"] = False
        before = sha256_file(source)
        start = time.monotonic()
        if args.text:
            if source.stat().st_size > rules.get("max_input_bytes", 25 * 1024 * 1024):
                raise ValueError("Text input exceeds its size budget")
            doc, issues, stats = engine.analyze_text(source.read_text(encoding="utf-8-sig"), args.mode, rules)
            stats["file_sha256"] = before
        else:
            with source.open("rb") as stream:
                signature = stream.read(8)
            if signature.startswith(b"%PDF-"):
                doc, issues, stats = engine.analyze_pdf(source, args.mode, rules)
            else:
                doc, issues, stats = engine.analyze(source, args.mode, rules)
        if sha256_file(source) != before:
            raise DocumentIntegrityError("Input changed during analysis")
        stats["analysis_duration_seconds"] = round(time.monotonic() - start, 3)
        if args.annotate and doc is None:
            raise ValueError("Annotation is available only for Word documents")
        if args.annotate:
            engine.annotate(doc, issues, args.annotate, max_comments=args.max_comments, per_par=None)
            stats["annotation_summary"] = getattr(doc, "_last_annotation_summary", {})
            stats["annotation_format"] = "docx"
            from unified_word_review import finalize_word_review
            stats["unified_word_review"] = finalize_word_review(
                source, args.annotate, issues, stats, rules)
        if args.annotate_pdf:
            if not stats.get("is_pdf"):
                raise ValueError("--annotate-pdf requires a PDF input")
            from pdf_annotator import annotate_pdf
            pdf_limit = min(5000, len(issues) if args.max_comments is None else args.max_comments)
            stats["annotation_summary"] = annotate_pdf(source, issues, args.annotate_pdf, max_comments=pdf_limit)
            stats["annotation_format"] = "pdf"
        if args.word_report:
            from word_review_report import write_word_review
            review = write_word_review(args.word_report, issues, stats, rules, source_path=source)
            stats["review_document"] = {
                "created": review["output_created"], "findings": review["requested"],
                "comments": review["comments"], "highlighted": review["highlighted_entries"], "filetype": "docx",
            }
        text = engine.build_report(issues, stats, rules)
        if args.output_dir:
            write_artifacts(args.output_dir, issues, stats, rules, source_path=source, text_report=text)
        data = report_data(issues, stats, rules)
        output = {"text": lambda: text, "json": lambda: json_report(data),
                  "html": lambda: html_report(data), "csv": lambda: csv_report(data)}[args.format]()
        if sha256_file(source) != before:
            raise DocumentIntegrityError("Input changed during report generation")
        print(output)
        if stats.get("analysis_status") == "failed" or any(i.rule_id == "pdf_read_error" for i in issues):
            return 2
        return int(bool(args.fail_on and any(i.confidence == args.fail_on for i in issues)))
    except (OSError, ValueError, DocumentIntegrityError, json.JSONDecodeError) as error:
        print(f"Review failed: {error}", file=sys.stderr)
        return 2
    except Exception as error:
        print(f"Review failed ({type(error).__name__}): the document could not be read or processed safely.", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
