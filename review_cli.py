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
    parser.add_argument("--mode", choices=("thesis", "tums", "article", "english"), default="thesis")
    parser.add_argument("--rules", help="فایل JSON قواعد سفارشی")
    parser.add_argument("--text", action="store_true", help="خواندن ورودی به‌عنوان متن UTF-8")
    parser.add_argument("--format", choices=("text", "json", "html", "csv"), default="text")
    parser.add_argument("--output-dir", help="ذخیرهٔ گزارش‌های HTML، JSON، CSV و متن")
    parser.add_argument("--annotate", help="مسیر کپی مستقل Word با کامنت")
    parser.add_argument("--include-ambiguous", action="store_true")
    parser.add_argument("--fail-on", choices=("A", "B", "C"), help="بازگرداندن کد ۱ در صورت یافتن این سطح")
    parser.add_argument("--list-rules", action="store_true")
    parser.add_argument("--version", action="version", version=PRODUCT_VERSION)
    args = parser.parse_args(argv)
    if args.list_rules:
        from rule_registry import rule_catalog
        extras = [row[0] for row in engine.TEXT_RULES]
        print(json.dumps(rule_catalog(extras), ensure_ascii=False, indent=2))
        return 0
    if not args.file:
        parser.error("مسیر فایل ورودی لازم است")
    try:
        source = Path(args.file)
        if args.annotate:
            ensure_distinct_output(source, args.annotate)
        rules = engine.load_rules(args.rules, args.mode)
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
        text = engine.build_report(issues, stats, rules)
        if args.output_dir:
            write_artifacts(args.output_dir, issues, stats, rules, source_path=source, text_report=text)
        if args.annotate:
            engine.annotate(doc, issues, args.annotate)
        data = report_data(issues, stats, rules)
        output = {"text": lambda: text, "json": lambda: json_report(data),
                  "html": lambda: html_report(data), "csv": lambda: csv_report(data)}[args.format]()
        print(output)
        if any(i.rule_id == "pdf_read_error" for i in issues):
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
