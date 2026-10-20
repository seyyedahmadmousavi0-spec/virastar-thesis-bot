"""Bounded, page-aware PDF extraction with source text preservation."""
import unicodedata
try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        pymupdf = None

def normalize_presentation_forms_only(text):
    """Comparison-only helper; extraction preserves original Unicode."""
    return "".join(unicodedata.normalize("NFKC", c) if 0xFB50 <= ord(c) <= 0xFDFF or 0xFE70 <= ord(c) <= 0xFEFC else c for c in text)

def _failure(code):
    messages = {
        "dependency": "ابزار استخراج PDF روی میزبان در دسترس نیست؛ بررسی انجام نشد.",
        "encrypted": "PDF رمزگذاری شده است؛ یک نسخهٔ قابل‌خواندن و بدون رمز ارسال کنید.",
        "budget": "تعداد صفحات یا حجم متن PDF از ظرفیت بررسی بیشتر است؛ سند را به بخش‌های کوچک‌تر تقسیم کنید.",
        "read_error": "استخراج قابل‌اعتماد PDF موفق نبود؛ فایل و متن آن نیازمند بررسی دستی است.",
    }
    return {"ok": False, "error_code": code, "error": messages[code], "is_scanned": False,
            "pages": [], "total_pages": 0, "total_chars": 0}

def extract_pdf_content(pdf_path, *, max_pages=1000, max_chars=4_000_000, max_blocks=25000):
    if pymupdf is None:
        return _failure("dependency")
    if any(not isinstance(value, int) or isinstance(value, bool) or value <= 0 for value in (max_pages, max_chars, max_blocks)):
        raise ValueError("Invalid PDF extraction budget")
    try:
        with pymupdf.open(pdf_path) as document:
            if document.needs_pass:
                return _failure("encrypted")
            if len(document) > max_pages:
                return _failure("budget")
            pages, total_chars, total_blocks = [], 0, 0
            for page_index, page in enumerate(document):
                raw_text = page.get_text("text", sort=False) or ""
                total_chars += len(raw_text)
                if total_chars > max_chars:
                    return _failure("budget")
                image_count = len(page.get_images(full=True))
                chars = len(raw_text.strip())
                status = "text" if chars >= 50 else "short_text" if chars else "image_only" if image_count else "no_text"
                warnings = []
                if "\ufffd" in raw_text:
                    warnings.append("replacement_characters")
                if any(0xFB50 <= ord(c) <= 0xFDFF or 0xFE70 <= ord(c) <= 0xFEFC for c in raw_text):
                    warnings.append("presentation_forms")
                if status in {"image_only", "no_text"}:
                    warnings.append("no_reliable_text")
                blocks = []
                for item in page.get_text("blocks", sort=False):
                    if len(item) >= 7 and item[6] == 0 and item[4].strip():
                        total_blocks += 1
                        if total_blocks > max_blocks:
                            return _failure("budget")
                        blocks.append({"block_id": f"pdf:p{page_index+1}:b{len(blocks)+1}", "text": item[4],
                                       "bbox": list(item[:4]), "page_index": page_index})
                if not blocks and raw_text.strip():
                    total_blocks += 1
                    if total_blocks > max_blocks:
                        return _failure("budget")
                    blocks = [{"block_id": f"pdf:p{page_index+1}:b1", "text": raw_text, "bbox": None, "page_index": page_index}]
                for left_index, left in enumerate(blocks):
                    if not left["bbox"]:
                        continue
                    for right in blocks[left_index+1:]:
                        if not right["bbox"]:
                            continue
                        a, b = left["bbox"], right["bbox"]
                        if min(a[3], b[3]) > max(a[1], b[1]) and (a[2] + 12 < b[0] or b[2] + 12 < a[0]):
                            warnings.append("possible_multiple_columns")
                            break
                    if "possible_multiple_columns" in warnings:
                        break
                pages.append({"page_num": page_index+1, "text": raw_text,
                              "paragraphs": [block["text"] for block in blocks], "blocks": blocks,
                              "quality": status, "quality_warnings": warnings,
                              "image_count": image_count, "text_chars": chars})
            no_text = not any(page["text_chars"] for page in pages)
            return {"ok": True, "total_pages": len(document), "total_chars": total_chars,
                    "no_text_document": no_text,
                    "is_scanned": no_text and any(page["image_count"] for page in pages), "pages": pages}
    except Exception:
        return _failure("read_error")
