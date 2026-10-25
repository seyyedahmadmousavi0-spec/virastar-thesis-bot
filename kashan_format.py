# -*- coding: utf-8 -*-
"""Word-format rules of the University of Kashan graduate office.

Source: the three-page «نکات مربوط به تهیه، تدوین و نگارش پایان‌نامه» of the
Kashan graduate office and the official master's/PhD templates (Bahman 1398),
supplied by the bot owner on 2026-10-25. Each finding names its clause.

Only native Word evidence is used: effective run fonts, sizes and toggles,
section properties and existing header/footer parts (read through relations,
never through python-docx accessors that would create parts). Every finding is
a confidence-B review item; nothing is changed in the document.
"""
import re

from docx.oxml.ns import qn

from analysis_model import Issue

GUIDE = "راهنمای سه‌صفحه‌ای تحصیلات تکمیلی دانشگاه کاشان (نکات تهیه، تدوین و نگارش پایان‌نامه)، بند {}"
TEMPLATE = "الگوی رسمی {} دانشگاه کاشان، بهمن ۱۳۹۸"

BODY_FONTS = ("nazanin", "lotus", "نازنین", "لوتوس")
TITR_FONTS = ("titr", "تیتر")
BAD_COVER_FONTS = ("nastaliq", "nastaleeq", "shekaste", "shekasteh", "نستعلیق", "شکسته", "kufi")
BODY_SIZE, FOOTNOTE_SIZE = 12, 10
HEADING_SIZES = {2: 16, 3: 14}          # depth of «۱-۱» / «۱-۱-۱»; deeper headings 12
MARGIN_CM, MARGIN_TOLERANCE = 3.0, 0.25
LINES_PER_PAGE = (24, 26)
KEYWORD_COUNT = (4, 5)
ABJAD_FORMATS = {"arabicAbjad", "arabicAlpha"}
FA_MONTHS = ("فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند")
EN_MONTHS = ("january", "february", "march", "april", "may", "june", "july", "august",
             "september", "october", "november", "december")
_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
_NUMBERED_HEADING = re.compile(r"^\s*[-–]?\s*([0-9۰-۹]+(?:\s*[-–.]\s*[0-9۰-۹]+)+)\s*[-–.]?\s+\S")
_FA_LETTER = re.compile("[؀-ۿ]")


def _issue(rule_id, message, source, record=None, span=None, suggestion="", loc=None, par_index=-1):
    issue = Issue(rule_id, "warn", message + "\nمستند: " + source,
                  loc=loc or (record["location"] if record else "ساختار سند"),
                  confidence="B", category="university_profile", suggestion=suggestion)
    if record:
        for key in ("par_index", "block_id", "source_part", "xml_path", "story", "page_index", "extraction_status"):
            setattr(issue, key, record[key])
        issue.paragraph_index = record["par_index"]
        if span:
            issue.char_start, issue.char_end = span
            issue.target_text = record["text"][span[0]:span[1]]
            issue.validate_evidence(record["text"])
        else:
            issue.snippet = record["text"][:160]
    elif par_index >= 0:
        issue.par_index = issue.paragraph_index = par_index
    return issue


# ----------------------------------------------------------------------------
# effective run properties
# ----------------------------------------------------------------------------
def _levels(doc, paragraph, run):
    import checker
    return checker._levels(doc, paragraph, "r", run)


def _toggle(doc, paragraph, run, *tags):
    """Effective bold/italic; Persian text uses the complex-script toggle."""
    for level in _levels(doc, paragraph, run):
        if level is None:
            continue
        for tag in tags:
            node = level.find(qn(tag))
            if node is not None:
                return node.get(qn("w:val")) not in ("0", "false", "off")
    return False


def _font(doc, paragraph, run, persian=True):
    import checker
    return checker.eff_rpr(doc, paragraph, run, "w:rFonts", "w:cs" if persian else "w:ascii")


def _size(doc, paragraph, run, persian=True):
    import checker
    value = checker.eff_rpr(doc, paragraph, run, "w:szCs" if persian else "w:sz", "w:val")
    return int(value) / 2 if value and value.isdigit() else None


def _text_runs(paragraph):
    return [run for run in paragraph.runs if run.text.strip()]


def _matches(name, families):
    return bool(name) and any(family in name.lower().replace(" ", "") for family in families)


def _all_bold(doc, paragraph):
    runs = _text_runs(paragraph)
    return bool(runs) and all(_toggle(doc, paragraph, run, "w:b", "w:bCs") for run in runs)


# ----------------------------------------------------------------------------
# covers (clauses 2-7 and 22-25)
# ----------------------------------------------------------------------------
def _persian_cover(doc, window, degree, out):
    window = [record for record in window if record["par_index"] >= 0 and record["story"] == "body"]
    if not window:
        return
    source = lambda n: GUIDE.format(n)
    paragraphs = doc.paragraphs
    first = window[0]["par_index"]
    # 2) the university logo above the cover lines
    start = max(0, first - 6)
    if not any(p._p.xpath(".//w:drawing|.//w:pict") for p in paragraphs[start:first + 1]):
        out.append(_issue("KASHAN_COVER_LOGO", "در بالای صفحهٔ عنوان فارسی تصویر آرم دانشگاه پیدا نشد.",
                          source("۲"), window[0], suggestion="آرم دانشگاه را به‌صورت شکیل بالای صفحهٔ عنوان درج کنید."))
    # 3) one font for university/faculty/group lines, never Nastaliq or Shekasteh
    head_fonts = set()
    for record in window:
        paragraph = paragraphs[record["par_index"]]
        for run in _text_runs(paragraph):
            name = _font(doc, paragraph, run)
            if _matches(name, BAD_COVER_FONTS):
                out.append(_issue("KASHAN_COVER_FONT", f"در صفحهٔ عنوان فارسی قلم «{name}» به کار رفته است؛ خط شکسته و نستعلیق مجاز نیست.",
                                  source("۳"), record, suggestion="قلم این سطر را با قلم سطرهای دیگر صفحهٔ عنوان یکسان کنید."))
                break
        text = record["text"].strip()
        if re.match(r"(?:دانشکده|گروه|پژوهشکده)", text):
            runs = _text_runs(paragraph)
            if runs:
                head_fonts.add((_font(doc, paragraph, runs[0]), _size(doc, paragraph, runs[0])))
    if len(head_fonts) > 1:
        out.append(_issue("KASHAN_COVER_FONT", "قلم یا اندازهٔ سطرهای دانشکده، پژوهشکده و گروه در صفحهٔ عنوان یکسان نیست.",
                          source("۳"), window[0], suggestion="قلم و فاصلهٔ این سطرها را یکسان کنید."))
    word = "رساله" if degree == "phd" else "پایان"
    for record in window:
        paragraph = paragraphs[record["par_index"]]
        text = record["text"]
        # 4) «پایان‌نامه» (PhD: «رساله») on its own line, bold
        if re.match(r"\s*(?:پایان[\s‌]*نامه|رساله)\b", text) and len(text.split()) <= 10:
            if not _all_bold(doc, paragraph):
                out.append(_issue("KASHAN_DEGREE_BOLD", f"سطر «{text.strip()[:40]}» در صفحهٔ عنوان فارسی باید Bold باشد.",
                                  source("۴"), record, suggestion="کل این سطر را Bold کنید."))
            if degree == "phd" and not text.strip().startswith("رساله"):
                out.append(_issue("KASHAN_DEGREE_WORD", "در رساله دکتری به جای «پایان‌نامه» از واژهٔ «رساله» استفاده شود.",
                                  source("۴"), record, suggestion="رساله"))
        # 6) «توسط» instead of «به وسیله»
        match = re.search(r"به[\s‌]*وسیله|بوسیله", text)
        if match:
            out.append(_issue("KASHAN_COVER_BY", "در صفحهٔ عنوان به جای «به وسیله» از «توسط» استفاده شود.",
                              source("۶"), record, match.span(), "توسط"))
    # 7) defence date as Solar month and year
    joined = " ".join(record["text"] for record in window).translate(_DIGITS)
    # Whole words only: «دی»، «تیر» and «مهر» also occur inside other words.
    month = re.search(r"(?<![؀-ۿ])(?:" + "|".join(FA_MONTHS) + r")(?![؀-ۿ])", joined)
    if not (month and re.search(r"(?<![0-9])1[34][0-9]{2}(?![0-9])", joined)):
        out.append(_issue("KASHAN_COVER_DATE", "تاریخ دفاع در صفحهٔ عنوان فارسی به‌صورت ماه و سال شمسی (مثلاً «شهریور ۱۴۰۳») پیدا نشد.",
                          source("۷"), window[-1], suggestion="ماه و سال شمسی دفاع را در پایین صفحهٔ عنوان بنویسید."))


def _english_cover(doc, window, degree, out):
    if not window:
        return
    paragraphs = doc.paragraphs
    for record in window:
        if record["par_index"] < 0:
            continue
        paragraph = paragraphs[record["par_index"]]
        text = record["text"]
        # 22) no italics on the English title page
        if any(_toggle(doc, paragraph, run, "w:i", "w:iCs") for run in _text_runs(paragraph)):
            out.append(_issue("KASHAN_EN_COVER_ITALIC", "در صفحهٔ عنوان انگلیسی از حروف ایتالیک استفاده شده است.",
                              GUIDE.format("۲۲"), record, suggestion="ایتالیک این سطر را بردارید."))
        # 23) «The» is removed from the university name
        match = re.search(r"\bthe\s+university\s+of\s+kashan\b", text, re.I)
        if match:
            out.append(_issue("KASHAN_EN_THE", "واژهٔ «The» از ابتدای نام دانشگاه حذف شود.",
                              GUIDE.format("۲۳"), record, match.span(), "University of Kashan"))
        # 24) «Thesis» / «Dissertation» on its own line, bold
        if re.search(r"\b(?:thesis|dissertation)\b", text, re.I) and len(text.split()) <= 8:
            if not _all_bold(doc, paragraph):
                out.append(_issue("KASHAN_DEGREE_BOLD", f"سطر «{text.strip()[:40]}» در صفحهٔ عنوان انگلیسی باید Bold باشد.",
                                  GUIDE.format("۲۴"), record, suggestion="کل این سطر را Bold کنید."))
    # 25) defence date as Gregorian month and year
    joined = " ".join(record["text"] for record in window).lower()
    if not (any(month in joined for month in EN_MONTHS) and re.search(r"\b(?:19|20)[0-9]{2}\b", joined)):
        out.append(_issue("KASHAN_COVER_DATE", "تاریخ دفاع در صفحهٔ عنوان انگلیسی به‌صورت ماه و سال میلادی (مثلاً «September 2024») پیدا نشد.",
                          GUIDE.format("۲۵"), window[-1], suggestion="ماه و سال میلادی دفاع را بنویسید."))


# ----------------------------------------------------------------------------
# opening pages (clauses 11, 12, 15-17, 21)
# ----------------------------------------------------------------------------
def _section_after(records, heading):
    """Records after a heading up to the next isolated heading."""
    import kashan_checker
    result = []
    for record in records[heading["index"] + 1:]:
        if kashan_checker._heading(record) and len(record["text"].split()) <= 4:
            break
        result.append(record)
    return result


def _keywords(doc, record, english, out):
    paragraph = doc.paragraphs[record["par_index"]]
    text = record["text"]
    label = re.match(r"\s*(key\s*words?|کلمات کلیدی|واژگان کلیدی|واژه[\s‌]*های کلیدی|کلیدواژه[\s‌]*ها)\s*[:：]?", text, re.I)
    if not label:
        return
    clause = "۲۱" if english else "۱۲"
    runs = [run for run in paragraph.runs if run.text.strip()]
    if runs and not _toggle(doc, paragraph, runs[0], "w:b", "w:bCs"):
        out.append(_issue("KASHAN_KEYWORDS", f"عبارت «{label.group(1)}» در پایان چکیده باید Bold باشد.",
                          GUIDE.format(clause), record, label.span(1), suggestion="این عبارت را Bold کنید."))
    items = [item for item in re.split(r"[،,؛;]", text[label.end():]) if item.strip()]
    if items and not KEYWORD_COUNT[0] <= len(items) <= KEYWORD_COUNT[1]:
        out.append(_issue("KASHAN_KEYWORDS", f"{len(items)} {'کلیدواژهٔ انگلیسی' if english else 'کلیدواژه'} نوشته شده است؛ ۴ یا ۵ کلیدواژه لازم است.",
                          GUIDE.format(clause), record, suggestion="۴ یا ۵ کلیدواژه بنویسید."))


def _opening_pages(doc, records, out):
    import kashan_checker
    find = kashan_checker._find_heading
    ack = find(records, ("سپاسگزاری", "سپاس گزاری", "تقدیر و تشکر", "تشکر و قدردانی", "سپاس و قدردانی"))
    if ack is not None:
        body = [record for record in _section_after(records, ack) if record["par_index"] >= 0]
        paragraphs = doc.paragraphs
        if body and not any(_toggle(doc, paragraphs[r["par_index"]], run, "w:b", "w:bCs")
                            for r in body for run in _text_runs(paragraphs[r["par_index"]])):
            out.append(_issue("KASHAN_ACK_NAMES", "در صفحهٔ سپاسگزاری نام استادان و داوران جلسهٔ دفاع باید Bold باشد؛ هیچ نام Boldی پیدا نشد.",
                              GUIDE.format("۱۱"), body[0], suggestion="نام استادان راهنما، مشاور و داوران را Bold کنید."))
    for record in records:
        if record["par_index"] < 0:
            continue
        text = record["text"].strip()
        if re.match(r"\s*(?:کلمات کلیدی|واژگان کلیدی|واژه[\s‌]*های کلیدی|کلیدواژه)", text):
            _keywords(doc, record, False, out)
        elif re.match(r"\s*key\s*words?\b", text, re.I):
            _keywords(doc, record, True, out)
        # 17) «جدول‌ها» and «شکل‌ها», not «جداول» and «اشکال»
        if kashan_checker._heading(record):
            match = re.search(r"فهرست\s+(جداول|اشکال)", text)
            if match:
                right = "جدول‌ها" if match.group(1) == "جداول" else "شکل‌ها"
                start = record["text"].index(match.group(1))
                out.append(_issue("KASHAN_LIST_TITLE", f"در عنوان فهرست به جای «{match.group(1)}» واژهٔ «{right}» به کار رود.",
                                  GUIDE.format("۱۷"), record, (start, start + len(match.group(1))), right))
    # 15-16) chapter titles bold in the table of contents; references listed
    toc = [r for r in records if r["par_index"] >= 0 and r["style"].lower().startswith("toc") and r["style"].lower() != "toc heading"]
    if toc:
        paragraphs = doc.paragraphs
        chapters = [r for r in toc if re.match(r"\s*فصل\s", r["text"])]
        plain = [r for r in chapters if not _all_bold(doc, paragraphs[r["par_index"]])]
        if plain:
            out.append(_issue("KASHAN_TOC", f"{len(plain)} عنوان فصل در فهرست مطالب Bold نیست.",
                              GUIDE.format("۱۵"), plain[0], suggestion="عنوان فصل‌ها را در فهرست مطالب Bold کنید (در Word: سبک TOC 1)."))
        if not any(re.search(r"منابع|مراجع|مآخذ", r["text"]) for r in toc):
            out.append(_issue("KASHAN_TOC", "در انتهای فهرست مطالب، فهرست منابع و مراجع با شمارهٔ صفحه نیامده است.",
                              GUIDE.format("۱۶"), toc[-1], suggestion="عنوان «منابع و مراجع» را با سبک Heading بنویسید و فهرست را به‌روز کنید."))


# ----------------------------------------------------------------------------
# body text and headings (clauses 26-27)
# ----------------------------------------------------------------------------
def _body(doc, records, out):
    import checker
    import kashan_checker
    find = kashan_checker._find_heading
    start = next((r for r in records if kashan_checker._heading(r) and kashan_checker._CHAPTER.fullmatch(kashan_checker._norm(r["text"]))), None)
    if start is None:
        return
    end = find(records[start["index"]:], ("منابع", "مراجع", "منابع و مراجع", "منابع و مآخذ", "references", "bibliography"))
    stop = end["index"] if end else len(records)
    paragraphs = doc.paragraphs
    findings = []
    for record in records[start["index"]:stop]:
        if record["par_index"] < 0 or record["story"] != "body":
            continue
        paragraph = paragraphs[record["par_index"]]
        text = record["text"].strip()
        runs = [run for run in _text_runs(paragraph) if _FA_LETTER.search(run.text)]
        if not runs:
            continue
        numbered = _NUMBERED_HEADING.match(text)
        if numbered and len(text.split()) <= 15:
            depth = len(re.findall(r"[0-9۰-۹]+", numbered.group(1)))
            want = HEADING_SIZES.get(depth, 12)
            names = {_font(doc, paragraph, run) for run in runs}
            sizes = {_size(doc, paragraph, run) for run in runs} - {None}
            level = {2: "عنوان اصلی", 3: "عنوان فرعی"}.get(depth, "عنوان فرعی‌تر")
            if not all(_matches(name, TITR_FONTS) for name in names if name):
                findings.append(_issue("KASHAN_HEADING_FONT", f"{level} باید با قلم تیتر نوشته شود.", GUIDE.format("۲۷"),
                                       record, suggestion=f"قلم تیتر با اندازهٔ {want}"))
            if sizes and sizes != {want}:
                findings.append(_issue("KASHAN_HEADING_FONT", f"اندازهٔ {level} {', '.join(f'{s:g}' for s in sorted(sizes))} است؛ باید {want} باشد.",
                                       GUIDE.format("۲۷"), record, suggestion=f"اندازه {want}"))
            continue
        if checker.is_heading(paragraph) or len(text) < 40 or re.match(r"\s*(?:جدول|شکل|نمودار)\s", text):
            continue
        names = {_font(doc, paragraph, run) for run in runs} - {None}
        sizes = {_size(doc, paragraph, run) for run in runs} - {None}
        bad = sorted(name for name in names if not _matches(name, BODY_FONTS))
        if bad:
            findings.append(_issue("KASHAN_BODY_FONT", f"قلم متن «{'، '.join(bad)}» است؛ متن اصلی با قلم نازنین یا لوتوس نوشته شود.",
                                   GUIDE.format("۲۶"), record, suggestion="نازنین یا لوتوس با اندازهٔ ۱۲"))
        if sizes and sizes != {BODY_SIZE}:
            findings.append(_issue("KASHAN_BODY_FONT", f"اندازهٔ قلم متن {', '.join(f'{s:g}' for s in sorted(sizes))} است؛ باید ۱۲ باشد.",
                                   GUIDE.format("۲۶"), record, suggestion="اندازه ۱۲"))
    out.extend(checker.group_repeated_format(findings, 3))


def _footnotes(doc, out):
    """Footnotes in Nazanin or Lotus, size 10 (clause 26)."""
    part = next((rel.target_part for rel in doc.part.rels.values()
                 if not rel.is_external and rel.reltype.endswith("/footnotes")), None)
    if part is None:
        return
    root = part.element if hasattr(part, "element") else None
    if root is None:
        from lxml import etree
        root = etree.fromstring(part.blob)
    styles = {}
    for style in doc.styles.element.iter(qn("w:style")):
        sz = style.find(qn("w:rPr") + "/" + qn("w:szCs"))
        if sz is not None:
            styles[style.get(qn("w:styleId"))] = int(sz.get(qn("w:val"))) / 2
    sizes, count = set(), 0
    for note in root.iter(qn("w:footnote")):
        if note.get(qn("w:type")) in {"separator", "continuationSeparator", "continuationNotice"}:
            continue
        for paragraph in note.iter(qn("w:p")):
            style = paragraph.find(qn("w:pPr") + "/" + qn("w:pStyle"))
            default = styles.get(style.get(qn("w:val"))) if style is not None else styles.get("FootnoteText")
            for run in paragraph.iter(qn("w:r")):
                if not "".join(run.itertext()).strip() or not _FA_LETTER.search("".join(run.itertext())):
                    continue
                sz = run.find(qn("w:rPr") + "/" + qn("w:szCs"))
                size = int(sz.get(qn("w:val"))) / 2 if sz is not None else default
                if size is not None and size != FOOTNOTE_SIZE:
                    sizes.add(size)
                    count += 1
    if sizes:
        # Anchor on the first body paragraph that holds a footnote reference.
        anchor = next((index for index, p in enumerate(doc.paragraphs) if p._p.xpath(".//w:footnoteReference")), -1)
        out.append(_issue("KASHAN_FOOTNOTE_FONT", f"اندازهٔ قلم {count} بخش از زیرنویس‌ها {', '.join(f'{s:g}' for s in sorted(sizes))} است؛ زیرنویس با قلم نازنین یا لوتوس و اندازهٔ ۱۰ نوشته شود.",
                          GUIDE.format("۲۶"), loc=f"پاراگراف {anchor + 1}" if anchor >= 0 else "زیرنویس‌ها",
                          par_index=anchor, suggestion="سبک Footnote Text را روی اندازهٔ ۱۰ تنظیم کنید."))


# ----------------------------------------------------------------------------
# page setup and numbering (clauses 13, 14, 28, 29)
# ----------------------------------------------------------------------------
def _section_ranges(doc):
    """(first body paragraph index, sectPr) per section, in order."""
    ranges, start = [], 0
    for index, paragraph in enumerate(doc.paragraphs):
        sect = paragraph._p.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        if sect is not None:
            ranges.append((start, index, sect))
            start = index + 1
    ranges.append((start, len(doc.paragraphs) - 1, doc.element.body.find(qn("w:sectPr"))))
    return ranges


def _has_page_number(doc, ranges, index):
    """Effective default header/footer of a section shows a PAGE field."""
    for position in range(index, -1, -1):
        sect = ranges[position][2]
        if sect is None:
            continue
        refs = [ref for ref in sect.findall(qn("w:headerReference")) + sect.findall(qn("w:footerReference"))
                if ref.get(qn("w:type"), "default") == "default"]
        if not refs:
            continue                      # linked to the previous section
        for ref in refs:
            part = doc.part.related_parts.get(ref.get(qn("r:id")))
            element = getattr(part, "element", None)
            if element is None:
                continue
            fields = [node.text or "" for node in element.iter(qn("w:instrText"))]
            fields += [node.get(qn("w:instr")) or "" for node in element.iter(qn("w:fldSimple"))]
            if any(field.split() and field.split()[0].upper() == "PAGE" for field in fields):
                return True
        return False
    return False


def _page_setup(doc, records, out, full=True):
    import kashan_checker
    ranges = _section_ranges(doc)
    for number, (_, _, sect) in enumerate(ranges, 1):
        if sect is None:
            continue
        margins = sect.find(qn("w:pgMar"))
        if margins is None:
            continue
        wrong = []
        for side, label in (("top", "بالا"), ("bottom", "پایین"), ("right", "راست"), ("left", "چپ")):
            value = margins.get(qn("w:" + side))
            if value and value.lstrip("-").isdigit() and abs(abs(int(value)) * 2.54 / 1440 - MARGIN_CM) > MARGIN_TOLERANCE:
                wrong.append(f"{label} {abs(int(value)) * 2.54 / 1440:.1f}")
        if wrong:
            out.append(_issue("KASHAN_MARGIN", f"حاشیهٔ {'، '.join(wrong)} سانتی‌متر است؛ هر چهار حاشیه ۳ سانتی‌متر باشد.",
                              GUIDE.format("۲۸"), loc=f"بخش {number}", suggestion="Layout ← Margins: بالا، پایین، راست و چپ ۳ سانتی‌متر"))
    # 13-14) no numbers before the table of contents; abjad/alphabet from it
    toc = kashan_checker._find_heading(records, ("فهرست مطالب",))
    if not full or toc is None or toc["par_index"] < 0 or len(ranges) < 2:
        return
    toc_section = next(number for number, (first, last, _) in enumerate(ranges) if first <= toc["par_index"] <= last)
    numbered_front = [number for number in range(toc_section) if _has_page_number(doc, ranges, number)]
    if numbered_front:
        out.append(_issue("KASHAN_PAGE_NUMBERS", "صفحه‌های پیش از فهرست مطالب (بسم‌الله تا چکیده) نباید شماره یا حرف داشته باشند.",
                          GUIDE.format("۱۳"), loc=f"بخش {numbered_front[0] + 1}",
                          suggestion="برای این بخش‌ها Link to Previous را بردارید و شمارهٔ صفحه را حذف کنید."))
    if _has_page_number(doc, ranges, toc_section):
        number_type = ranges[toc_section][2].find(qn("w:pgNumType")) if ranges[toc_section][2] is not None else None
        fmt = number_type.get(qn("w:fmt")) if number_type is not None else "decimal"
        if fmt not in ABJAD_FORMATS:
            out.append(_issue("KASHAN_PAGE_NUMBERS", "از فهرست مطالب به بعد، شماره‌گذاری صفحه با حروف الفبا یا ابجد (الف، ب، ج…) باشد.",
                              GUIDE.format("۱۴"), toc, suggestion="Page Number Format ← حروف ابجد یا الفبا"))
    elif not numbered_front:
        out.append(_issue("KASHAN_PAGE_NUMBERS", "در صفحهٔ فهرست مطالب شمارهٔ صفحه (با حروف الفبا یا ابجد) پیدا نشد.",
                          GUIDE.format("۱۴"), toc, suggestion="از این بخش به بعد شمارهٔ صفحه با حروف ابجد درج کنید."))


def _lines_per_page(doc, records, out):
    """24-26 lines per page; measurable only with exact or minimum spacing."""
    import checker
    import kashan_checker
    start = next((r for r in records if kashan_checker._heading(r) and kashan_checker._CHAPTER.fullmatch(kashan_checker._norm(r["text"]))), None)
    if start is None:
        return
    ranges = _section_ranges(doc)
    estimates = {}
    for record in records[start["index"]:]:
        if record["par_index"] < 0 or len(record["text"]) < 40:
            continue
        paragraph = doc.paragraphs[record["par_index"]]
        line = checker.eff_ppr(doc, paragraph, "w:spacing", "w:line")
        rule = checker.eff_ppr(doc, paragraph, "w:spacing", "w:lineRule") or "auto"
        if not line or rule == "auto":
            continue
        sect = next(s for first, last, s in ranges if first <= record["par_index"] <= last)
        size, margins = sect.find(qn("w:pgSz")) if sect is not None else None, sect.find(qn("w:pgMar")) if sect is not None else None
        if size is None or margins is None:
            continue
        usable = int(size.get(qn("w:h"))) - abs(int(margins.get(qn("w:top"), 0))) - abs(int(margins.get(qn("w:bottom"), 0)))
        estimates.setdefault(int(usable // int(line)), record)
    wrong = {count: record for count, record in estimates.items()
             if not LINES_PER_PAGE[0] <= count <= LINES_PER_PAGE[1]}
    if wrong:
        count, record = sorted(wrong.items())[0]
        out.append(_issue("KASHAN_LINES", f"با فاصلهٔ خطوط این بند حدود {count} سطر در هر صفحه جا می‌شود؛ باید ۲۴ تا ۲۶ سطر باشد.",
                          GUIDE.format("۲۹"), record, suggestion="فاصلهٔ خطوط (Line spacing) را طوری تنظیم کنید که ۲۴ تا ۲۶ سطر در صفحه جا شود."))


# ----------------------------------------------------------------------------
# references (clause 19) and page order (clauses 1, 8, 9, 14, 20)
# ----------------------------------------------------------------------------
def _reference_numbers(records, out):
    import kashan_checker
    start = kashan_checker._find_heading(records, ("منابع", "مراجع", "منابع و مراجع", "منابع و مآخذ", "references", "bibliography"))
    if start is None:
        return
    for record in records[start["index"] + 1:]:
        if kashan_checker._find_heading([record], ("پیوست", "پیوست ها", "پیوستها", "abstract")):
            break
        label = re.match(r"\s*(\[?)\s*([0-9۰-۹]+)\s*(\]?)[.)\-]?\s", record["text"])
        if not label:
            continue
        rest = record["text"][label.end():]
        persian = bool(_FA_LETTER.search(rest[:30]))
        digits = label.group(2)
        wrong_script = any(ch in "0123456789" for ch in digits) if persian else any(ch in "۰۱۲۳۴۵۶۷۸۹" for ch in digits)
        if not (label.group(1) and label.group(3)) or wrong_script:
            want = f"[{digits.translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹')) if persian else digits.translate(_DIGITS)}]"
            out.append(_issue("KASHAN_REFERENCE_NUMBER", "شمارهٔ منبع داخل کروشه [ ] و برای منبع فارسی با رقم فارسی، برای منبع انگلیسی با رقم انگلیسی باشد.",
                              GUIDE.format("۱۹"), record, label.span(0), want))


def _order(records, checklist, out):
    """Sequence of the opening and closing pages as in the official templates."""
    import kashan_checker
    positions = {row["id"]: row["par_index"] for row in checklist
                 if row.get("status") == "detected" and row.get("par_index", -1) >= 0}
    by_paragraph = {r["par_index"]: r for r in records if r["par_index"] >= 0}
    order = [item[0] for item in kashan_checker._MASTER]
    labels = {item[0]: item[1] for item in kashan_checker._MASTER}
    previous = None
    for item in order:
        if item not in positions:
            continue
        if previous and positions[item] < positions[previous]:
            record = by_paragraph.get(positions[item])
            out.append(_issue("KASHAN_PAGE_ORDER", f"«{labels[item]}» پیش از «{labels[previous]}» آمده است؛ ترتیب صفحه‌ها مطابق راهنما نیست.",
                              GUIDE.format("۱، ۸، ۹، ۱۴ و ۲۰"), record,
                              suggestion="ترتیب: بسم‌الله، عنوان فارسی، صورت‌جلسه، تعهدنامه، تقدیم، سپاسگزاری، چکیده، فهرست‌ها، فصل‌ها، منابع، پیوست‌ها، Abstract، عنوان انگلیسی"))
        else:
            previous = item


def check_format(doc, records, degree, scope, checklist, english_window, out):
    """Run every Word-format clause that the document gives evidence for."""
    import kashan_checker
    if doc is None:
        return
    _persian_cover(doc, kashan_checker.persian_cover_window(records), degree, out)
    _english_cover(doc, english_window, degree, out)
    _opening_pages(doc, records, out)
    _body(doc, records, out)
    _footnotes(doc, out)
    _reference_numbers(records, out)
    _lines_per_page(doc, records, out)
    # Margins hold for any file; numbering and page order need the whole thesis.
    _page_setup(doc, records, out, full=scope == "full")
    if scope == "full":
        _order(records, checklist, out)
