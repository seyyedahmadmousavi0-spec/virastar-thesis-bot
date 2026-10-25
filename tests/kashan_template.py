# -*- coding: utf-8 -*-
"""Rebuild the official Kashan template (Bahman 1398) as a filled .docx.

Usage: python tests/kashan_template.py OUTPUT.docx master|phd clean|seeded
Lines marked SEED break one clause of the guide in the seeded variant.
Layout follows the 3-page guide: Nazanin 12, Titr 16/14/12 headings, 3 cm
margins, exact line spacing for ~25 lines, sections with no numbering on the
opening pages, abjad numbering from the table of contents, Arabic afterwards.
"""
import struct
import sys
import zlib

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

Z = "‌"
out, degree, variant = sys.argv[1], sys.argv[2], sys.argv[3]
seeded = variant == "seeded"
logo = out + ".png"


def png(path):
    """A 4x4 grey PNG standing in for the logo and basmala images."""
    raw = b"".join(b"\x00" + b"\x80" * 12 for _ in range(4))
    chunk = lambda kind, data: struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    open(path, "wb").write(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 4, 4, 8, 2, 0, 0, 0))
                           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


png(logo)
doc = Document()
from docx.enum.style import WD_STYLE_TYPE
doc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)


def setup(section, numbering=None, start=None):
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, side, Cm(3))
    sect = section._sectPr
    for old in sect.findall(qn("w:pgNumType")):
        sect.remove(old)
    if numbering:
        node = OxmlElement("w:pgNumType")
        node.set(qn("w:fmt"), numbering)
        if start:
            node.set(qn("w:start"), str(start))
        sect.append(node)


def page_field(section):
    section.footer.is_linked_to_previous = False
    paragraph = section.footer.paragraphs[0]
    for kind, text in (("begin", None), ("instr", " PAGE "), ("end", None)):
        run = paragraph.add_run()
        node = OxmlElement("w:instrText" if kind == "instr" else "w:fldChar")
        if text:
            node.text = text
        else:
            node.set(qn("w:fldCharType"), kind)
        run._r.append(node)


def para(text, *, font="B Nazanin", size=12, bold=False, italic=False, center=False, latin=False, style=None):
    p = doc.add_paragraph(style=style)
    if not latin:
        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(26)
    for part in text.split("**"):
        if not part:
            bold = not bold
            continue
        r = p.add_run(part.replace("_", Z))
        rpr = r._r.get_or_add_rPr()
        fonts = OxmlElement("w:rFonts")
        for key in ("w:ascii", "w:hAnsi"):
            fonts.set(qn(key), "Times New Roman")
        fonts.set(qn("w:cs"), font)
        rpr.insert(0, fonts)
        r.bold, r.italic = bold, italic
        r.font.size = Pt(size)
        sz = OxmlElement("w:szCs")
        sz.set(qn("w:val"), str(size * 2))
        rpr.append(sz)
        if bold:
            rpr.append(OxmlElement("w:bCs"))
        if not latin:
            rpr.append(OxmlElement("w:rtl"))
        bold = not bold
    return p


def picture():
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(logo, width=Cm(3))


def new_section(numbering=None, start=None, numbered=False):
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    setup(section, numbering, start)
    if numbered:
        page_field(section)
    else:
        section.footer.is_linked_to_previous = False
    return section


title = "رساله دکتری" if degree == "phd" else "پایان_نامه کارشناسی ارشد"
en_title = "Ph.D. Dissertation" if degree == "phd" else "M.Sc. Thesis"

setup(doc.sections[0])
picture()                                                   # 1) basmala image
new_section()
picture()                                                   # 2) logo
para("دانشکده مهندسی مکانیک", bold=True, center=True,
     font="IranNastaliq" if seeded else "B Nazanin")        # SEED cover font
para("گروه آموزشی حرارت و سیالات", bold=True, center=True)
para(f"{title} در رشته" if not seeded else f"**{title}** در رشته", bold=not seeded, center=True)  # SEED bold
para("مهندسی مکانیک - تبدیل انرژی", bold=True, center=True)
para("بررسی عددی انتقال حرارت در مبدل_های حرارتی صفحه_ای", bold=True, center=True, size=16)
para("استاد راهنما", bold=True, center=True)
para("دکتر علی احمدی", bold=True, center=True)
para("استاد مشاور", bold=True, center=True)
para("دکتر مریم رضایی", bold=True, center=True)
para("به وسیله:" if seeded else "توسط:", bold=True, center=True)          # SEED by
para("سیداحمد موسوی", bold=True, center=True)
para("1403" if seeded else "شهریور ۱۴۰۳", bold=True, center=True)          # SEED date
new_section()
para(f"صورتجلسه دفاع از {'رساله دکتری' if degree == 'phd' else 'پایان نامه کارشناسی ارشد'}", bold=True, center=True)
para("جلسه دفاع از پایان_نامهٔ آقای سیداحمد موسوی با حضور هیأت داوران تشکیل و با درجهٔ عالی مورد تأیید قرار گرفت.")
new_section()
para("تعهدنامه", bold=True, center=True)
para("مطالب مندرج در این پایان_نامه حاصل تحقیق و پژوهش اینجانب بوده و صحت مطالب نگارش شده مورد تأیید است.")
new_section()
para("تقدیم", bold=True)
para("تقدیم به پدر و مادرم.")
new_section()
para("سپاس_گزاری", bold=True)
para("از استاد راهنمای گرامی، **دکتر علی احمدی**، و استاد مشاور، **دکتر مریم رضایی**، سپاسگزارم." if not seeded
     else "از استاد راهنمای گرامی، دکتر علی احمدی، و استاد مشاور، دکتر مریم رضایی، سپاسگزارم.")   # SEED ack bold
new_section()
para("چکیده", bold=True)
para("در این پژوهش انتقال حرارت در مبدل_های حرارتی صفحه_ای به روش عددی بررسی شد و نتایج با داده_های تجربی مقایسه گردید. "
     "نتایج نشان داد که افزایش عدد رینولدز ضریب انتقال حرارت را افزایش می_دهد.")
para("**واژگان کلیدی:** مبدل حرارتی، انتقال حرارت، شبیه_سازی عددی، عدد رینولدز" if not seeded
     else "**واژگان کلیدی:** مبدل حرارتی، انتقال حرارت")                                          # SEED keywords
new_section("arabicAbjad", 1, numbered=True)                # 14) abjad from TOC
para("فهرست مطالب", bold=True, center=True)
para("فصل اول: مقدمه\t۱", style="toc 1", bold=True)
para("منابع و مراجع\t۷", style="toc 1")
para("فهرست جداول" if seeded else "فهرست جدول_ها", bold=True, center=True)               # SEED list title
para("جدول ۱-۱ مشخصات هندسی مبدل\t۳", style="toc 1")
para("فهرست شکل_ها", bold=True, center=True)
para("شکل ۱-۱ طرح_وارهٔ مبدل\t۴", style="toc 1")
para("فهرست نشانه_های اختصاری", bold=True, center=True)
para("ر.ک. رجوع کنید به")
body = new_section("decimal", 1, numbered=True)
if seeded:
    body.left_margin = Cm(2.5)                              # SEED margin
para("فصل اول", font="B Titr", size=16, bold=True, center=True, style="Heading 1")
para("مقدمه و تاریخچه", font="B Titr", size=16, bold=True, center=True)
para("۱-۱- پیشینهٔ پژوهش", font="B Titr", size=14 if seeded else 16, bold=True, style="Heading 2")   # SEED heading size
para("مبدل_های حرارتی صفحه_ای به دلیل راندمان بالا در صنایع گوناگون به کار می_روند (Holmberg et al., 2004). "
     "مشخصات هندسی مبدل در جدول ۱-۱ آمده است.",
     font="B Mitra" if seeded else "B Nazanin")              # SEED body font
para("۱-۱-۱- مبدل_های صفحه_ای", font="B Titr", size=14, bold=True, style="Heading 3")
para("طرح_وارهٔ مبدل در شکل ۱-۱ نشان داده شده است و جزئیات آن در ادامه آمده است (رحیمیان، ۱۳۹۲).")
para("جدول ۱-۱ مشخصات هندسی مبدل", center=True)
para("شکل ۱-۱ طرح_وارهٔ مبدل", center=True)
new_section("decimal", numbered=True)
para("منابع و مراجع", font="B Titr", size=16, bold=True)
para("رحیمیان، ج. (۱۳۹۲). جنبه_های صوری و معنایی عناصر وجهی در جمله_های فارسی. زبان و زبان_شناسی، ۱، ۳۳-۵۰.")
para("Holmberg, D., Orbuch, T., & Veroff, J. (2004). Thrice-told tales: Married couples tell their stories. Mahwah, NJ: Erlbaum.", latin=True)
if seeded:
    para("[1] Stein, A. (1997). Sex after sexuality. London: Sage.", latin=True)            # SEED numbering
new_section("decimal", numbered=True)
para("پیوست_ها", bold=True, center=True)
para("پیوست ۱: داده_های خام آزمایش.")
new_section()
para("Abstract", bold=True, center=True, latin=True)
para("In this study, heat transfer in plate heat exchangers was investigated numerically.", latin=True)
para("**Keywords:** heat exchanger, heat transfer, numerical simulation, Reynolds number", latin=True)
new_section()
picture()
para("Faculty of Mechanical Engineering", bold=True, center=True, latin=True)
para("Department of Thermal and Fluid Sciences", bold=True, center=True, latin=True)
para(f"{en_title} in", bold=not seeded, center=True, latin=True)                           # SEED en bold
para("Mechanical Engineering - Energy Conversion", bold=True, center=True, latin=True)
para("Numerical Study of Heat Transfer in Plate Heat Exchangers", bold=True, center=True, latin=True,
     italic=seeded)                                                                        # SEED italic
para("Supervisor:", bold=True, center=True, latin=True)
para("Dr. Ali Ahmadi", bold=True, center=True, latin=True)
para("By:", bold=True, center=True, latin=True)
para("Seyed Ahmad Mousavi", bold=True, center=True, latin=True)
para("The University of Kashan" if seeded else "September 2024", bold=True, center=True, latin=True)  # SEED The + date
doc.save(out)
