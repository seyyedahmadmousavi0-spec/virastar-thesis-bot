# -*- coding: utf-8 -*-
"""
لایه حفاظت از محتوای علمی، فرمول‌ها، ارقام و نمادها (Protected Scientific Content Layer).
اصل بنیادین:
نمادها، فرمول‌های ریاضی و شیمیایی، یکاهای اندازه‌گیری (SI)، ارقام آماری، کدهای نرم‌افزاری،
ارجاعات و پیوندها نباید مشمول قواعد ویرایش متن عمومی، فاصله‌گذاری و تغییر رسم‌الخط شوند.
"""
import re
from dataclasses import dataclass
from typing import List, Tuple


@dataclass(frozen=True)
class ProtectedSpan:
    start: int
    end: int
    kind: str  # formula, unit, chemical, citation, stat, latin, url, greek, number
    text: str


# ۱. عبارات و آدرس‌های اینترنتی و شناسه‌های DOI
URL_DOI_RE = re.compile(
    r"(?:https?://[^\s)\]>\"\'«»]+|"
    r"ftp://[^\s)\]>\"\'«»]+|"
    r"www\.[a-zA-Z0-9\-]+\.[^\s)\]>\"\'«»]+|"
    r"(?:doi\s*:\s*|https?://doi\.org/)?10\.\d{4,9}/[-._;()/:A-Za-z0-9]+|"
    r"[a-zA-Z0-9_\-\.]+\.(?:docx?|pdf|xlsx?|csv|png|jpe?g|svg|py|cpp|java|json|xml))(?!\w)",
    re.IGNORECASE
)

EMAIL_RE = re.compile(r"(?<![\w.+-])[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,63}(?![\w.-])")
CODE_RE = re.compile(r"```[\s\S]*?```|~~~[\s\S]*?~~~|`[^`\n]+`")
_DIGIT = r"0-9\u0660-\u0669\u06f0-\u06f9"
_NUM = rf"[+\-\u2212]?[{_DIGIT}]+(?:[.\u066b][{_DIGIT}]+)?(?:[eE][+\-\u2212]?[{_DIGIT}]+)?"
_LETTER = r"A-Za-z\u0370-\u03ff"
_VARIABLE = rf"[{_LETTER}]{{1,3}}(?:[_^](?:[{{(]?[{_LETTER}{_DIGIT}]+[}})]?))?(?![{_LETTER}])"
_TERM = rf"(?:{_NUM}|{_VARIABLE})"

# ۲. فرمول‌های ریاضی درون‌متنی و برون‌متنی (LaTeX / Math / Equations)
MATH_FORMULA_RE = re.compile(
    r"(?:\$\$[^\$]+\$\$|"                          # LaTeX block math $$ ... $$
    r"\$[^\$\n]{1,150}\$|"                        # LaTeX inline math $ ... $
    r"\\\[[\s\S]*?\\\]|"                          # LaTeX display math \[ ... \]
    r"\\\([^\n]*?\\\)|"                           # LaTeX inline math \( ... \)
    rf"(?<!\w){_VARIABLE}\s*(?:<=|>=|!=|==|[<>=≤≥≈~])\s*{_TERM}"
    rf"(?:\s*[*/±×÷+\-]\s*{_TERM})*(?!\w))"  # x = 5.2, E = mc^2, p < 0.05
)

INTERVAL_RE = re.compile(rf"[\[(]\s*{_NUM}\s*[,،]\s*{_NUM}\s*[\])]")

# ۳. ارجاعات درون‌متنی (Citations)
_FA_LETTER = r"\u0621-\u063a\u0641-\u064a\u067e\u0686\u0698\u06a9\u06af\u06cc\u06c0\u200c"
_NAME = rf"(?:[A-Z][A-Za-z'’\-]+|[{_FA_LETTER}]+)"
_AUTHOR = rf"{_NAME}(?:\s+(?:{_NAME}|de|del|van|von|da|and|&|و)){{0,6}}(?:\s+(?:et\s+al\.?|و\s+همکاران))?"
_AUTHOR_YEAR = rf"{_AUTHOR}\s*[,،\s]\s*[{_DIGIT}]{{4}}[a-z]?(?:\s*[:،,]\s*(?:pp?\.?\s*)?[{_DIGIT}]+(?:\s*[-–—]\s*[{_DIGIT}]+)?)?"
CITATION_RE = re.compile(
    rf"(?:\[[{_DIGIT}\s,،;؛\-–—]+\]|"
    rf"\({_AUTHOR_YEAR}(?:\s*[;؛]\s*{_AUTHOR_YEAR})*\))"
)

# ۴. شاخص‌های آماری و ریاضیاتی (Statistical indicators)
STAT_RE = re.compile(
    rf"(?<!\w)(?:[pPrRtTfFzZnN]\s*(?:<=|>=|[<>=≤≥≈])\s*{_NUM}|"
    rf"[{_DIGIT}]+(?:[.\u066b][{_DIGIT}]+)?[%٪]\s*CI\s*(?:\[[^\]\n]+\]|\([^\)\n]+\))|"
    rf"(?:Mean|SD|SEM|IQR|OR|RR|HR|df|AUC|ROC)\s*=\s*{_NUM}|"
    rf"{_NUM}\s*±\s*{_NUM})(?!\w)",
    re.IGNORECASE
)

# ۵. اصطلاحات لاتین علمی و پزشکی متداول (Academic Latin Phrases & Abbreviations)
LATIN_ACADEMIC_RE = re.compile(
    r"(?<!\w)(?:et\s+al\.|i\.e\.|e\.g\.|cf\.|vs\.|ibid\.|op\.\s*cit\.|in\s+vitro|in\s+vivo|"
    r"in\s+situ|ex\s+vivo|de\s+novo|ad\s+hoc|per\s+se|inter\s+alia|vice\s+versa|ca\.|"
    r"status\s+quo|a\s+priori|a\s+posteriori)(?!\w)",
    re.IGNORECASE
)

# ۶. فرمول‌های شیمیایی و نام‌های زیست‌مولکولی رایج
CHEMICAL_RE = re.compile(
    r"(?<!\w)(?:H2O|CO2|CaCO3|NaCl|H2SO4|NaOH|HCl|CH4|C6H12O6|NOx|SO2|"
    r"Fe2\+|Fe3\+|Ca2\+|Mg2\+|K\+|Na\+|Cl\-|PO43\-|SO42\-|"
    r"mRNA|tRNA|rRNA|cDNA|dsDNA|ssRNA|p53|BRCA1|BRCA2|TNF-α|TNF-alpha|IL-6|IL-10|COVID-19|SARS-CoV-2)(?!\w)"
)

# ۷. اعداد همراه با واحدهای اندازه‌گیری استاندارد علمی (SI Units & Compound Units)
# مثال: 15 mg/dL, 25 °C, 50 µm, 100 nm, 4.2 GHz, 70 kg, 220 V, 50 Hz, 10 ppm, 95%, 15٪
_PREFIX = r"(?:da|[µμumcdhkMGTnp])?"
_EXPONENT = r"(?:[²³23]|[⁻⁺]?[⁰¹²³⁴⁵⁶⁷⁸⁹]+|\^[+\-−]?[0-9]+)?"
# Unit spelling is case sensitive: lower-case 'a' in prose is not ampere A.
# Compound units must be consumed as a whole, including their denominator.
_UNIT_ATOM = (
    rf"(?:°\s*[CFKcf]|{_PREFIX}(?:mol|Hz|Pa|[gLlmsVAWJK])|"
    r"[munpµμ]?M|bar|atm|mmHg|torr|rpm|[cC][fF][uU]|IU|"
    r"(?:k|M|G)?bp|kDa|Da|ppm|ppb|ppt|dB|pH|‰|%|٪|"
    r"(?i:sec|second|seconds|min|minute|minutes|hr|hour|hours|day|days|week|weeks|month|months|year|years)|h)"
    + _EXPONENT
)
SCIENTIFIC_VALUE_UNIT_RE = re.compile(
    rf"(?<![\w.]){_NUM}(?:\s*[-–—~]\s*{_NUM})?\s*"
    rf"{_UNIT_ATOM}(?:\s*(?:/|·|⋅)\s*{_UNIT_ATOM})*(?![\w/])"
)

# ۸. حروف یونانی و نمادهای ریاضی (Greek letters & symbols)
GREEK_SYMBOLS_RE = re.compile(
    r"[\u0370-\u03FF\u1F00-\u1FFF\u2200-\u22FF\u2A00-\u2AFF]"
)


class ProtectedContentDetector:
    """
    موتور شناسایی و استخراج محدوده‌های حفاظت‌شده در متن.
    تمام بازه‌های حساس به صورت بازه‌های [start, end) ثبت می‌شوند تا قواعد ویرایشی عمومی روی آن‌ها اعمال نشوند.
    """

    @classmethod
    def find_spans(cls, text: str) -> List[ProtectedSpan]:
        """یافتن تمام محدوده‌های محافظت‌شده در متن."""
        if not text:
            return []

        spans: List[ProtectedSpan] = []

        # ۱. آدرس‌های وب و DOI
        for m in URL_DOI_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "url_doi", m.group(0)))
        for m in EMAIL_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "email", m.group()))
        for m in CODE_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "code", m.group()))

        # ۲. فرمول‌های ریاضی و لاتک
        for m in MATH_FORMULA_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "formula", m.group(0)))
        for m in INTERVAL_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "interval", m.group()))

        # ۳. ارجاعات متنی
        for m in CITATION_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "citation", m.group(0)))

        # ۴. شاخص‌های آماری
        for m in STAT_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "stat", m.group(0)))

        # ۵. اصطلاحات لاتین علمی
        for m in LATIN_ACADEMIC_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "latin_academic", m.group(0)))

        # ۶. فرمول‌های شیمیایی
        for m in CHEMICAL_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "chemical", m.group(0)))

        # ۷. اعداد همراه با واحدهای علمی
        for m in SCIENTIFIC_VALUE_UNIT_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "scientific_unit", m.group(0)))

        for m in GREEK_SYMBOLS_RE.finditer(text):
            spans.append(ProtectedSpan(m.start(), m.end(), "greek", m.group(0)))
        for m in re.finditer(r"[-+−]?[0-9٠-٩۰-۹]+(?:[.٫][0-9٠-٩۰-۹]+)*(?:[eE][−+-]?[0-9]+)?[%٪²³⁰¹⁴⁵⁶⁷⁸⁹]*", text):
            spans.append(ProtectedSpan(m.start(), m.end(), "number", m.group(0)))
        for m in re.finditer(r"(?<![A-Za-z])(?:[A-Z][a-z]?[₀-₉0-9]+(?:[A-Z][a-z]?[₀-₉0-9]*)*|(?:[A-Z][a-z]?){2,}[₀-₉0-9]*)(?![A-Za-z])", text):
            spans.append(ProtectedSpan(m.start(), m.end(), "symbol", m.group(0)))

        # ادغام و حذف بازه‌های تداخلی (Interval Merging)
        return [ProtectedSpan(s.start, s.end, s.kind, text[s.start:s.end]) for s in cls._merge_spans(spans)]

    @staticmethod
    def _merge_spans(spans: List[ProtectedSpan]) -> List[ProtectedSpan]:
        """مرتب‌سازی و ادغام بازه‌های همپوشان."""
        if not spans:
            return []

        sorted_spans = sorted(spans, key=lambda s: (s.start, -(s.end - s.start)))
        merged: List[ProtectedSpan] = []

        current = sorted_spans[0]
        for next_span in sorted_spans[1:]:
            if next_span.start <= current.end:
                # تداخل یا پوشش
                new_end = max(current.end, next_span.end)
                current = ProtectedSpan(
                    start=current.start,
                    end=new_end,
                    kind="+".join(sorted(set(current.kind.split("+")) | set(next_span.kind.split("+")))),
                    text=""  # در صورت ادغام متن با برش رشته پر می‌شود
                )
            else:
                merged.append(current)
                current = next_span
        merged.append(current)
        return merged

    @classmethod
    def is_protected(cls, start: int, end: int, spans: List[ProtectedSpan]) -> bool:
        """بررسی آیا بازه داده‌شده با هیچ‌یک از بازه‌های محافظت‌شده تداخل دارد یا خیر."""
        if start < 0 or end <= start:
            return False
        for s in spans:
            if not (end <= s.start or start >= s.end):
                return True
        return False

    @classmethod
    def filter_text_spans(cls, text: str, matches: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
        """فیلتر کردن مواردی که درون محدوده‌های علمی قرار دارند."""
        spans = cls.find_spans(text)
        return [m for m in matches if not cls.is_protected(m[0], m[1], spans)]


def quoted_spans(text):
    return [ProtectedSpan(m.start(), m.end(), "quote", m.group())
            for m in re.finditer(r'''«[^»]*»|“[^”]*”|"[^"\n]*"|‘[^\n]*?’(?!\w)|(?<!\w)'[^'\n]+'(?!\w)''', text)]
