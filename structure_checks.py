# -*- coding: utf-8 -*-
"""Document-wide consistency checks for Word theses.

* bibliography entries that are never cited in the body;
* one term written both joined and with ZWNJ («پرسشنامه» / «پرسش‌نامه»);
* table and figure captions never referred to in the body text.

Each check abstains when the document does not give it reliable evidence,
and every finding is a confidence-B review item anchored on an existing
paragraph.
"""
from collections import Counter, defaultdict
import re

from analysis_model import Issue

ZWNJ = "‌"
_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
_YEAR = re.compile(r"(?<![0-9])((?:1[3-4]|19|20)[0-9]{2})[a-zأ-ی]?(?![0-9])")
_WORD = re.compile(r"[A-Za-z؀-ۿ" + ZWNJ + "'-]+")
_CAPTION = re.compile(r"^\s*(جدول|شکل|نمودار|تصویر)\s*(?:شماره\s*)?\(?([0-9۰-۹]+(?:\s*[-–ـ.]\s*[0-9۰-۹]+)?)\)?")
_MARKS = re.compile("[ً-ٰٕـ]")
_CITATION_WINDOW = 90
_MIN_REFERENCES = 3


def _plain(text):
    text = text.translate(_DIGITS).replace("ي", "ی").replace("ك", "ک")
    return " ".join(text.replace(ZWNJ, " ").split()).casefold()


def _first_author(entry):
    """Surname that starts a bibliography entry, before its first comma/period."""
    head = re.split(r"[،,.(]", entry, maxsplit=1)[0]
    words = _WORD.findall(head)
    return words[0].casefold() if words and len(words[0]) >= 2 else None


def uncited_references(paragraphs, reference_index, end_index, numbered_labels=None):
    """Return issues for bibliography entries that the body never cites."""
    if reference_index is None:
        return []
    entries = [(index, paragraphs[index].text) for index in range(reference_index + 1, end_index)
               if paragraphs[index].text.strip()]
    if len(entries) < _MIN_REFERENCES:
        return []
    body = "\n".join(_plain(paragraph.text) for paragraph in paragraphs[:reference_index])
    issues = []
    if numbered_labels:
        cited = set()
        for match in re.finditer(r"\[([0-9,،\s\-–]+)\]", body):
            for part in re.split(r"[,،]", match.group(1)):
                bounds = [int(value) for value in re.findall(r"[0-9]+", part)]
                if bounds and bounds[-1] - bounds[0] <= 1000:
                    cited.update(range(bounds[0], bounds[-1] + 1))
        if not cited:
            return []      # another citation style; numbers cannot be judged
        for (index, text), label in zip(entries, numbered_labels):
            if label not in cited:
                issues.append(_uncited(index, text, f"منبع شمارهٔ {label}"))
        return issues
    checked = 0
    for index, text in entries:
        plain = _plain(text)
        author, year = _first_author(plain), _YEAR.search(plain)
        if not author or not year:
            continue                     # unparsable entry: no claim
        checked += 1
        pattern = re.escape(author) + r"[^\n]{0," + str(_CITATION_WINDOW) + r"}?" + year.group(1)
        if not re.search(r"(?<![\w])" + pattern, body):
            shown = _WORD.findall(re.split(r"[،,.(]", text, maxsplit=1)[0])
            issues.append(_uncited(index, text, f"«{shown[0] if shown else author}» ({year.group(1)})"))
    # If most entries look uncited, the citation style was not understood.
    if checked and len(issues) > checked / 2:
        return []
    return issues


def _uncited(index, text, label):
    return Issue("reference_uncited", "warn",
                 f"در متن پیش از فهرست منابع، ارجاعی به {label} پیدا نشد؛ اگر این منبع استفاده نشده، "
                 "آن را حذف یا در متن به آن ارجاع دهید.",
                 f"پاراگراف {index + 1}", text[:60], index, category="citation_consistency",
                 confidence="B")


def inconsistent_terms(paragraphs, skip=()):
    """One term written joined in some places and with ZWNJ in others."""
    joined, split = Counter(), defaultdict(list)
    for index, paragraph in enumerate(paragraphs):
        if index in skip:
            continue
        for match in _WORD.finditer(paragraph.text):
            word = _MARKS.sub("", match.group()).replace("ي", "ی").replace("ك", "ک").strip(ZWNJ)
            if ZWNJ in word:
                split[word.replace(ZWNJ, "")].append((index, match.start(), match.end(), word))
            else:
                joined[word] += 1
    issues = []
    for key, occurrences in split.items():
        if len(key) < 5 or not joined[key]:
            continue
        # Plural/verb affixes are covered by their own rules.
        parts = occurrences[0][3].split(ZWNJ)
        if parts[-1] in {"ها", "های", "هایی", "اند", "ایم", "اید", "ای", "ام", "ی"} or parts[0] in {"می", "نمی"}:
            continue
        index, start, end, word = occurrences[0]
        issues.append(Issue(
            "term_inconsistency", "warn",
            f"این واژه در سند به دو صورت آمده است: «{word}» ({len(occurrences)} بار) و «{key}» "
            f"({joined[key]} بار)؛ یک صورت را در کل متن یکسان به کار ببرید.",
            f"پاراگراف {index + 1}", paragraphs[index].text[max(0, start - 18):end + 18], index,
            "یک صورت واحد (ترجیحاً با نیم‌فاصله)", word, category="۱. غلط املایی و تایپوگرافی",
            confidence="B", char_start=start, char_end=end, location_status="exact"))
    return issues


def _number_key(value):
    return tuple(int(part) for part in re.findall(r"[0-9]+", value.translate(_DIGITS)))


def unreferenced_captions(paragraphs, skip=()):
    """Captions whose «جدول/شکل N» never occurs in body text."""
    captions, mentions = [], defaultdict(set)
    kinds = {"تصویر": "شکل", "نمودار": "شکل"}
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text
        match = _CAPTION.match(text)
        if index in skip:
            continue
        if match and len(text.split()) < 40:
            captions.append((index, kinds.get(match.group(1), match.group(1)), _number_key(match.group(2)), text))
            body = text[match.end():]
        else:
            body = text
        for kind in ("جدول", "شکل", "نمودار", "تصویر"):
            if kind not in body:
                continue
            # «جدول ۲-۳»، «جدول‌های ۲-۳ و ۲-۴»، «(شکل ۱)»
            for number in re.finditer(r"[0-9۰-۹]+(?:\s*[-–ـ.]\s*[0-9۰-۹]+)?", body):
                mentions[kinds.get(kind, kind)].add(_number_key(number.group()))
    if not any(mentions.values()):
        return []      # the text never cites captions by number at all
    issues = []
    for index, kind, key, text in captions:
        if key and key not in mentions[kind]:
            issues.append(Issue(
                "caption_unreferenced", "info",
                f"در متن به «{text.strip()[:40]}» ارجاع داده نشده است؛ هر {kind} باید پیش از آمدن، در متن معرفی شود.",
                f"پاراگراف {index + 1}", text[:60], index, category="۷. جدول‌ها، نمودارها و شکل‌ها",
                confidence="B"))
    return issues
