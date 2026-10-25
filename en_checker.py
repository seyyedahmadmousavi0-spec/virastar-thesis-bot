# -*- coding: utf-8 -*-
"""
ماژول بررسی و ویرایش متون علمی انگلیسی (English Academic Text Checker).
شامل بررسی:
- گرامر و تطابق فاعل و فعل (EN_SVA)
- حروف تعریف a/an با در نظر گرفتن استثناهای آوایی (EN_ARTICLE)
- حروف اضافه متداول علمی (EN_PREP)
- نشانه‌گذاری و فاصله‌گذاری ایمن (EN_PUNCT)
- لحن و سبک آکادمیک (EN_TONE)
- حشو و زوائد در انگلیسی (EN_WORDINESS)
- تقارن پرانتزها و براکت‌ها (EN_BRACKET)
"""
import re
from typing import List

from analysis_model import ConfidenceLevel, Issue
from rule_registry import FindingBuffer, bounded_findings
from protected_content import ProtectedContentDetector, quoted_spans

# واژه‌ها و عبارات غیررسمی در نگارش آکادمیک انگلیسی و جایگزین‌های پیشنهادی
INFORMAL_ACADEMIC_REPLACEMENTS = {
    r"\ba lot of\b": ("numerous / a considerable amount of", "در نگارش علمی از عبارات دقیق‌تر مانند numerous یا substantial استفاده کنید."),
    r"\blots of\b": ("numerous / substantial", "عبارت غیررسمی است؛ از numerous یا considerable استفاده کنید."),
    r"\bkind of\b": ("somewhat / relatively", "عبارت عامیانه در نگارش علمی."),
    r"\bsort of\b": ("partially / relatively", "عبارت عامیانه در نگارش علمی."),
    r"\bget\b": ("obtain / acquire / receive", "فعل get در نگارش علمی معمولاً با افعال دقیق‌تر مانند obtain جایگزین می‌شود."),
    r"\bgot\b": ("obtained / acquired / received", "فعل got غیررسمی است."),
    r"\bgive\b": ("provide / yield / present", "فعل give را با provide یا present جایگزین کنید."),
    r"\bgave\b": ("provided / yielded / presented", "فعل gave غیررسمی است."),
    r"\btill\b": ("until", "واژه till عامیانه است؛ از until استفاده کنید."),
    r"\bkids\b": ("children / adolescents", "واژه عامیانه است."),
    r"\bcops\b": ("police officers", "واژه عامیانه است."),
    r"\bbad\b": ("adverse / unfavorable / detrimental", "صفت bad بسیار ساده است؛ از adverse یا detrimental استفاده کنید."),
    r"\bgood\b": ("beneficial / favorable / advantageous", "صفت good ساده است؛ از favorable یا beneficial استفاده کنید."),
    r"\bbig\b": ("significant / substantial / large-scale", "صفت big در نگارش آکادمیک با substantial یا considerable جایگزین می‌شود."),
    r"\bhuge\b": ("substantial / extensive", "صفت huge محاوره‌ای است."),
    r"\bstuff\b": ("materials / elements / substances", "واژه stuff کاملاً غیرعلمی است."),
    r"\bthings\b": ("factors / aspects / variables", "واژه things مبهم است؛ مفهوم دقیق را بنویسید."),
    r"\bso\b(?=\s+[a-z]+)": ("therefore / thus / consequently", "به‌جای so برای بیان نتیجه، از therefore یا consequently استفاده کنید."),
    r"\btotally\b": ("completely / entirely", "واژه totally غیررسمی است."),
    r"\breally\b": ("markedly / substantially", "واژه really لحن عامیانه به متن می‌دهد."),
}

# عبارات حشو و پرگو در انگلیسی (Wordiness and Redundancy)
WORDINESS_PATTERNS = {
    r"\bin order to\b": ("to", "حذف کلمات اضافی: «in order to» را می‌توان به «to» خلاصه کرد."),
    r"\bdue to the fact that\b": ("because / since", "عبارت طولانی؛ از «because» استفاده کنید."),
    r"\bat this point in time\b": ("now / currently", "عبارت حشو؛ از «currently» استفاده کنید."),
    r"\bat the present time\b": ("currently / presently", "عبارت حشو؛ از «currently» استفاده کنید."),
    r"\bhas the ability to\b": ("can", "حشو؛ می‌توان با «can» ساده‌تر بیان کرد."),
    r"\ba majority of\b": ("most", "می‌توان به «most» خلاصه کرد."),
    r"\bfor the purpose of\b": ("for / to", "عبارت حشو."),
    r"\bin spite of the fact that\b": ("although / despite", "از «although» استفاده کنید."),
    r"\bthe reason why is that\b": ("because", "عبارت حشو."),
    r"\bperiod of time\b": ("period / time", "حشو (Tautology)."),
    r"\bfuture plans\b": ("plans", "حشو."),
    r"\bpast history\b": ("history", "حشو."),
    r"\bclose proximity\b": ("proximity", "حشو."),
}

# خطاهای رایج حروف تعریف a/an
ARTICLE_ERRORS = [
    (re.compile(r"\b(a)\s+([aeiouAEIOU][a-z]+)\b"), "an", "پیش از کلماتی که با حرف صدادار شروع می‌شوند از «an» استفاده می‌شود."),
    (re.compile(r"\b(an)\s+([bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZ][a-z]+)\b"), "a", "پیش از کلماتی که با حرف بی‌صدا شروع می‌شوند از «a» استفاده می‌شود."),
]

# خطاهای رایج حروف اضافه (Prepositions) در متون آکادمیک
PREPOSITION_ERRORS = [
    (r"\binterested on\b", "interested in", "حرف اضافه صحیح interested کلمه in است."),
    (r"\baccording with\b", "according to", "ترکیب صحیح «according to» است."),
    (r"\bdepend of\b", "depend on", "ترکیب صحیح «depend on» است."),
    (r"\bconsist in\b", "consist of", "ترکیب صحیح «consist of» است."),
    (r"\bresponsible of\b", "responsible for", "ترکیب صحیح «responsible for» است."),
    (r"\bsimilar with\b", "similar to", "ترکیب صحیح «similar to» است."),
    (r"\bdifferent than\b", "different from", "در نگارش علمی استاندارد «different from» ارجح است."),
    (r"\bcomply to\b", "comply with", "ترکیب صحیح «comply with» است."),
    (r"\bregardless about\b", "regardless of", "ترکیب صحیح «regardless of» است."),
]

# خطاهای تطابق فاعل و فعل و ساختارهای جمع/مفرد رایج
SVA_ERRORS = [
    (r"\b(these|those)\s+(result|datum|finding|study|method|effect)\b", "جمع اسم بعد از these/those", "کلمه پس از these/those باید جمع باشد (مانند these results)."),
    (r"\b(this|that)\s+(results|data|findings|studies|methods|effects)\b", "مفرد بعد از this/that", "کلمه پس از this/that باید مفرد باشد."),
    (r"\b(each|every)\s+([a-z]+s)\b", "اسم مفرد بعد از each/every", "بعد از each و every اسم به صورت مفرد می‌آید."),
    (r"\bmany\s+(research|information|equipment)\b", "اسم غیرقابل شمارش با many", "کلماتی مانند research/information غیرقابل شمارش هستند؛ از much یا a great deal of استفاده کنید."),
]

# قوانین نشانه‌گذاری انگلیسی
PUNCTUATION_RULES = [
    (re.compile(r"\s+([,.:;?!])"), "فاصله قبل از علامت نگارشی نباید بیاید."),
    (re.compile(r"([,;:?!])(?=[A-Za-z])"), "بعد از علائم نگارشی باید یک فاصله قرار گیرد."),
    (re.compile(r"\b([A-Za-z]{2,})\s+\1\b", re.IGNORECASE), "تکرار ناخواسته یک کلمه انگلیسی پشت‌سرهم."),
]


def _tone_term_is_technical(pattern, match, text):
    """Exclude a few explicit technical senses; this is not semantic parsing."""
    term = match.group().lower()
    tail = text[match.end():]
    if term == "good" and re.match(r"\s+(?:agreement|fit|clinical|laboratory|manufacturing|practice)\b", tail, re.I):
        return True
    if term == "big" and re.match(r"\s+(?:data|bang|five)\b", tail, re.I):
        return True
    if term == "so" and re.match(r"\s+(?:that|far|much|many|as|to)\b", tail, re.I):
        return True
    if term == "till" and re.search(r"\bglacial\s+$", text[:match.start()], re.I):
        return True
    if term in {"kind of", "sort of"}:
        return bool(re.search(r"\b(?:a|an|the|this|that|another|any|some|each|every|what)\s+(?:[A-Za-z]+\s+){0,2}$",
                              text[max(0, match.start() - 40):match.start()], re.I))
    return False


def _expected_article(word):
    """Return a known reading, or abstain for variable acronym pronunciations."""
    lower = word.lower()
    if re.match(r"^[A-Za-z]-", word):
        return "an" if word[0].upper() in "AEFHILMNORSX" else "a"
    if word.isupper():
        if word in {"MRI", "RNA", "FDA", "HIV", "HTML", "ICU", "X"} or len(word) == 1 and word in "AEFHILMNORSX":
            return "an"
        if word in {"DNA", "CT", "PCR"}:
            return "a"
        return None
    if lower.startswith(("honest", "honor", "honour", "heir")) or re.match(r"^hour(?:s|ly|glass|long)?$", lower) or lower.startswith("hour-"):
        return "an"
    if (lower in {"one", "once", "ewe"} or lower.startswith("one-")
            or lower.startswith(("university", "uniform", "unique", "unit", "user", "useful", "usual",
                                 "europe", "euro", "eulog", "euphem", "euphor", "eukary", "euclid", "ubiquit"))):
        return "a"
    if lower.startswith(("herb", "historic", "homage", "sql", "uni", "eu", "one")):
        return None
    return "an" if lower[0] in "aeiou" else "a"


def _grammar_issue(text, match, loc, par_index, suggestion, confidence="A", message=""):
    return Issue("EN_SVA", "error", message or "تطابق مفرد/جمع در این گروه اسمی را بررسی کنید.",
                 loc, text[max(0, match.start() - 15):match.end() + 15], par_index,
                 suggestion, match.group(), category="english_grammar", confidence=confidence,
                 char_start=match.start(), char_end=match.end())


@bounded_findings
def check_english_text(text: str, loc: str = "", par_index: int = -1) -> List[Issue]:
    """بررسی کامل یک قطعه متن انگلیسی با فیلتر دقیق استثناهای علمی."""
    issues: List[Issue] = FindingBuffer()
    if not text or not text.strip():
        return issues

    prot_spans = ProtectedContentDetector.find_spans(text) + quoted_spans(text)

    # ۱. بررسی نقطه‌گذاری
    for rx, msg in PUNCTUATION_RULES:
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            # Both sequences can be grammatical: "had had" and "that that".
            if rx.groups > 0 and len(m.groups()) == 1 and re.fullmatch(r"(?:had\s+had|that\s+that)", m.group(), re.I):
                continue
            issues.append(Issue(
                rule_id="EN_PUNCT",
                category="english_style",
                severity="warn",
                confidence=ConfidenceLevel.HIGH_A.value,
                message=msg,
                loc=loc,
                snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                par_index=par_index,
                char_start=m.start(),
                char_end=m.end(),
                target_text=m.group(0)
            ))

    # Pronunciation is lexicon-based; unknown initials and variable readings abstain.
    for m in re.finditer(r"\b(a|an)\s+([A-Za-z][A-Za-z-]*)\b", text, re.IGNORECASE):
        if ProtectedContentDetector.is_protected(m.start(1), m.end(1), prot_spans):
            continue
        word = m.group(2)
        expected = _expected_article(word)
        if expected and m.group(1).lower() != expected:
            replacement = expected.capitalize() if m.group(1)[0].isupper() else expected
            issues.append(Issue("EN_ARTICLE", "error",
                f"Review the article before {word}: known pronunciation suggests {expected}.",
                loc, m.group(), par_index, replacement, m.group(1),
                category="english_grammar", char_start=m.start(1), char_end=m.end(1)))

    # ۳. بررسی واژگان محاوره‌ای آکادمیک (Informal Academic Tone)
    for pat, (rep, expl) in INFORMAL_ACADEMIC_REPLACEMENTS.items():
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            if _tone_term_is_technical(pat, m, text):
                continue
            issues.append(Issue(
                rule_id="EN_TONE",
                category="english_style",
                severity="info",
                confidence=ConfidenceLevel.REVIEW_B.value,
                message=f"سبک علمی: «{m.group(0)}» — {expl}",
                loc=loc,
                snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                par_index=par_index,
                char_start=m.start(),
                char_end=m.end(),
                target_text=m.group(0),
                suggestion=rep
            ))

    # ۴. بررسی حشو و پرگویی (Wordiness)
    for pat, (rep, expl) in WORDINESS_PATTERNS.items():
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            issues.append(Issue(
                rule_id="EN_WORDINESS",
                category="english_style",
                severity="info",
                confidence=ConfidenceLevel.REVIEW_B.value,
                message=f"ساده‌نویسی علمی: «{m.group(0)}» — {expl}",
                loc=loc,
                snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                par_index=par_index,
                char_start=m.start(),
                char_end=m.end(),
                target_text=m.group(0),
                suggestion=rep
            ))

    # ۵. بررسی حروف اضافه (Prepositions)
    for pat, rep, expl in PREPOSITION_ERRORS:
        if "consist in" in pat or "different than" in pat:
            continue
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            issues.append(Issue(
                rule_id="EN_PREP",
                category="english_style",
                severity="error",
                confidence=ConfidenceLevel.HIGH_A.value,
                message=f"حرف اضافه نادرست: «{m.group(0)}» ← پیشنهاد اصلاح: «{rep}». {expl}",
                loc=loc,
                snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                par_index=par_index,
                char_start=m.start(),
                char_end=m.end(),
                target_text=m.group(0),
                suggestion=rep
            ))

    # Finite local patterns avoid treating "these result from" as a noun error,
    # or singular noun modifiers in "many research studies" as mass nouns.
    noun_forms = {"result": "results", "datum": "data", "finding": "findings",
                  "study": "studies", "method": "methods", "effect": "effects"}
    nominal = r"\b(these|those)\s+(result|datum|finding|study|method|effect)\b(?=\s+(?:indicate|show|suggest|demonstrate|support|are|were|have|is|was|has|can|may)\b|\s*[,.;:!?]|\s*$)"
    for m in re.finditer(nominal, text, re.I):
        if not ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
            suggestion = m.group(1) + " " + noun_forms[m.group(2).lower()]
            issues.append(_grammar_issue(text, m, loc, par_index, suggestion, "B"))

    for m in re.finditer(r"\bmany\s+(research|information|equipment)\b(?=\s+(?:is|was|are|were|has|have|can|could)\b|\s*[,.;:!?]|\s*$)", text, re.I):
        if not ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
            issues.append(_grammar_issue(text, m, loc, par_index, "much " + m.group(1), "B"))

    finite_patterns = [
        (r"\b(?:these|those)\s+(?:results|findings|studies|methods|effects)\s+(is|was|has)\b",
         {"is": "are", "was": "were", "has": "have"}),
        (r"\b(?:this|that)\s+(?:result|finding|study|method|effect)\s+(are|were|have)\b",
         {"are": "is", "were": "was", "have": "has"}),
    ]
    for pattern, replacements in finite_patterns:
        for m in re.finditer(pattern, text, re.I):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            original = m.group(1)
            replacement = replacements[original.lower()]
            if original[0].isupper():
                replacement = replacement.capitalize()
            issues.append(Issue("EN_SVA", "error", "تطابق فعل با فاعل (مفرد/جمع) را بررسی کنید.",
                                loc, m.group(), par_index, replacement, original,
                                category="english_grammar", char_start=m.start(1), char_end=m.end(1)))

    # A stack catches inverted/crossed delimiters that equal counts miss.
    # Quotes, code, citations and numeric intervals already have their own syntax.
    stack, unmatched = [], []
    pairs = {")": "(", "]": "["}
    for index, ch in enumerate(text):
        if ch not in "()[]" or ProtectedContentDetector.is_protected(index, index + 1, prot_spans):
            continue
        if ch in "([":
            stack.append((ch, index))
        elif stack and stack[-1][0] == pairs[ch]:
            stack.pop()
        else:
            unmatched.append(index)
    unmatched.extend(index for _, index in stack)
    for index in sorted(unmatched):
        issues.append(Issue("EN_BRACKET", "warn", "جفت‌بودن یا ترتیب این پرانتز/کروشه را بررسی کنید.",
                            loc, text[max(0, index - 15):index + 16], par_index, target_text=text[index],
                            category="english_style", confidence="B", char_start=index, char_end=index + 1))

    return issues
