# -*- coding: utf-8 -*-
"""
ماژول بررسی و ویرایش متون علمی انگلیسی (English Academic Text Checker).
شامل بررسی گرامر، اسپلینگ، نقطه‌گذاری، حروف اضافه، تطابق فاعل و فعل،
لحن و سبک آکادمیک (Academic Tone)، حشو و عبارات نامناسب.
"""
import re
from dataclasses import dataclass
from typing import List

@dataclass
class EnglishIssue:
    rule: str
    severity: str  # error, warn, info
    message: str
    snippet: str = ""
    loc: str = ""
    suggestion: str = ""
    target_text: str = ""

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
    (r"\bdata\s+is\b", "data are (در نگارش علمی سنتی)", "در بسیاری از شیوه‌نامه‌های پزشکی و علمی data جمع است (data are)."),
    (r"\bmany\s+(research|information|equipment)\b", "اسم غیرقابل شمارش با many", "کلماتی مانند research/information غیرقابل شمارش هستند؛ از much یا a great deal of استفاده کنید."),
]

# خطاهای نقطه‌گذاری انگلیسی
PUNCTUATION_RULES = [
    (re.compile(r"\s+([,.:;?!])"), "فاصله قبل از علامت نگارشی نباید بیاید."),
    (re.compile(r"([,;:?!])(?=[A-Za-z])"), "بعد از علائم نگارشی باید یک فاصله قرار گیرد."),
    (re.compile(r"\b([A-Za-z]+)\s+\1\b", re.IGNORECASE), "تکرار ناخواسته یک کلمه انگلیسی پشت‌سرهم."),
]


def check_english_text(text: str, loc: str = "") -> List[EnglishIssue]:
    """بررسی کامل یک قطعه متن انگلیسی."""
    issues = []
    if not text or not text.strip():
        return issues

    # ۱. بررسی نقطه‌گذاری
    for rx, msg in PUNCTUATION_RULES:
        for m in rx.finditer(text):
            issues.append(EnglishIssue(
                rule="en_punctuation",
                severity="warn",
                message=msg,
                snippet=text[max(0, m.start()-15):min(len(text), m.end()+15)],
                loc=loc,
                target_text=m.group(0)
            ))

    # ۲. بررسی حروف تعریف (Articles)
    for rx, correct_art, msg in ARTICLE_ERRORS:
        for m in rx.finditer(text):
            # استثناهای فونتیک و اختصارات علمی زبان انگلیسی
            word = m.group(2).lower()
            a_exceptions = (
                "university", "uniform", "unique", "european", "one", "user", "unit",
                "universal", "usage", "useful", "useless", "usual", "ubiquitous",
                "unilateral", "utopia", "eukaryotic", "eukaryote", "one-way", "once",
                "ufo", "historic", "historical"
            )
            an_exceptions = (
                "hour", "honest", "honor", "honorable", "honorary", "heir", "heiress",
                "mri", "rna", "dna", "x-ray", "fda", "who", "nih", "icu", "ct", "pcr",
                "hiv", "sd", "snp", "snps", "html", "sql", "suv", "sos"
            )
            if (correct_art == "an" and (word in a_exceptions or word.startswith("uni-") or word.startswith("euro-"))) or \
               (correct_art == "a" and (word in an_exceptions or word.startswith("hour-"))):
                continue
            issues.append(EnglishIssue(
                rule="en_article",
                severity="error",
                message=f"{msg} (اصلاح به «{correct_art} {m.group(2)}»)",
                snippet=m.group(0),
                loc=loc,
                suggestion=f"{correct_art} {m.group(2)}",
                target_text=m.group(0)
            ))

    # ۳. بررسی واژگان محاوره‌ای آکادمیک (Informal Academic Tone)
    for pat, (rep, expl) in INFORMAL_ACADEMIC_REPLACEMENTS.items():
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            issues.append(EnglishIssue(
                rule="en_academic_tone",
                severity="info",
                message=f"سبک علمی: «{m.group(0)}» ← پیشنهاد: {rep}. {expl}",
                snippet=text[max(0, m.start()-15):min(len(text), m.end()+15)],
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۴. بررسی حشو و پرگویی (Wordiness)
    for pat, (rep, expl) in WORDINESS_PATTERNS.items():
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            issues.append(EnglishIssue(
                rule="en_wordiness",
                severity="info",
                message=f"ساده‌نویسی علمی: «{m.group(0)}» ← {rep}. {expl}",
                snippet=text[max(0, m.start()-15):min(len(text), m.end()+15)],
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۵. بررسی حروف اضافه (Prepositions)
    for pat, rep, expl in PREPOSITION_ERRORS:
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            issues.append(EnglishIssue(
                rule="en_preposition",
                severity="error",
                message=f"کاربرد حرف اضافه: «{m.group(0)}» ← اصلاح: «{rep}». {expl}",
                snippet=m.group(0),
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۶. بررسی تطابق فاعل و فعل / شمارش اسامی (SVA / Plurality)
    for pat, rep, expl in SVA_ERRORS:
        rx = re.compile(pat, re.IGNORECASE)
        for m in rx.finditer(text):
            issues.append(EnglishIssue(
                rule="en_grammar_agreement",
                severity="warn",
                message=f"تطابق دستوری: «{m.group(0)}» — {expl} (پیشنهاد: {rep})",
                snippet=m.group(0),
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۷. بررسی پرانتز و گیومه‌های جفت‌نشده در انگلیسی
    open_p = text.count("(")
    close_p = text.count(")")
    if open_p != close_p:
        issues.append(EnglishIssue(
            rule="en_unpaired_brackets",
            severity="warn",
            message=f"عدم تقارن در پرانتزها: {open_p} پرانتز باز و {close_p} پرانتز بسته یافت شد.",
            loc=loc
        ))

    return issues

