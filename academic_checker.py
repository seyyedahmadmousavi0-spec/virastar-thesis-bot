# -*- coding: utf-8 -*-
"""
ماژول ویرایش علمی، سبک دانشگاهی و انسجام متن فارسی (Academic Style & Cohesion Checker).
شامل:
- بررسی لحن علمی، واژگان احساسی و مبالغه‌آمیز (ACAD_EMOTIONAL)
- ادعاهای قطعی بدون شاهد
- حشو و عبارات زائد (ACAD_PLEONASM)
- ضمایر اول‌شخص در نگارش دانشگاهی (ACAD_FIRST_PERSON)
- پیوستگی و انسجام پاراگراف‌ها (ACAD_COHESION)
تمامی پیشنهادات این بخش در سطح قطعیت B (نیازمند قضاوت پژوهشگر) ارائه می‌شوند.
"""
import re
from typing import List, Tuple

from analysis_model import ConfidenceLevel, Issue
from rule_registry import FindingBuffer, bounded_findings
from protected_content import ProtectedContentDetector, quoted_spans

FA = "آ-ی\u0621-\u063A\u0641-\u064A\u067E\u0686\u0698\u06A9\u06AF\u06CC\u06C0\u200c"

# ۱. واژه‌ها و عبارات مبالغه‌آمیز، احساسی یا فاقد استناد علمی
EMOTIONAL_EXAGGERATED_PATTERNS = [
    (r"به\s*طرز\s*معجزه‌?آسای?ی", "به شکل قابل توجهی", "در نگارش علمی از تعابیر احساسی/مبالغه‌آمیز مانند «معجزه‌آسا» پرهیز کنید."),
    (r"شاهکار", "دستاورد برجسته / نوآوری", "واژه «شاهکار» لحن ژورنالیستی و احساسی دارد."),
    (r"فوق‌?العاده\s*(?:عالی|خوب|زیاد|مهم)", "بسیار حائز اهمیت / قابل توجه", "از صفت‌های اغراق‌آمیز پرهیز کنید."),
    (r"بی‌?نظیر|کم‌?نظیر|شگفت‌?انگیز", "کم‌سابقه / قابل توجه", "تعابیر اغراق‌آمیز؛ نتایج را با داده و عدد دقیق توصیف کنید."),
    (r"کاملاً\s*واضح\s*و\s*مبرهن\s*است", "شواهد نشان می‌دهد", "ادعای بدیهی بودن بدون ارائه شواهد علمی مناسب نیست."),
    (r"جای\s*هیچ\s*(?:گونه\s*)?تردیدی\s*نیست", "داده‌ها مؤید آن است که", "در علم احکام مطلق وجود ندارد؛ بیان محتاطانه علمی به کار برید."),
    (r"بدون\s*(?:کوچک‌?ترین\s*)?شک\s*و\s*شبهه", "بر پایه شواهد موجود", "لحن ادعای قطعی در متن علمی مناسب نیست."),
    (r"قطعاً\s*(?:ثابت|مشخص|مبرهن)\s*شد", "شواهد به روشنی حاکی از آن است", "به جای ادعای قطعیت مطلق از شواهد آماری و تجربی سخن بگویید."),
    (r"صد\s*در\s*صد", "به طور کامل / با ضریب اطمینان بالا", "بیان غیرعلمی و عامیانه."),
    (r"متأسفانه", "با این حال / برخلاف انتظار", "ابراز احساسات شخصی («متأسفانه») در گزارش علمی مجاز نیست."),
    (r"خوشبختانه", "نتایج نشان داد که", "ابراز احساسات شخصی («خوشبختانه») در متن علمی نامناسب است."),
    (r"جای\s*(?:بسی\s*)?تعجب\s*دارد", "یافته‌ای غیرمنتظره بود", "لحن احساسی؛ از تعبیر «یافته غیرمنتظره» استفاده کنید."),
    (r"نتایج\s*بسیار\s*عالی\s*بود", "نتایج نشان‌دهنده اثربخشی معنی‌دار بود", "عبارت مبهم و احساسی؛ شاخص عددی یا آماری ارائه دهید."),
    (r"(?<![" + FA + r"])به\s*شدت(?!\s+(?:نور|جریان|میدان|صوت|تابش|زلزله|درد|بیماری|عفونت|امواج|حرارت|گرما|تحریک|تنش|صدمه|آسیب|صدا|موج|باد|باران|انرژی|برخورد|جوشش|مقاومت|کشش|سختی|جذب|واکنش))(?!" + FA + r")",
     "به میزان چشمگیری / به طور معناداری", "«به شدت» در اینجا تعبیری عامیانه است؛ از «به طور چشمگیری» یا «معناداری» استفاده کنید."),
]

# ۲. عبارات حشو و پرگو در فارسی (Wordiness / Pleonasm)
PLEONASM_PATTERNS = [
    (r"به\s*طور\s*مداوم\s*و\s*پیوسته", "به‌طور پیوسته", "حشو (دو واژه مترادف)."),
    (r"سیر\s*صعودی\s*رو\s*به\s*بالا", "سیر صعودی", "حشو قبیح (صعودی خود به معنی رو به بالا است)."),
    (r"سیر\s*نزولی\s*رو\s*به\s*پایین", "سیر نزولی", "حشو (نزولی یعنی رو به پایین)."),
    (r"به\s*عقیده\s*و\s*نظر", "به نظر / بر پایه دیدگاه", "حشو مترادف."),
    (r"همان‌?طوری\s*که\s*قبلاً\s*ذکر\s*شد\s*و\s*بیان\s*گردید", "چنان‌که پیش‌تر اشاره شد", "اطناب و پرگویی غیرضروری."),
    (r"در\s*راستای\s*این\s*موضوع", "در این راستا", "عبارت زائد."),
    (r"لازم\s*به\s*ذکر\s*است\s*که\s*باید\s*گفت", "شایان ذکر است که", "حشو و پرگویی."),
    (r"به\s*دلیل\s*اینکه\s*چون", "زیرا / از آنجا که", "تکرار دو حرف ربط سببیت."),
    (r"پیرامون\s*حول\s*و\s*حوش", "درباره / در خصوص", "ترکیب عامیانه و حشو."),
]

# ۳. کلمات ربط و نشانگرهای گذار برای ارزیابی انسجام متن (Cohesion markers)
TRANSITION_MARKERS = [
    "بنابراین", "از این رو", "در نتیجه", "با این حال", "با وجود این",
    "علاوه بر این", "افزون بر این", "از سوی دیگر", "در مقابل", "بر این اساس",
    "به عبارت دیگر", "همچنین", "بر همین اساس", "بدین ترتیب", "شایان ذکر است",
    "نکته قابل توجه اینکه", "در همین راستا", "با توجه به نتایج"
]

# ۴. ضمایر شخصی اول شخص در نگارش دانشگاهی
FIRST_PERSON_PATTERNS = [
    (r"(?<![" + FA + r"])(به\s*نظر\s*من|به\s*عقیده\s*من|من\s*معتقدم)(?![" + FA + r"])", "به نظر می‌رسد / بر اساس شواهد موجود"),
    (r"(?<![" + FA + r"])(من|خودم)(?![" + FA + r"]|\s+(?:جمله|باب|حیث|بعد))", "نگارنده / پژوهشگر / استفاده از ساختار مجهول"),
    (r"(?<![" + FA + r"])(به\s*نظر\s*ما|ما\s*معتقدیم)(?![" + FA + r"])", "بر اساس نتایج حاصل"),
]


@bounded_findings
def check_academic_style(text: str, loc: str = "", in_declaration: bool = False, par_index: int = -1) -> List[Issue]:
    """بررسی سبک علمی، واژگان احساسی، حشو و ضمایر در یک پاراگراف."""
    issues: List[Issue] = FindingBuffer()
    if not text or not text.strip():
        return issues

    prot_spans = ProtectedContentDetector.find_spans(text) + quoted_spans(text)

    # ۱. بررسی لحن مبالغه‌آمیز و احساسی
    for pat, rep, expl in EMOTIONAL_EXAGGERATED_PATTERNS:
        if "معنادار" in rep or "معنی‌دار" in rep:
            rep = "توصیف دقیق‌تر نتیجه با اتکا به داده‌ها و شواهد موجود"
        rx = re.compile(pat, re.UNICODE)
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            issues.append(Issue(
                rule_id="ACAD_EMOTIONAL",
                category="academic_style",
                severity="warn",
                confidence=ConfidenceLevel.REVIEW_B.value,
                message=f"لحن علمی: عبارت «{m.group(0)}» — {expl}",
                loc=loc,
                snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                par_index=par_index,
                char_start=m.start(),
                char_end=m.end(),
                target_text=m.group(0),
                suggestion=rep
            ))

    # ۲. بررسی حشو و زوائد بیانی
    for pat, rep, expl in PLEONASM_PATTERNS:
        rx = re.compile(pat, re.UNICODE)
        for m in rx.finditer(text):
            if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                continue
            issues.append(Issue(
                rule_id="ACAD_PLEONASM",
                category="academic_style",
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

    # ۳. بررسی ضمیر اول شخص (به جز در صفحه تعهدنامه)
    if not in_declaration:
        for pat, rep in FIRST_PERSON_PATTERNS:
            rx = re.compile(pat, re.UNICODE)
            for m in rx.finditer(text):
                if ProtectedContentDetector.is_protected(m.start(), m.end(), prot_spans):
                    continue
                issues.append(Issue(
                    rule_id="ACAD_FIRST_PERSON",
                    category="academic_style",
                    severity="warn",
                    confidence=ConfidenceLevel.REVIEW_B.value,
                    message=f"کاربرد ضمیر اول‌شخص «{m.group(0)}» در نگارش علمی دانشگاهی؛ پیشنهاد: ساختار مجهول یا «پژوهشگر/نگارنده».",
                    loc=loc,
                    snippet=text[max(0, m.start() - 15):min(len(text), m.end() + 15)],
                    par_index=par_index,
                    char_start=m.start(),
                    char_end=m.end(),
                    target_text=m.group(0),
                    suggestion=rep
                ))

    # Phrase-level and pronoun-level patterns express the same review concern.
    first_person = sorted((issue for issue in issues if issue.rule_id == "ACAD_FIRST_PERSON"),
                          key=lambda issue: -(issue.char_end - issue.char_start))
    kept = []
    for issue in first_person:
        if not any(other.char_start <= issue.char_start and issue.char_end <= other.char_end for other in kept):
            kept.append(issue)
    return [issue for issue in issues if issue.rule_id != "ACAD_FIRST_PERSON" or issue in kept]


def evaluate_cohesion(paragraphs: List[str], enabled: bool = False) -> List[Issue]:
    """ارزیابی انسجام کلان و پیوستگی استدلال میان پاراگراف‌ها."""
    issues: List[Issue] = []
    if not enabled or len(paragraphs) < 3:
        return issues

    long_runs_without_marker = 0
    start_par = -1

    for idx, p in enumerate(paragraphs):
        p_str = p.strip()
        # پاراگراف‌های کوتاه یا عناوین را ارزیابی نمی‌کنیم
        if len(p_str.split()) < 35:
            long_runs_without_marker = 0
            continue

        has_marker = any(m in p_str for m in TRANSITION_MARKERS)
        if not has_marker:
            if long_runs_without_marker == 0:
                start_par = idx + 1
            long_runs_without_marker += 1
        else:
            long_runs_without_marker = 0

        # اگر بیش از ۳ پاراگراف متوالی بلند هیچ نشانگر گذاری نداشته باشند
        if long_runs_without_marker >= 4:
            issues.append(Issue(
                rule_id="ACAD_COHESION",
                category="academic_style",
                severity="info",
                confidence=ConfidenceLevel.AMBIGUOUS_C.value,
                message=(
                    f"بررسی دستی پیوند متن: از پاراگراف {start_par} تا {idx + 1}، "
                    "چند پاراگراف متوالی بدون کلمات ربط و گذار (مانند بنابراین، با این حال، در نتیجه) نگاشته شده‌اند. "
                    "استفاده از نشانگرهای گذار به درک پیوستگی منطقی استدلال کمک می‌کند."
                ),
                loc=f"پاراگراف {start_par} تا {idx + 1}",
                snippet=p_str[:80] + "…",
                par_index=idx,
                suggestion="استفاده از کلمات ربط (بنابراین، از سوی دیگر، در نتیجه...)"
            ))
            long_runs_without_marker = 0

    return issues
