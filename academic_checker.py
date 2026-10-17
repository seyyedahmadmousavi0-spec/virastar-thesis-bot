# -*- coding: utf-8 -*-
"""
ماژول ویرایش علمی، سبک دانشگاهی و انسجام متن فارسی (Academic Style & Cohesion Checker).
شامل بررسی لحن علمی، واژگان احساسی و مبالغه‌آمیز، ادعاهای قطعی بدون شاهد،
حشو و عبارات زائد، ضمایر اول شخص، پیوستگی و انسجام پاراگراف‌ها.
"""
import re
from dataclasses import dataclass
from typing import List, Tuple

@dataclass
class AcademicIssue:
    rule: str
    severity: str  # error, warn, info
    message: str
    snippet: str = ""
    loc: str = ""
    suggestion: str = ""
    target_text: str = ""

# ۱. واژه‌ها و عبارات مبالغه‌آمیز، احساسی یا فاقد استناد علمی
EMOTIONAL_EXAGGERATED_PATTERNS = [
    (r"به\s*طرز\s*معجزه‌?آسای?ی", "به شکل قابل توجهی", "در نگارش علمی از تعابیر احساسی/مبالغه‌آمیز مانند «معجزه‌آسا» پرهیز کنید."),
    (r"شاهکار", "دستاورد برجسته / نوآوری", "واژه «شاهکار» لحن ژورنالیستی و احساسی دارد."),
    (r"فوق‌?العاده\s*(عالی|خوب|زیاد|مهم)", "بسیار حائز اهمیت / قابل توجه", "از صفت‌های اغراق‌آمیز پرهیز کنید."),
    (r"بی‌?نظیر|کم‌?نظیر|شگفت‌?انگیز", "کم‌سابقه / قابل توجه", "تعابیر اغراق‌آمیز؛ نتایج را با داده و عدد دقیق توصیف کنید."),
    (r"کاملاً\s*واضح\s*و\s*مبرهن\s*است", "شواهد نشان می‌دهد", "ادعای بدیهی بودن بدون ارائه شواهد علمی مناسب نیست."),
    (r"جای\s*هیچ\s*(گونه\s*)?تردیدی\s*نیست", "داده‌ها قویاً مؤید آن است که", "در علم احکام مطلق وجود ندارد؛ بیان محتاطانه علمی به کار برید."),
    (r"بدون\s*(کوچک‌?ترین\s*)?شک\s*و\s*شبهه", "بر پایه شواهد موجود", "لحن ادعای قطعی در متن علمی مناسب نیست."),
    (r"قطعاً\s*(ثابت|مشخص|مبرهن)\s*شد", "شواهد به روشنی حاکی از آن است", "به جای ادعای قطعیت مطلق از شواهد آماری و تجربی سخن بگویید."),
    (r"صد\s*در\s*صد", "به طور کامل / با ضریب اطمینان بالا", "بیان غیرعلمی و محاوره‌ای."),
    (r"متأسفانه", "با این حال / برخلاف انتظار", "ابراز احساسات شخصی («متأسفانه») در گزارش علمی مجاز نیست."),
    (r"خوشبختانه", "نتایج نشان داد که", "ابراز احساسات شخصی («خوشبختانه») در متن علمی نامناسب است."),
    (r"جای\s*(بسی\s*)?تعجب\s*دارد", "یافته‌ای غیرمنتظره بود", "لحن احساسی؛ از تعبیر «یافته غیرمنتظره» استفاده کنید."),
    (r"نتایج\s*بسیار\s*عالی\s*بود", "نتایج نشان‌دهنده اثربخشی معنی‌دار بود", "عبارت مبهم و احساسی؛ شاخص عددی یا آماری ارائه دهید."),
    (r"\bبه\s*شدت\b(?!\s+(?:نور|جریان|میدان|صوت|تابش|زلزله|درد|بیماری|عفونت|امواج|حرارت|گرما|تحریک|تنش|صدمه|آسیب|صدا|موج|باد|باران|انرژی|برخورد|جوشش|مقاومت|کشش|سختی|جذب|واکنش))\b", "به میزان چشمگیری / به طور معناداری", "«به شدت» در اینجا تعبیری عامیانه است؛ از «به طور چشمگیری» یا «معناداری» استفاده کنید."),
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

# ۳. کلمات ربط و گذار برای ارزیابی انسجام متن (Cohesion markers)
TRANSITION_MARKERS = [
    "بنابراین", "از این رو", "در نتیجه", "با این حال", "با وجود این",
    "علاوه بر این", "افزون بر این", "از سوی دیگر", "در مقابل", "بر این اساس",
    "به عبارت دیگر", "همچنین", "بر همین اساس", "بدین ترتیب", "شایان ذکر است"
]

# ۴. ضمایر شخصی اول شخص در نگارش دانشگاهی
FIRST_PERSON_PATTERNS = [
    (r"(?<![\u0600-\u06FF\u200c])(به\s*نظر\s*من|به\s*عقیده\s*من|من\s*معتقدم)(?![\u0600-\u06FF\u200c])", "به نظر می‌رسد / بر اساس شواهد موجود"),
    (r"(?<![\u0600-\u06FF\u200c])(من|خودم)(?![\u0600-\u06FF\u200c]|\s+(?:جمله|باب|حیث|بعد))", "نگارنده / پژوهشگر / استفاده از ساختار مجهول"),
    (r"(?<![\u0600-\u06FF\u200c])(به\s*نظر\s*ما|ما\s*معتقدیم)(?![\u0600-\u06FF\u200c])", "بر اساس نتایج حاصل"),
]


def check_academic_style(text: str, loc: str = "", in_declaration: bool = False) -> List[AcademicIssue]:
    """بررسی سبک علمی، واژگان احساسی، حشو و ضمایر در یک پاراگراف."""
    issues = []
    if not text or not text.strip():
        return issues

    # ۱. بررسی لحن مبالغه‌آمیز و احساسی
    for pat, rep, expl in EMOTIONAL_EXAGGERATED_PATTERNS:
        rx = re.compile(pat, re.UNICODE)
        for m in rx.finditer(text):
            issues.append(AcademicIssue(
                rule="academic_tone_emotional",
                severity="warn",
                message=f"لحن علمی: عبارت «{m.group(0)}» — {expl} (پیشنهاد: «{rep}»)",
                snippet=m.group(0),
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۲. بررسی حشو و زوائد بیانی
    for pat, rep, expl in PLEONASM_PATTERNS:
        rx = re.compile(pat, re.UNICODE)
        for m in rx.finditer(text):
            issues.append(AcademicIssue(
                rule="academic_pleonasm",
                severity="info",
                message=f"ساده‌نویسی علمی: «{m.group(0)}» ← پیشنهاد اصلاح: «{rep}». {expl}",
                snippet=m.group(0),
                loc=loc,
                suggestion=rep,
                target_text=m.group(0)
            ))

    # ۳. بررسی اول شخص (مگر در صفحه تعهدنامه/سپاسگزاری)
    if not in_declaration:
        for pat, rep in FIRST_PERSON_PATTERNS:
            rx = re.compile(pat, re.UNICODE)
            for m in rx.finditer(text):
                issues.append(AcademicIssue(
                    rule="academic_first_person",
                    severity="info",
                    message=f"کاربرد اول‌شخص در متن علمی: «{m.group(0)}» ← پیشنهاد: {rep}",
                    snippet=m.group(0),
                    loc=loc,
                    suggestion=rep,
                    target_text=m.group(0)
                ))

    return issues


def evaluate_cohesion(paragraphs: List[str]) -> List[AcademicIssue]:
    """ارزیابی پیوستگی و ترنزیشن‌ها در زنجیره پاراگراف‌های یک بخش."""
    issues = []
    long_no_transition_count = 0

    for i, p in enumerate(paragraphs):
        p_clean = p.strip()
        words = p_clean.split()
        if len(words) < 25:
            continue

        # بررسی اینکه آیا پاراگراف با یک حرف ربط یا نشانگر گذار پیوند خورده یا نه
        first_few_words = " ".join(words[:6])
        has_transition = any(tm in first_few_words for tm in TRANSITION_MARKERS)

        if not has_transition:
            long_no_transition_count += 1
            if long_no_transition_count >= 4:
                issues.append(AcademicIssue(
                    rule="academic_cohesion",
                    severity="info",
                    message="چند پاراگراف پیاپی بدون عبارات گذار یا پیونددهنده (Transition Markers) آغاز شده‌اند؛ برای روانی و پیوستگی استدلال از عباراتی نظیر «از سوی دیگر»، «بنابراین»، «با این حال» یا «علاوه بر این» استفاده کنید.",
                    loc=f"پاراگراف {i+1}",
                    snippet=p_clean[:60] + "…"
                ))
                long_no_transition_count = 0
        else:
            long_no_transition_count = 0

    return issues

