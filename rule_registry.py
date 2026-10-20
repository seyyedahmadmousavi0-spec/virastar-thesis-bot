"""Versioned rule metadata and early, explicit configuration validation."""
from dataclasses import asdict, dataclass
import math
from functools import wraps


class FindingBudgetExceeded(Exception):
    def __init__(self, findings):
        super().__init__("Finding budget reached")
        self.findings = list(findings)


class FindingBuffer(list):
    """Stop accumulating before a pathological document can exhaust worker memory."""
    def __init__(self, limit=500):
        super().__init__()
        self.limit = limit

    def append(self, value):
        if len(self) >= self.limit:
            raise FindingBudgetExceeded(self)
        super().append(value)

    def extend(self, values):
        for value in values:
            self.append(value)

    def __iadd__(self, values):
        self.extend(values)
        return self


def bounded_findings(function):
    @wraps(function)
    def checked(*args, **kwargs):
        try:
            return function(*args, **kwargs)
        except FindingBudgetExceeded as error:
            from analysis_model import Issue
            return error.findings + [Issue("findings_limit", "warn",
                "تعداد یافته‌های این بخش از بودجهٔ بررسی بیشتر است؛ بررسی دستی لازم است.",
                confidence="B", extraction_status="analysis_limited")]
    return checked


class RuleConfigurationError(ValueError):
    pass


@dataclass(frozen=True)
class RuleSpec:
    rule_id: str
    title: str
    category: str
    confidence_ceiling: str = "B"
    version: str = "2"
    validation: str = "not_independently_validated"


_HIGH = {"arabic_letters", "mi_prefix", "plural_space", "common_mistake", "space_before_punct",
         "no_space_after_punct", "double_space", "EN_ARTICLE", "EN_PREP"}
_TITLES = {
    "arabic_letters": "حروف عربی در متن فارسی", "mi_prefix": "نیم‌فاصلهٔ پیشوند فعلی",
    "plural_space": "فاصلهٔ پسوند جمع", "common_mistake": "صورت نوشتاری قابل بازبینی",
    "space_before_punct": "فاصلهٔ پیش از نشانه", "no_space_after_punct": "فاصلهٔ پس از نشانه",
    "double_space": "فاصلهٔ تکراری", "EN_ARTICLE": "حرف تعریف انگلیسی", "EN_PREP": "حرف اضافهٔ انگلیسی",
    "EN_SVA": "ساختار شمار و تطابق انگلیسی", "EN_PUNCT": "نشانه‌گذاری انگلیسی",
    "EN_WORDINESS": "پیشنهاد ایجاز انگلیسی", "EN_TONE": "پیشنهاد سبک انگلیسی", "EN_BRACKET": "پرانتز انگلیسی",
    "ACAD_EMOTIONAL": "بازبینی لحن علمی", "ACAD_PLEONASM": "پیشنهاد ایجاز فارسی",
    "ACAD_FIRST_PERSON": "بازبینی شیوهٔ بیان نویسنده", "ACAD_COHESION": "بازبینی نشانگرهای پیوند",
    "cite_missing_ref": "بازبینی تطبیق نام و سال", "cite_ambiguous_ref": "چند نامزد برای ارجاع",
    "numeric_cite": "بازبینی ارجاع عددی", "numeric_cite_ambiguous": "بازهٔ عددی نامطمئن",
    "comparative_space": "بازبینی پسوند تفضیل", "ye_space": "بازبینی یای پیوندی",
    "repeated_word": "بازبینی تکرار واژه", "straight_quotes": "پیشنهاد گیومهٔ فارسی",
    "latin_punct": "بازبینی نشانهٔ لاتین", "leading_space": "تورفتگی با فاصله",
    "arabic_digits": "بازبینی شکل رقم عربی", "latin_digits": "بازبینی شکل رقم لاتین",
    "paren_space": "فاصلهٔ پرانتز", "long_sentence": "طول جمله", "long_paragraph": "طول پاراگراف",
    "page_size": "ابعاد صفحه", "page_margin": "حاشیهٔ صفحه", "font_fa": "قلم فارسی",
    "font_en": "قلم انگلیسی", "size_fa": "اندازهٔ قلم فارسی", "size_en": "اندازهٔ قلم انگلیسی",
    "line_spacing": "فاصلهٔ خطوط", "alignment": "تراز پاراگراف", "rtl": "جهت پاراگراف",
    "toc": "فهرست مطالب", "page_numbers": "شمارهٔ صفحه", "chapters": "ساختار فصل‌ها",
    "heading_style": "سبک تیتر", "required_section": "بخش مورد انتظار", "recommended_section": "بخش پیشنهادی",
    "abstract_length": "طول چکیده", "keywords_count": "تعداد کلیدواژه", "caption_missing": "عنوان شکل و جدول",
    "caption_numbering": "شماره‌گذاری عنوان‌ها", "blank_lines": "پاراگراف خالی", "informal_word": "واژگان محاوره‌ای",
    "analysis_limit": "محدودیت بررسی بلوک", "findings_limit": "محدودیت تعداد یافته‌ها",
    "pdf_read_error": "خطای خواندن PDF", "pdf_scanned_warning": "متن استخراج‌نشدهٔ PDF",
    "pdf_page_unreadable": "صفحهٔ فاقد متن قابل‌اعتماد", "pdf_extraction_review": "بازبینی استخراج صفحه",
    "TUMS_CHECKLIST_MISSING": "چک‌لیست دانشگاه", "TUMS_CHAPTER_PROPORTIONS": "نسبت تقریبی فصل‌ها",
    "TUMS_ET_AL_FORBIDDEN": "بازبینی نام نویسندگان منبع", "TUMS_REF_PAGE_FORMAT": "بازبینی صفحهٔ منبع",
    "TUMS_PAGE_MARGIN": "حاشیه‌ها در پروفایل دانشگاه",
}


def rule_spec(rule_id):
    if rule_id.startswith("EN_"):
        category = "english_grammar" if rule_id in {"EN_ARTICLE", "EN_PREP", "EN_SVA"} else "english_style"
    elif rule_id.startswith("ACAD_"):
        category = "academic_style"
    elif rule_id.startswith("TUMS_"):
        category = "university_profile"
    elif "cite" in rule_id or "ref" in rule_id.lower():
        category = "citation_consistency"
    else:
        category = "persian_orthography" if rule_id in _TITLES else "document_review"
    return RuleSpec(rule_id, _TITLES.get(rule_id, rule_id), category,
                    confidence_ceiling="A" if rule_id in _HIGH else "B")


def rule_catalog(extra_ids=()):
    return [asdict(rule_spec(rule_id)) for rule_id in sorted(set(_TITLES) | set(extra_ids))]


def enrich_issue(issue):
    spec = rule_spec(issue.rule_id)
    issue.rule_version = spec.version
    if issue.category in {"general", ""}:
        issue.category = spec.category
    if issue.confidence == "A" and spec.confidence_ceiling != "A":
        issue.confidence = spec.confidence_ceiling
    return issue


def _positive(value, path, *, integer=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise RuleConfigurationError(f"{path} must be a positive finite number")
    if integer and not isinstance(value, int):
        raise RuleConfigurationError(f"{path} must be an integer")


def validate_rules(rules, mode="thesis"):
    if not isinstance(rules, dict):
        raise RuleConfigurationError("Rules must be an object")
    pending = [(rules, 0)]
    count = 0
    while pending:
        value, depth = pending.pop()
        count += 1
        if depth > 50 or count > 50000:
            raise RuleConfigurationError("Configuration exceeds its structure budget")
        if isinstance(value, dict):
            if any(not isinstance(key, str) for key in value):
                raise RuleConfigurationError("Configuration keys must be strings")
            pending.extend((item, depth+1) for item in value.values())
        elif isinstance(value, list):
            pending.extend((item, depth+1) for item in value)
        elif isinstance(value, float) and not math.isfinite(value):
            raise RuleConfigurationError("Configuration numbers must be finite")
        elif value is not None and not isinstance(value, (str, bool, int, float)):
            raise RuleConfigurationError("Configuration must contain JSON-compatible values")
    for name in ("page", "body", "limits", "modes"):
        if not isinstance(rules.get(name), dict):
            raise RuleConfigurationError(f"{name} must be an object")
    if mode not in rules["modes"]:
        raise RuleConfigurationError("Unsupported document mode")
    for profile in rules["modes"].values():
        if not isinstance(profile, dict):
            raise RuleConfigurationError("Each mode must be an object")
        for key in ("required_sections", "recommended_sections"):
            sections = profile.get(key, [])
            if not isinstance(sections, list):
                raise RuleConfigurationError("Mode sections must be arrays")
            for section in sections:
                if not isinstance(section, dict) or not isinstance(section.get("name"), str):
                    raise RuleConfigurationError("Each section needs a name and patterns")
                patterns = section.get("patterns")
                if not isinstance(patterns, list) or any(not isinstance(item, str) or not item for item in patterns):
                    raise RuleConfigurationError("Section patterns must be nonempty strings")
    if not isinstance(rules.get("strict", True), bool):
        raise RuleConfigurationError("strict must be true or false")
    for key in ("protected_terms", "disabled_rules"):
        values = rules.get(key, [])
        if not isinstance(values, list) or any(not isinstance(item, str) or not item.strip() for item in values):
            raise RuleConfigurationError(f"{key} must contain nonempty strings")
        if len(values) > 2000 or any(len(item) > 500 for item in values):
            raise RuleConfigurationError(f"{key} exceeds its configuration budget")
    for key in ("width_cm", "height_cm", "size_tolerance_cm", "margin_tolerance_cm"):
        _positive(rules["page"].get(key), "page." + key)
    margins = rules["page"].get("margins_cm")
    if not isinstance(margins, dict):
        raise RuleConfigurationError("page.margins_cm must be an object")
    for side in ("top", "bottom", "left", "right"):
        _positive(margins.get(side), "page.margins_cm." + side)
    for key in ("fa_fonts", "en_fonts"):
        values = rules["body"].get(key)
        if not isinstance(values, list) or not values or any(not isinstance(v, str) or not v.strip() for v in values):
            raise RuleConfigurationError("body." + key + " must contain font names")
    for key in ("fa_size_pt", "en_size_pt"):
        _positive(rules["body"].get(key), "body." + key)
    for key in ("max_paragraph_words", "max_sentence_words", "min_chapters_thesis"):
        if key in rules["limits"]:
            _positive(rules["limits"][key], "limits." + key, integer=True)
    for key in ("abstract_words", "keywords_count"):
        if key in rules["limits"]:
            bounds = rules["limits"][key]
            if not isinstance(bounds, list) or len(bounds) != 2:
                raise RuleConfigurationError("Limit ranges must contain two values")
            for bound in bounds:
                _positive(bound, "limits." + key, integer=True)
            if bounds[0] > bounds[1]:
                raise RuleConfigurationError("Limit range is reversed")
    for key in ("max_analysis_block_chars", "max_input_bytes", "max_document_blocks", "max_document_chars", "max_pdf_pages", "max_findings", "max_findings_per_block"):
        if key in rules:
            _positive(rules[key], key, integer=True)
    mistakes = rules.get("common_mistakes", {})
    if not isinstance(mistakes, dict):
        raise RuleConfigurationError("common_mistakes must be an object")
    for key, value in mistakes.items():
        if not isinstance(key, str) or not key.strip() or not isinstance(value, str):
            raise RuleConfigurationError("common_mistakes requires nonempty string keys and string suggestions")
    return rules
