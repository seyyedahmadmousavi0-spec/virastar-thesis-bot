# -*- coding: utf-8 -*-
"""
ماژول نمایش تحلیلی، نگاشت آفست و ساختار استاندارد گزارش ایرادها (Analysis Model & Offset Mapper).
اصول بنیادین:
1. هیچ‌گاه آفست رشته نرمال‌شده مستقیماً روی متن اصلی اعمال نمی‌شود مگر با نگاشت دوطرفه و قطعی.
2. هر ایراد دارای شناسه قاعده (rule_id) یکتا، سطح قطعیت (Confidence Level: A, B, C) و دلیل است.
3. تفکیک بین خطای قطعی (A)، پیشنهاد نیازمند بررسی انسانی (B) و هشدار احتمالی/مرزی (C).
4. فیلد auto_applied همیشه False است؛ متن علمی کاربر تغییر خودکار داده نمی‌شود.
"""
from dataclasses import dataclass
from enum import Enum
from typing import List, Tuple
import unicodedata


class ConfidenceLevel(str, Enum):
    """سطوح قطعیت گزارش بر پایه موازین ویرایش علمی و دقت حداکثری."""
    HIGH_A = "A"       # قطعی با شواهد کافی و کمترین احتمال خطای کاذب
    REVIEW_B = "B"     # پیشنهاد ارتقای سبکی یا نیازمند قضاوت پژوهشگر
    AMBIGUOUS_C = "C"  # احتمالی، ساختار نامطمئن یا همپوشان با فرمول/نام خاص


@dataclass
class Issue:
    """
    ساختار استاندارد، جامع و ممیزی‌پذیر برای هر ایراد یا پیشنهاد گزارش‌شده.
    سازگار کامل با امضای متدهای قبلی و افزوده شدن فیلدهای اعتبارسنجی و ممیزی.
    """
    rule_id: str
    severity: str = "warn"          # "error" | "warn" | "info"
    message: str = ""               # شرح به زبان فارسی
    loc: str = ""                   # موقعیت خوانا (مثلاً: «پاراگراف ۴» یا «جدول ۲، سطر ۱»)
    snippet: str = ""               # شاهد از متن اصلی
    par_index: int = -1             # اندیس پاراگراف یا ردیف
    suggestion: str = ""            # پیشنهاد صریح اصلاح
    target_text: str = ""           # متن دقیق عبارت مورد نظر
    category: str = "general"       # دسته‌بندی موضوعی ایراد
    confidence: str = "A"           # "A" | "B" | "C"
    char_start: int = -1            # آفست شروع در متن اصلی
    char_end: int = -1              # آفست پایان در متن اصلی
    protected_overlap: bool = False # آیا با محتوای علمی تداخل داشته است؟
    auto_applied: bool = False      # همواره False: هرگز تغییر خودکار اعمال نمی‌شود

    rule_version: str = "1"
    extraction_status: str = "native"
    location_status: str = "unverified"

    # Appended fields preserve the positional constructor used by older rules.
    block_id: str = ""
    source_part: str = ""
    xml_path: str = ""
    story: str = ""
    paragraph_index: int = -1
    page_index: int = -1

    def __post_init__(self):
        self.auto_applied = False
        if isinstance(self.confidence, ConfidenceLevel):
            self.confidence = self.confidence.value
        if self.confidence not in {"A", "B", "C"}:
            raise ValueError("Issue confidence must be A, B or C")
        if self.severity not in {"error", "warn", "info"}:
            raise ValueError("Issue severity must be error, warn or info")

    def validate_evidence(self, source_text: str) -> bool:
        """Validate offsets against their own source block, never search a substitute.

        Structural findings can have no text target. This method leaves those
        unlocated and cannot certify a claimed exact location without evidence.
        """
        self.auto_applied = False
        exact = (isinstance(self.char_start, int) and not isinstance(self.char_start, bool)
                 and isinstance(self.char_end, int) and not isinstance(self.char_end, bool)
                 and isinstance(self.target_text, str) and bool(self.target_text)
                 and 0 <= self.char_start < self.char_end <= len(source_text)
                 and source_text[self.char_start:self.char_end] == self.target_text)
        self.location_status = "exact" if exact else "unverified"
        if exact:
            self.snippet = source_text[max(0, self.char_start - 18):self.char_end + 18]
        elif self.target_text or self.char_start != -1 or self.char_end != -1:
            self.confidence = ConfidenceLevel.AMBIGUOUS_C.value
            self.char_start = self.char_end = -1
        elif self.confidence == ConfidenceLevel.HIGH_A.value:
            self.confidence = ConfidenceLevel.REVIEW_B.value
        return exact

    def to_dict(self):
        from dataclasses import asdict
        result = asdict(self)
        result["original_text"] = self.target_text if self.location_status == "exact" else None
        result["auto_applied"] = False
        return result

    @property
    def rule(self) -> str:
        """سازگاری با کدهای پیشین که به .rule دسترسی دارند."""
        return self.rule_id


class OffsetMapper:
    """
    نگاشتار مطمئن و دوطرفه بین کاراکترهای متن اصلی و متن تحلیلی/نرمال‌شده.
    اگر در طول نرمال‌سازی حروفی تغییر کنند، حذف شوند یا اضافه شوند،
    موقعیت هر کاراکتر در متن تحلیل دقیقا به موقعیت متناظر در متن اصلی متصل می‌ماند.
    """
    def __init__(self, original_text: str):
        self.original_text = original_text
        self.norm_to_orig: List[int] = []
        self._norm_to_orig_end: List[int] = []
        self.normalized_text = ""
        self._build_default_mapping()

    def _build_default_mapping(self):
        """ایجاد نگاشت پایه (یک به یک)."""
        self.norm_to_orig = list(range(len(self.original_text)))
        self._norm_to_orig_end = list(range(1, len(self.original_text) + 1))
        self.normalized_text = self.original_text

    @classmethod
    def create_normalized(cls, text: str) -> "OffsetMapper":
        """
        ایجاد نسخه تحلیلی استاندارد با حفظ جدول نگاشت کاراکترها:
        - تبدیل ي به ی و ك به ک
        - حفظ و ثبت دقیق جابجایی کاراکترها
        """
        mapper = cls(text)
        norm_chars, starts, ends = [], [], []
        index = 0
        # Normalize Persian letter spelling and Arabic presentation forms only.
        # A general NFKC would silently turn scientific µ, ², subscripts, etc.
        # into different symbols. Combining marks stay with their source base.
        replacements = str.maketrans({"ي": "ی", "ك": "ک", "ە": "ه"})
        while index < len(text):
            end = index + 1
            while end < len(text) and unicodedata.category(text[end]).startswith("M"):
                end += 1
            cluster = text[index:end]
            normalized = "".join(
                unicodedata.normalize("NFKC", ch)
                if 0xFB50 <= ord(ch) <= 0xFDFF or 0xFE70 <= ord(ch) <= 0xFEFF
                else ch for ch in cluster
            ).translate(replacements)
            norm_chars.append(normalized)
            starts.extend([index] * len(normalized))
            ends.extend([end] * len(normalized))
            index = end

        mapper.normalized_text = "".join(norm_chars)
        mapper.norm_to_orig = starts
        mapper._norm_to_orig_end = ends
        return mapper

    def map_slice_to_original(self, norm_start: int, norm_end: int) -> Tuple[int, int, bool]:
        """
        تبدیل یک برش (slice) از متن نرمال‌شده به برش دقیق در متن اصلی.
        خروجی: (orig_start, orig_end, is_reliable)
        """
        if (not isinstance(norm_start, int) or isinstance(norm_start, bool)
                or not isinstance(norm_end, int) or isinstance(norm_end, bool)
                or not self.norm_to_orig):
            return 0, 0, False

        if norm_start < 0 or norm_end > len(self.norm_to_orig):
            return 0, 0, False

        if norm_start >= norm_end:
            return 0, 0, False

        # A match may not select half an expanded ligature or combining cluster.
        if norm_start and self.norm_to_orig[norm_start] == self.norm_to_orig[norm_start - 1]:
            return 0, 0, False
        if norm_end < len(self.norm_to_orig) and self.norm_to_orig[norm_end] == self.norm_to_orig[norm_end - 1]:
            return 0, 0, False

        orig_start = self.norm_to_orig[norm_start]
        # انتهای بازه
        orig_end = self._norm_to_orig_end[norm_end - 1]

        # اعتبارسنجی
        is_reliable = (orig_start < orig_end <= len(self.original_text))
        return orig_start, orig_end, is_reliable

    def get_original_snippet(self, norm_start: int, norm_end: int) -> str:
        """استخراج دقیق متن متناظر از نسخه دست‌نخورده اصلی."""
        s, e, ok = self.map_slice_to_original(norm_start, norm_end)
        if ok:
            return self.original_text[s:e]
        return ""
