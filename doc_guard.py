# -*- coding: utf-8 -*-
"""
ماژول حفاظت از اسناد و تضمین تغییرناپذیری فایل ورودی (Document Guard & Immutability Engine).
اصول بنیادین:
1. فایل اصلی کاربر داده‌ای غیرقابل‌تغییر است و تحت هیچ شرایطی نباید بازنویسی یا ویرایش شود.
2. کنترل یکپارچگی فایل با هش SHA-256 قبل و بعد از پردازش الزامی است.
3. تفکیک کامل مسیرهای ورودی، موقت و خروجی.
4. ایمن‌سازی نام فایل‌ها در برابر حملات تزریق مسیر (Path Traversal).
"""
import hashlib
import os
import re
import shutil
import tempfile
from typing import Optional, Tuple


class DocumentIntegrityError(Exception):
    """خطای تغییر ناخواسته یا آسیب به تمامیت فایل ورودی."""
    pass


def sha256_file(path: str, block_size: int = 65536) -> str:
    """محاسبه هش امن SHA-256 یک فایل به شیوه جریانی."""
    if block_size <= 0:
        raise ValueError("block_size must be positive")
    if not os.path.exists(path):
        raise FileNotFoundError(f"فایل در مسیر {path} یافت نشد.")
    hasher = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(block_size)
            if not chunk:
                break
            hasher.update(chunk)
    return hasher.hexdigest()


def safe_filename(filename: str, default: str = "document") -> str:
    """
    ایمن‌سازی نام فایل ورودی برای جلوگیری از حملات Path Traversal و کاراکترهای مخرب.
    نام پایه استخراج شده و کاراکترهای خطرناک حذف یا جایگزین می‌شوند.
    """
    if not filename:
        return default
    # استخراج نام بدون مسیر
    base = os.path.basename(filename.strip().replace("\\", "/"))
    # حذف کاراکترهای ممنوعه در ویندوز و لینوکس
    cleaned = re.sub(r'[\/:*?"<>|\x00-\x1f\u202a-\u202e\u2066-\u2069]', "_", base).strip(" .")
    # جلوگیری از نام‌های نامتعارف یا نقطه تنها
    if not cleaned or cleaned in (".", ".."):
        cleaned = default
    stem, extension = os.path.splitext(cleaned)
    stem = stem.rstrip(" .")
    if stem.upper() in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9]", stem.upper()):
        stem = "_" + stem
    # Linux limits filename components by UTF-8 bytes, not character count.
    # Leave room for work/annotation prefixes and nested Windows temp paths.
    extension = extension[:16]
    budget = 180 - len(extension.encode("utf-8"))
    stem = stem[:120].encode("utf-8")[:budget].decode("utf-8", errors="ignore").rstrip(" .")
    return (stem or "document") + extension


class DocumentGuard:
    """
    مدیر چرخه حیات فایل و تضمین تغییرناپذیری (Immutability Manager).
    مسیر فایل اصلی را گرفته، هش اولیه را ثبت می‌کند و در پایان صحت آن را بررسی می‌نماید.
    """
    def __init__(self, original_path: str, work_dir: Optional[str] = None):
        self.original_path = os.path.abspath(original_path)
        if not os.path.exists(self.original_path):
            raise FileNotFoundError(f"فایل اصلی یافت نشد: {self.original_path}")

        self.initial_hash = sha256_file(self.original_path)
        self.initial_size = os.path.getsize(self.original_path)
        self._work_dir = work_dir
        self._own_temp = False
        self.temp_dir = ""
        self.isolated_working_copy = ""

    def __enter__(self):
        if self._work_dir:
            os.makedirs(self._work_dir, exist_ok=True)
            self.temp_dir = tempfile.mkdtemp(prefix="docguard_", dir=self._work_dir)
            self._own_temp = True
        else:
            self.temp_dir = tempfile.mkdtemp(prefix="docguard_")
            self._own_temp = True

        # ایجاد یک کپی کاملاً مجزا برای پردازش بدون کوچک‌ترین ارتباط با فایل اصلی
        filename = safe_filename(os.path.basename(self.original_path))
        self.isolated_working_copy = os.path.join(self.temp_dir, f"work_{filename}")
        try:
            shutil.copy2(self.original_path, self.isolated_working_copy)
            if sha256_file(self.isolated_working_copy) != self.initial_hash:
                raise DocumentIntegrityError("Working snapshot does not match input")
        except BaseException:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            raise
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            # بررسی الزامی عدم تغییر فایل اصلی (Zero-Mutation Verification)
            self.verify_original_intact()
        finally:
            if self._own_temp and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)

    def verify_original_intact(self) -> bool:
        """بررسی اکید یکپارچگی فایل اصلی."""
        if not os.path.exists(self.original_path):
            raise DocumentIntegrityError(f"فایل اصلی در طول پردازش مفقود شده است: {self.original_path}")

        current_size = os.path.getsize(self.original_path)
        if current_size != self.initial_size:
            raise DocumentIntegrityError(
                f"حجم فایل اصلی تغییر کرده است! قبل: {self.initial_size} بایت، بعد: {current_size} بایت"
            )

        current_hash = sha256_file(self.original_path)
        if current_hash != self.initial_hash:
            raise DocumentIntegrityError(
                f"هش فایل اصلی تغییر کرده است! تمامیت فایل نقض شده است:\n"
                f"Initial SHA256: {self.initial_hash}\n"
                f"Current SHA256: {current_hash}"
            )
        return True

    def create_output_path(self, prefix: str = "reviewed_") -> str:
        """ایجاد مسیر ایمن و مجزا برای ذخیره فایل خروجی."""
        filename = safe_filename(os.path.basename(self.original_path))
        return os.path.join(self.temp_dir, safe_filename(f"{prefix}{filename}"))



def ensure_distinct_output(source, output):
    """Reject lexical aliases, symlinks and existing hard links before writing."""
    if os.path.normcase(os.path.realpath(source)) == os.path.normcase(os.path.realpath(output)):
        raise DocumentIntegrityError("Output must be independent of input")
    if os.path.exists(output) and os.path.samefile(source, output):
        raise DocumentIntegrityError("Output aliases input")


def validate_document(path, kind, max_bytes=25 * 1024 * 1024):
    """Bound parser input and ZIP expansion before loading any document."""
    import zipfile
    if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
        raise ValueError("Invalid input size budget")
    if not os.path.isfile(path):
        raise ValueError("Input must be a regular file")
    if os.path.getsize(path) > max_bytes:
        raise ValueError("too_big")
    with open(path, "rb") as stream:
        signature = stream.read(8)
    if kind == "pdf":
        if not signature.startswith(b"%PDF-"):
            raise ValueError("Invalid PDF signature")
        return
    if kind != "docx" or not signature.startswith(b"PK"):
        raise ValueError("Invalid DOCX signature")
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        names = [i.filename for i in entries]
        if len({name.casefold() for name in names}) != len(names):
            raise ValueError("Ambiguous ZIP member names")
        if len(entries) > 5000 or len(names) != len(set(names)):
            raise ValueError("Invalid ZIP directory")
        if not {"[Content_Types].xml", "word/document.xml"}.issubset(names):
            raise ValueError("Not a Word package")
        if sum(i.file_size for i in entries) > 256 * 1024 * 1024:
            raise ValueError("ZIP expansion limit")
        for item in entries:
            if item.flag_bits & 1 or item.file_size > 64 * 1024 * 1024:
                raise ValueError("Unsupported ZIP member")
            if item.file_size > max(1, item.compress_size) * 200:
                raise ValueError("ZIP compression ratio limit")
            if item.filename.startswith(("/", "\\")) or ".." in item.filename.replace("\\", "/").split("/"):
                raise ValueError("Unsafe ZIP member")
        for item in entries:
            if item.filename.lower().endswith((".xml", ".rels")):
                with archive.open(item) as stream:
                    tail = b""
                    while True:
                        chunk = stream.read(65536)
                        if not chunk:
                            break
                        reject_unsafe_xml(tail + chunk)
                        tail = chunk[-128:]
        from lxml import etree
        types = etree.fromstring(archive.read("[Content_Types].xml"),
                                etree.XMLParser(resolve_entities=False, no_network=True))
        main_types = [node.get("ContentType") for node in types
                      if node.get("PartName") == "/word/document.xml"]
        if main_types != ["application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"]:
            raise ValueError("Unsupported Word document type")


def immutable_input(function):
    """Check input integrity on success and on parser/rule exceptions."""
    from functools import wraps
    @wraps(function)
    def checked(path, *args, **kwargs):
        before = sha256_file(path)
        try:
            return function(path, *args, **kwargs)
        finally:
            if not os.path.isfile(path) or sha256_file(path) != before:
                raise DocumentIntegrityError("Input integrity verification failed")
    return checked


def reject_unsafe_xml(data):
    """Document XML never needs DTDs or custom entity declarations."""
    # UTF-16/32 XML has zero bytes between ASCII markup characters.
    upper = data.replace(b"\x00", b"").upper()
    if b"<!DOCTYPE" in upper or b"<!ENTITY" in upper:
        raise ValueError("DTD/entity declarations are not supported")


def atomic_write(path, data, *, source_path=None):
    """Publish complete artifacts while preventing source replacement."""
    destination = os.path.abspath(path)
    if source_path:
        ensure_distinct_output(source_path, destination)
    descriptor, temporary = tempfile.mkstemp(prefix=".artifact_", dir=os.path.dirname(destination))
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        if source_path:
            ensure_distinct_output(source_path, destination)
        os.replace(temporary, destination)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
