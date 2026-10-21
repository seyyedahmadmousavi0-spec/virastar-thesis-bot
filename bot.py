# -*- coding: utf-8 -*-
"""
بازوی بله «ویراستار پژوهش» برای ارزیابی، ویرایش علمی و بررسی پایان‌نامه و مقاله.
پشتیبانی از فایل‌های Word (.docx)، فایل‌های PDF و متون مستقیم ارسالی در چت.
مجهز به پروفایل تخصصی دانشگاه علوم پزشکی تهران (TUMS) و بررسی متون انگلیسی.
"""
import logging
import os
import shutil
import signal
import socket
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote

import requests

import checker
from doc_guard import DocumentGuard, safe_filename, validate_document
from bot_runtime import JobGate, ThreadLocalHTTP, env_integer, validated_base_url
from report_artifacts import PRODUCT_VERSION, write_artifacts
from report_store import ReportStore

def _load_env_file():
    """اگر فایل .env کنار bot.py باشد، متغیرها را از آن می‌خواند (BALE_TOKEN=...)."""
    if os.environ.get("BALE_SKIP_DOTENV") == "1":
        return
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env_file()

TOKEN = os.environ.get("BALE_TOKEN", "").strip().strip('"').strip("'")
API = os.environ.get("BALE_API_BASE", "https://tapi.bale.ai").rstrip("/")
FILE_API = os.environ.get("BALE_FILE_BASE", f"{API}/file").rstrip("/")
_config_errors = []


def _integer(name, default, minimum, maximum):
    value, error = env_integer(os.environ, name, default, minimum, maximum)
    if error:
        _config_errors.append(error)
    return value


MAX_MB = _integer("MAX_FILE_MB", 25, 1, 100)
COOLDOWN = _integer("COOLDOWN_SEC", 10, 0, 3600)
WORKERS = _integer("WORKERS", 4, 1, 32)
MAX_QUEUE = _integer("MAX_QUEUE", 8, 0, 1000)
MAX_TEXT_CHARS = _integer("MAX_TEXT_CHARS", 50000, 100, 200000)
HEALTH_STALE_SEC = _integer("HEALTH_STALE_SEC", 120, 45, 3600)
MAX_COMMENTS = _integer("MAX_COMMENTS", 250, 1, 1000)
REPORT_TTL_MINUTES = _integer("REPORT_TTL_MINUTES", 15, 1, 60)
REPORT_CACHE_MB = _integer("REPORT_CACHE_MB", 32, 1, 128)
PDF_ANNOTATION_EXPERIMENTAL = os.environ.get("PDF_ANNOTATION_EXPERIMENTAL", "0") == "1"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("bot")

_bot_status = {
    "started": False,
    "start_time": None,
    "username": None,
    "last_poll_time": None,
    "last_poll_count": 0,
    "last_poll_error": None,
    "total_polls": 0,
    "last_message": None,
    "last_reply": None
}

modes = {}        # chat_id -> "tums" | "thesis" | "article" | "english"
document_scopes = {}  # Explicit full-document / fragment scope for Kashan reviews.
jobs = JobGate(WORKERS + MAX_QUEUE, COOLDOWN)
busy, last_job, lock = jobs.busy, jobs.last_started, jobs.lock
http = ThreadLocalHTTP(requests.Session)
_stop_event = threading.Event()
_last_poll_monotonic = None
_health_server = None
reports = ReportStore(ttl_seconds=REPORT_TTL_MINUTES * 60,
                      max_bytes=REPORT_CACHE_MB * 1024 * 1024)

WELCOME = (
    "سلام 👋\n"
    "به سامانه بازوی «ویراستار پژوهش» خوش آمدید.\n\n"
    "این سامانه برای بازبینی قاعده‌محور پایان‌نامه‌ها، رساله‌ها و مقالات طراحی شده است. یافته‌ها پیشنهاد بازبینی‌اند و صحت علمی یا نبود خطا را تضمین نمی‌کنند:\n"
    "• املای دقیق و تایپوگرافی فارسی (نیم‌فاصله، ی/ک، ارقام، پیشوندها و پسوندها)\n"
    "• سبک و لحن علمی (حذف عبارات مبالغه‌آمیز، محاوره‌ای و ادعاهای قطعی بدون شاهد)\n"
    "• دستور زبان و ساختار متون فارسی و انگلیسی (Grammar, Articles, Prepositions)\n"
    "• بررسی برخی نشانه‌های ساختاری و بخش‌های سند\n"
    "• ارجاع‌دهی درون‌متنی و تطبیق با فهرست منابع (ونکوور، هاروارد و APA)\n"
    "• صفحه‌آرایی فایل Word (حاشیه‌ها، فونت‌ها، تراز و شماره‌گذاری)\n"
    "• دریافت کپی Word با هایلایت و کامنت‌های کوتاه؛ پیشنهادها خودکار اعمال نمی‌شوند\n\n"
    "📌 انتخاب نوع سند:\n"
    "🎓 /thesis : پایان‌نامه و رساله دانشگاهی (پیش‌فرض)\n"
    "🏛 /tums : بررسی با پروفایل دانشگاه علوم پزشکی تهران\n"
    "🏛 /kashan : دانشگاه کاشان؛ انتخاب ارشد یا دکتری\n"
    "📑 /article : مقاله پژوهشی / علمی\n"
    "🇬🇧 /english : بررسی متون و مقالات انگلیسی\n"
    "ℹ️ /status : مشاهده وضعیت جاری\n"
    "📖 /help : راهنمای کامل\n\n"
    "📎 فایل Word (.docx) یا PDF بفرستید یا متن را در چت بنویسید.\n"
    "📋 /report گزارش کامل آخرین بررسی شما را می‌فرستد؛ /comments راهنمای کامنت‌هاست."
)

HELP = (
    "📖 راهنمای کار با ربات ویراستار پژوهش:\n\n"
    "۱) انتخاب نوع سند:\n"
    "• دستور /thesis برای پایان‌نامه‌ها و رساله‌ها (پیش‌فرض).\n"
    "• دستور /tums برای پروفایل دانشگاه علوم پزشکی تهران؛ انطباق قطعی نیازمند بازبینی شیوه‌نامه است.\n"
    "• دکمهٔ «دانشگاه کاشان» یا /kashan: انتخاب کارشناسی ارشد یا دکتری و سند کامل یا بخشی از آن.\n"
    "  انتخاب مستقیم: /kashan_master و /kashan_phd؛ /kashan_info مرجع و محدودیت‌ها را نشان می‌دهد.\n"
    "• دستور /article برای مقالات پژوهشی و علمی.\n"
    "• دستور /english برای بازبینی قاعده‌محور متون انگلیسی.\n\n"
    "۲) ارسال ورودی:\n"
    "• فایل Word (.docx): یک کپی با هایلایت و کامنت‌های کوتاه دریافت می‌کنید؛ دستور /annotate لازم نیست. فایل اصلی و متن تغییر نمی‌کنند.\n"
    "• فایل PDF: خلاصه و گزارش صفحه‌دار؛ برای صفحات تصویری استخراج متن انجام نمی‌شود. نشانه‌گذاری PDF فقط در حالت آزمایشی میزبان و برای محل‌های قابل تطبیق انجام می‌شود.\n"
    "• متن مستقیم: یک یا چند بند از متن خود را بفرستید تا فوراً بررسی گردد.\n\n"
    "۳) /report گزارش HTML و TXT آخرین بررسی شما، /report html یا /report txt فقط یک قالب را می‌فرستد.\n"
    f"گزارش تا {REPORT_TTL_MINUTES} دقیقه در حافظه می‌ماند؛ بررسی جدید جای قبلی را می‌گیرد و راه‌اندازی مجدد یا پرشدن ظرفیت ممکن است آن را زودتر پاک کند.\n"
    "۴) /comments روش دیدن کامنت‌ها، /forget حذف گزارش و /settings تنظیمات را نشان می‌دهد.\n"
    "🔒 فایل‌های موقت این پردازش پاک‌سازی می‌شوند؛ نگهداری داده در بله و میزبان تابع سیاست‌های آن‌هاست.\n"
    "نتایج پیشنهاد بازبینی‌اند؛ بررسی علمی محتوا، صحت منابع و استخراج PDF کامل یا قطعی نیست."
)

MAIN_KEYBOARD = {
    "keyboard": [
        [{"text": "🎓 پایان‌نامه"}, {"text": "📑 مقاله پژوهشی"}],
        [{"text": "🏛 دانشگاه کاشان"}],
        [{"text": "🇬🇧 متن انگلیسی"}, {"text": "ℹ️ وضعیت جاری"}],
        [{"text": "📋 گزارش کامل"}, {"text": "💬 راهنمای کامنت‌ها"}],
        [{"text": "📖 راهنما"}, {"text": "🗑 حذف گزارش"}]
    ],
    "resize_keyboard": True
}

KASHAN_MODE_TITLES = {
    "kashan": "دانشگاه کاشان؛ مقطع هنوز انتخاب نشده",
    "kashan_master": "دانشگاه کاشان — کارشناسی ارشد",
    "kashan_phd": "دانشگاه کاشان — دکتری",
}
KASHAN_KEYBOARD = {
    "keyboard": [
        [{"text": "🎓 کاشان — کارشناسی ارشد"}, {"text": "🎓 کاشان — دکتری"}],
        [{"text": "📄 کاشان — سند کامل"}, {"text": "📑 کاشان — بخشی از سند"}],
        [{"text": "ℹ️ راهنمای کاشان"}, {"text": "⬅️ منوی اصلی"}],
    ],
    "resize_keyboard": True,
}
KASHAN_GUIDANCE = (
    "🏛 پروفایل دانشگاه کاشان\n"
    "مبنای این پروفایل، خلاصهٔ راهنماهای ارشد و دکتری بهمن ۱۳۹۸ در مشخصات ارائه‌شدهٔ مالک ربات است؛ "
    "اعتبار آخرین نسخه و الزامات رشته باید با تحصیلات تکمیلی تطبیق داده شود.\n"
    "ساختار و نشانه‌های قابل استخراج، سازگاری مقطع و منابع با رویکرد APA بررسی می‌شوند؛ نسخهٔ APA حدس زده نمی‌شود. "
    "تقدیم اختیاری است و نمونهٔ سه فصل، تعداد اجباری فصل‌ها محسوب نمی‌شود.\n"
    "قلم ۱۲، حاشیهٔ ۳ سانتی‌متر و سایر قواعد راهنمای سه‌صفحه‌ای با دانشگاه نامشخص، الزام کاشان تلقی نمی‌شوند. "
    "امضا، مهر، صفحه‌بندی واقعی، اعتبار علمی و الزامات روز ایرانداک به بررسی انسانی یا مدرک معتبر نیاز دارند.\n"
    "برای یک فصل یا چکیده، گزینهٔ «بخشی از سند» یا /scope part را انتخاب کنید تا نبود صفحات یک پایان‌نامهٔ کامل ایراد محسوب نشود. "
    "بازگشت به سند کامل: /scope full. فایل اصلی بازنویسی نمی‌شود."
)


# ---------------------------- API بله ----------------------------
class BaleAPIError(RuntimeError):
    """A sanitized failure: never include response bodies or token-bearing URLs."""


def _api_result(response, method):
    try:
        if not 200 <= response.status_code < 300:
            raise BaleAPIError(f"{method}: HTTP {response.status_code}")
        data = response.json()
        if not isinstance(data, dict) or data.get("ok") is not True or "result" not in data:
            raise BaleAPIError(f"{method}: API rejected or malformed response")
        return data["result"]
    except (ValueError, TypeError) as exc:
        raise BaleAPIError(f"{method}: invalid JSON response") from exc


def call(method, **payload):
    try:
        with http.post(f"{API}/bot{TOKEN}/{method}", json=payload,
                       timeout=(10, 45), allow_redirects=False) as response:
            return _api_result(response, method)
    except requests.RequestException as exc:
        raise BaleAPIError(f"{method}: transport failure") from exc


def send(chat_id, text, reply_markup=MAIN_KEYBOARD, receipt=None):
    parts = checker.split_message(text, 3800)
    for i, part in enumerate(parts):
        payload = {"chat_id": chat_id, "text": part}
        if i == len(parts) - 1 and reply_markup is not None:
            payload["reply_markup"] = reply_markup
        try:
            result = call("sendMessage", **payload)
            if receipt is not None:
                receipt.append(result)
            _bot_status["last_reply"] = {"time": time.strftime("%Y-%m-%d %H:%M:%S")}
        except Exception as e:
            log.warning("sendMessage failed (%s)", type(e).__name__)
            return False
    return True


def send_file(chat_id, path, caption="", reply_markup=MAIN_KEYBOARD):
    data = {"chat_id": chat_id, "caption": caption}
    if reply_markup is not None:
        import json
        data["reply_markup"] = json.dumps(reply_markup)
    try:
        with open(path, "rb") as f:
            with http.post(f"{API}/bot{TOKEN}/sendDocument", data=data,
                           files={"document": (os.path.basename(path), f)},
                           timeout=(10, 120), allow_redirects=False) as response:
                return _api_result(response, "sendDocument")
    except requests.RequestException as exc:
        raise BaleAPIError("sendDocument: transport failure") from exc


def download(file_id, dest):
    info = call("getFile", file_id=file_id)
    file_path = info.get("file_path") if isinstance(info, dict) else None
    if not isinstance(file_path, str) or not file_path or file_path.startswith(("/", "\\")) or "\\" in file_path or any(part in {".", ".."} for part in file_path.split("/")):
        raise BaleAPIError("getFile: invalid file path")
    url = f"{FILE_API}/bot{TOKEN}/{quote(file_path, safe='/')}"
    try:
        with http.get(url, stream=True, timeout=(10, 120), allow_redirects=False) as r:
            if r.status_code != 200:
                raise BaleAPIError(f"download: HTTP {r.status_code}")
            size = 0
            with open(dest, "wb") as f:
                for chunk in r.iter_content(65536):
                    size += len(chunk)
                    if size > MAX_MB * 1024 * 1024:
                        raise ValueError("too_big")
                    f.write(chunk)
    except requests.RequestException:
        raise BaleAPIError("download: transport failure") from None


# ---------------------------- پردازش اسناد ----------------------------
class Progress:
    """One editable status message; a failed edit never loses the final result."""
    def __init__(self, chat_id, text="⏳ فایل شما در حال بررسی است…"):
        self.chat_id = chat_id
        receipt = []
        send(chat_id, text, reply_markup=None, receipt=receipt)
        result = receipt[-1] if receipt else None
        self.message_id = result.get("message_id") if isinstance(result, dict) else None

    def update(self, text, final=False):
        if self.message_id is not None:
            try:
                call("editMessageText", chat_id=self.chat_id,
                     message_id=self.message_id, text=text)
                return
            except BaleAPIError:
                pass
        if final and not send(self.chat_id, text):
            raise BaleAPIError("Could not deliver summary")


COMMENTS_HELP = (
    "💬 فایل «نسخهٔ بازبینی» را در Word باز کنید و از Review ← Show Comments کامنت‌ها را ببینید. "
    "در نسخهٔ موبایل، نمایش کامنت‌ها معمولاً از منوی بازبینی یا علامت گفت‌وگو فعال می‌شود.\n"
    "هایلایت، محل پیشنهاد است؛ داخل کامنت شناسهٔ قاعده، دلیل و پیشنهاد کوتاه آمده است.\n"
    "A: اطمینان بالا · B: نیازمند بازبینی · C: مبهم و فقط برای بررسی دستی در گزارش. "
    "این حروف شدت خطا یا درصد دقت نیستند.\n"
    "هیچ پیشنهادی خودکار اعمال نشده است. گزارش کامل و علت درج‌نشدن بعضی کامنت‌ها: /report"
)


def _cache_reports(chat_id, sender_id, artifacts):
    return reports.put(chat_id, sender_id,
                       {kind: artifacts[kind].read_bytes() for kind in ("html", "txt")})


def deliver_reports(chat_id, sender_id=None, kind="all"):
    sender_id = chat_id if sender_id is None else sender_id
    payloads = reports.get(chat_id, sender_id)
    if not payloads:
        send(chat_id, "گزارش فعالی برای شما در این گفت‌وگو موجود نیست؛ ممکن است منقضی یا پاک شده باشد. فایل یا متن را دوباره بفرستید.")
        return
    try:
        with tempfile.TemporaryDirectory(prefix="review_report_") as tmp:
            for extension in (("html", "txt") if kind == "all" else (kind,)):
                from pathlib import Path
                path = Path(tmp) / ("گزارش-بازبینی." + extension)
                path.write_bytes(payloads[extension])
                send_file(chat_id, path, "گزارش کامل آخرین بررسی شما؛ HTML را با مرورگر باز کنید." if extension == "html" else "گزارش کامل متنی آخرین بررسی شما")
    except BaleAPIError:
        send(chat_id, "ارسال گزارش به‌دلیل مشکل ارتباط با بله تکمیل نشد؛ کمی بعد /report را دوباره بزنید.")


def _report_reaper():
    while not _stop_event.wait(30):
        reports.reap()


def _failure_notice(chat_id, progress, text):
    if progress is not None:
        progress.update(text, final=True)
    else:
        send(chat_id, text)


def process_document(chat_id, doc, annotate_requested=False, mode=None, sender_id=None, progress=None, document_scope="full"):
    tmp = None
    sender_id = chat_id if sender_id is None else sender_id
    try:
        name = safe_filename(doc.get("file_name") or "file")
        tmp = tempfile.mkdtemp(prefix="thesis_")
        mode = mode or modes.get(chat_id, "thesis")
        rules = checker.load_rules(mode=mode)
        if mode in KASHAN_MODE_TITLES:
            rules["document_scope"] = document_scope
        rules["max_input_bytes"] = MAX_MB * 1024 * 1024
        is_docx = name.lower().endswith(".docx")
        is_pdf = name.lower().endswith(".pdf")

        if not (is_docx or is_pdf):
            send(chat_id, "❌ لطفاً فایل با پسوند .docx (Word) یا .pdf ارسال نمایید (برای فایل‌های .doc، در نرم‌افزار Word گزینه Save As به .docx را انتخاب کنید).")
            return

        if (doc.get("file_size") or 0) > MAX_MB * 1024 * 1024:
            send(chat_id, f"❌ حجم فایل فراتر از سقف مجاز ({MAX_MB} مگابایت) است.")
            return

        reports.delete(chat_id, sender_id)
        progress = progress or Progress(chat_id)
        progress.update("⏳ فایل دریافت شد؛ بررسی نگارشی و ساختاری در حال اجراست…")
        src = os.path.join(tmp, name)
        download(doc["file_id"], src)
        validate_document(src, "pdf" if is_pdf else "docx", MAX_MB * 1024 * 1024)

        with DocumentGuard(src, work_dir=tmp) as guard:
            work_src = guard.isolated_working_copy

            analyze = checker.analyze_pdf if is_pdf else checker.analyze
            d, issues, stats = analyze(work_src, mode, rules)
            guard.verify_original_intact()
            progress.update("⏳ بررسی انجام شد؛ خروجی بازبینی در حال آماده‌سازی است…")
            output = None
            if is_docx and d is not None:
                out_docx = os.path.join(tmp, "نسخه-بازبینی-" + name)
                checker.annotate(d, issues, out_docx, max_comments=MAX_COMMENTS, per_par=None)
                stats["annotation_summary"] = getattr(d, "_last_annotation_summary", {})
                stats["annotation_format"] = "docx"
                if stats["annotation_summary"].get("output_created"):
                    output = out_docx
            elif is_pdf and PDF_ANNOTATION_EXPERIMENTAL and stats.get("analysis_status") != "failed":
                from pdf_annotator import annotate_pdf
                out_pdf = os.path.join(tmp, "نسخه-بازبینی-" + name)
                stats["annotation_summary"] = annotate_pdf(work_src, issues, out_pdf, max_comments=MAX_COMMENTS)
                stats["annotation_format"] = "pdf"
                if stats["annotation_summary"].get("output_created"):
                    output = out_pdf
            guard.verify_original_intact()
            artifacts = write_artifacts(os.path.join(tmp, "reports"), issues, stats, rules,
                                        source_path=work_src,
                                        text_report=checker.build_report(issues, stats, rules))
            cached = _cache_reports(chat_id, sender_id, artifacts)
            if output:
                send_file(chat_id, output, "نسخهٔ بازبینی؛ پیشنهادها خودکار اعمال نشده‌اند. راهنمای کامنت‌ها: /comments")
            summary = checker.build_summary(issues, stats, rules)
            if not output and is_docx and stats.get("analysis_status") != "failed":
                summary += "\nخروجی Word آماده نشد؛ جزئیات در گزارش است."
            if cached:
                summary += f"\n📋 گزارش کامل: /report (تا {REPORT_TTL_MINUTES} دقیقه؛ آخرین بررسی شما)"
            else:
                summary += "\nگزارش برای دریافت بعدی در ظرفیت حافظه جا نشد؛ همین حالا ارسال می‌شود."
            if not output and is_pdf:
                send_file(chat_id, artifacts["txt"], "گزارش PDF با محل و شمارهٔ صفحه؛ بررسی علمی محتوا نیازمند بازبینی انسانی است.")
            if not cached:
                send_file(chat_id, artifacts["html"], "گزارش کامل؛ نسخه‌ای برای دریافت بعدی نگهداری نشد.")
                if not (is_pdf and not output):
                    send_file(chat_id, artifacts["txt"], "گزارش کامل متنی")
            progress.update(summary, final=True)

    except BaleAPIError:
        log.warning("Document transport failed")
        _failure_notice(chat_id, progress, "❌ دریافت فایل یا ارسال نتیجه به‌دلیل مشکل ارتباط با بله تکمیل نشد. "
             "ممکن است بخشی از گزارش رسیده باشد؛ کمی بعد دوباره تلاش کنید.")
    except ValueError as e:
        _failure_notice(chat_id, progress, f"❌ حجم فایل بیشتر از {MAX_MB} مگابایت است." if str(e) == "too_big"
             else "❌ فایل ارسال‌شده قابل پردازش نبود.")
    except Exception:
        log.error("Document processing failed")
        _failure_notice(chat_id, progress, "❌ بررسی یا آماده‌سازی خروجی تکمیل نشد؛ کمی بعد دوباره تلاش کنید. این پیام به معنی خراب‌بودن فایل نیست.")
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)


def process_text_message(chat_id, text, mode=None, sender_id=None, progress=None):
    """بررسی متن مستقیم ارسالی کاربر."""
    try:
        sender_id = chat_id if sender_id is None else sender_id
        reports.delete(chat_id, sender_id)
        progress = progress or Progress(chat_id, "⏳ متن در حال بررسی است…")
        mode = mode or modes.get(chat_id, "thesis")
        rules = checker.load_rules(mode=mode)
        _, issues, stats = checker.analyze_text(text, mode, rules)
        with tempfile.TemporaryDirectory(prefix="review_text_") as tmp:
            artifacts = write_artifacts(tmp, issues, stats, rules,
                                        text_report=checker.build_report(issues, stats, rules))
            cached = _cache_reports(chat_id, sender_id, artifacts)
            summary = checker.build_summary(issues, stats, rules)
            if cached:
                summary += f"\n📋 گزارش کامل: /report (تا {REPORT_TTL_MINUTES} دقیقه)"
            else:
                send_file(chat_id, artifacts["html"], "گزارش کامل؛ ظرفیت نگهداری برای دریافت بعدی کافی نبود.")
                send_file(chat_id, artifacts["txt"], "گزارش کامل متنی")
            progress.update(summary, final=True)
    except Exception:
        log.error("Direct text analysis failed")
        _failure_notice(chat_id, progress, "❌ در بررسی متن مشکلی پیش آمد؛ لطفاً دوباره تلاش کنید.")


# ---------------------------- مدیریت پیام‌ها و دستورات ----------------------------
def _submit_job(pool, chat_id, function, *args, _report_request=False, **kwargs):
    token, reason, wait = jobs.reserve(chat_id, cooldown=2 if _report_request else None)
    if reason:
        replies = {
            "busy": "⏳ بررسی قبلی شما در صف یا در حال اجراست؛ لطفاً شکیبا باشید.",
            "cooldown": f"لطفاً {int(wait) + 1} ثانیه دیگر دوباره تلاش کنید.",
            "full": "⏳ ظرفیت پردازش فعلاً پر است؛ کمی بعد دوباره تلاش کنید.",
            "stopping": "سامانه در حال راه‌اندازی مجدد است؛ کمی بعد دوباره تلاش کنید.",
        }
        send(chat_id, replies[reason])
        return False

    if function in (process_document, process_text_message):
        kwargs["progress"] = Progress(chat_id, "⏳ درخواست ثبت شد؛ در صف بررسی…")

    def run():
        try:
            return function(chat_id, *args, **kwargs)
        finally:
            jobs.release(token)

    def done(future):
        if future.cancelled():
            jobs.release(token, refund=True)
        else:
            # Also covers a broken executor that fails before run() begins.
            jobs.release(token)
            if future.exception() is not None:
                log.error("Background job failed (%s)", type(future.exception()).__name__)

    try:
        future = pool.submit(run)
        future.add_done_callback(done)
    except Exception as exc:
        jobs.release(token, refund=True)
        log.warning("Job submission failed (%s)", type(exc).__name__)
        send(chat_id, "ثبت درخواست ممکن نشد؛ لطفاً کمی بعد دوباره تلاش کنید.")
        return False
    return True


def handle(update, pool):
    msg = update.get("message")
    if not msg:
        return
    chat_id = msg["chat"]["id"]
    sender_id = (msg.get("from") or {}).get("id")
    if sender_id is None:
        # Private messages have one owner. Unattributed group posts cannot
        # retrieve another member's reports through a shared fallback identity.
        if msg["chat"].get("type") not in (None, "private"):
            send(chat_id, "برای حفظ دسترسی اختصاصی گزارش، فایل را در گفت‌وگوی خصوصی ربات بفرستید.")
            return
        sender_id = chat_id
    text = (msg.get("text") or "").strip()
    command = text.split(maxsplit=1)[0].split("@", 1)[0].lower() if text.startswith("/") else ""
    # Group commands addressed to another bot must not change this bot's mode.
    first = text.split(maxsplit=1)[0] if text else ""
    if first.startswith("/") and "@" in first:
        target = first.split("@", 1)[1].lower()
        own_name = (_bot_status.get("username") or "").lower()
        if target != own_name:
            return
    _bot_status["last_message"] = {
        "kind": "document" if msg.get("document") else "text",
        "time": time.strftime("%Y-%m-%d %H:%M:%S")
    }

    if command == "/start" or text in ("🔄 شروع مجدد", "شروع مجدد", "شروع", "start"):
        send(chat_id, WELCOME)
    elif command == "/help" or text in ("📖 راهنما", "راهنما", "help"):
        send(chat_id, HELP)
    elif command == "/thesis" or text in ("🎓 پایان‌نامه", "پایان‌نامه", "پایان نامه"):
        modes[chat_id] = "thesis"
        send(chat_id, "🎓 الگوی «پایان‌نامه» فعال شد. فایل Word، PDF یا متن خود را ارسال کنید.")
    elif command == "/tums":
        modes[chat_id] = "tums"
        send(chat_id, "🏛 پروفایل دانشگاه علوم پزشکی تهران فعال شد. نتایج شیوه‌نامه برای بازبینی ارائه می‌شوند.")
    elif command == "/kashan" or text in ("🏛 دانشگاه کاشان", "دانشگاه کاشان", "کاشان"):
        args = text.split()[1:] if command else []
        choices = {"master": "kashan_master", "phd": "kashan_phd"}
        if len(args) > 1 or (args and args[0].lower() not in choices):
            send(chat_id, "دانشگاه کاشان: /kashan\nانتخاب مستقیم: /kashan master یا /kashan phd", reply_markup=KASHAN_KEYBOARD)
            return
        selected = choices[args[0].lower()] if args else "kashan"
        modes[chat_id] = selected
        document_scopes.setdefault(chat_id, "fragment")
        selection = ("🏛 " + KASHAN_MODE_TITLES[selected] + " فعال شد. دامنهٔ سند را مشخص کنید؛ سپس فایل را بفرستید.\n"
                     if args else "🏛 دانشگاه کاشان انتخاب شد. مقطع را از دکمه‌ها انتخاب کنید؛ سپس مشخص کنید سند کامل است یا بخشی از آن.\n")
        send(chat_id, selection +
             "مبنای پروفایل، مشخصات ارائه‌شدهٔ مالک ربات است؛ /kashan_info منابع و محدودیت‌ها را نشان می‌دهد.", reply_markup=KASHAN_KEYBOARD)
    elif command in {"/kashan_master", "/kashan_phd"} or text in ("🎓 کاشان — کارشناسی ارشد", "🎓 کاشان — دکتری"):
        is_phd = command == "/kashan_phd" or text == "🎓 کاشان — دکتری"
        modes[chat_id] = "kashan_phd" if is_phd else "kashan_master"
        scope = document_scopes.setdefault(chat_id, "fragment")
        send(chat_id, "🏛 " + KASHAN_MODE_TITLES[modes[chat_id]] + " فعال شد.\n"
             + ("دامنهٔ فعلی: سند کامل." if scope == "full" else "دامنهٔ فعلی: بخشی از سند.")
             + " در صورت نیاز دامنه را از دکمه‌ها تغییر دهید؛ سپس فایل را بفرستید.\n"
             "مبنای قواعد و موارد قابل تأیید نیست: /kashan_info", reply_markup=KASHAN_KEYBOARD)
    elif command == "/scope" or text in ("📄 کاشان — سند کامل", "📑 کاشان — بخشی از سند"):
        if modes.get(chat_id) not in KASHAN_MODE_TITLES:
            send(chat_id, "ابتدا دانشگاه کاشان را از /kashan انتخاب کنید.")
            return
        args = text.split()[1:] if command else []
        options = {"full": "full", "part": "fragment", "fragment": "fragment"}
        if command and (len(args) != 1 or args[0].lower() not in options):
            send(chat_id, "سند کامل: /scope full\nیک فصل یا بخشی از سند: /scope part", reply_markup=KASHAN_KEYBOARD)
            return
        scope = options[args[0].lower()] if command else ("full" if text == "📄 کاشان — سند کامل" else "fragment")
        document_scopes[chat_id] = scope
        send(chat_id, "دامنهٔ بررسی کاشان: " + ("سند کامل" if scope == "full" else "بخشی از سند؛ نبود بخش‌های یک پایان‌نامهٔ کامل ایراد محسوب نمی‌شود")
             + ". فایل را بفرستید.", reply_markup=KASHAN_KEYBOARD)
    elif command == "/kashan_info" or text == "ℹ️ راهنمای کاشان":
        send(chat_id, KASHAN_GUIDANCE, reply_markup=KASHAN_KEYBOARD)
    elif text == "⬅️ منوی اصلی":
        send(chat_id, "منوی اصلی؛ حالت انتخاب‌شده حفظ شد. فایل یا متن خود را بفرستید.")
    elif command == "/article" or text in ("📑 مقاله پژوهشی", "مقاله پژوهشی", "مقاله"):
        modes[chat_id] = "article"
        send(chat_id, "📑 الگوی «مقاله علمی/پژوهشی» فعال شد. فایل Word، PDF یا متن خود را ارسال کنید.")
    elif command == "/english" or text in ("🇬🇧 متن انگلیسی", "متن انگلیسی", "انگلیسی"):
        modes[chat_id] = "english"
        send(chat_id, "🇬🇧 الگوی «متن/مقاله انگلیسی» فعال شد. متن یا فایل انگلیسی خود را ارسال فرمایید.")
    elif command == "/report" or text == "📋 گزارش کامل":
        args = text.split()[1:] if command else []
        if len(args) > 1 or (args and args[0].lower() not in {"html", "txt", "all"}):
            send(chat_id, "گزارش آخرین بررسی شما: /report\nفقط یک قالب: /report html یا /report txt")
            return
        _submit_job(pool, chat_id, deliver_reports, sender_id=sender_id,
                    kind=args[0].lower() if args else "all", _report_request=True)
    elif command == "/comments" or text == "💬 راهنمای کامنت‌ها":
        send(chat_id, COMMENTS_HELP)
    elif command == "/forget" or text == "🗑 حذف گزارش":
        reports.delete(chat_id, sender_id)
        send(chat_id, "گزارش نگهداری‌شدهٔ شما در این گفت‌وگو پاک شد. اگر بررسی تازه‌ای در حال اجرا باشد، گزارش آن پس از پایان آماده می‌شود.")
    elif command in {"/status", "/settings"} or text in ("ℹ️ وضعیت جاری", "وضعیت جاری", "وضعیت"):
        cur_mode = modes.get(chat_id, "thesis")
        title_map = {
            "thesis": "پایان‌نامه دانشگاهی 🎓",
            "tums": "دانشگاه علوم پزشکی تهران 🏛",
            "article": "مقاله پژوهشی 📑",
            "english": "متن انگلیسی 🇬🇧"
        }
        title_map.update(KASHAN_MODE_TITLES)
        with lock:
            is_busy = chat_id in busy
        scope_description = ""
        if cur_mode in KASHAN_MODE_TITLES:
            scope_description = ("دامنهٔ کاشان: " + ("سند کامل" if document_scopes.get(chat_id, "fragment") == "full" else "بخشی از سند")
                                 + "; تغییر: /scope full یا /scope part\n")
        send(chat_id, f"📌 حالت فعال شما: {title_map.get(cur_mode, cur_mode)}\n"
             f"وضعیت درخواست: {'در صف یا در حال اجرا' if is_busy else 'آماده دریافت'}\n"
             f"سقف فایل: {MAX_MB} مگابایت؛ سقف متن: {MAX_TEXT_CHARS:,} نویسه\n"
             f"Word با کامنت، خروجی پیش‌فرض است؛ سقف کامنت: {MAX_COMMENTS}.\n"
             + scope_description +
             f"گزارش کامل: /report؛ نگهداری در حافظه تا {REPORT_TTL_MINUTES} دقیقه.\n"
             "انتخاب حالت: /thesis، /tums، /kashan، /article یا /english.")
    elif command == "/privacy":
        send(chat_id, "🔒 فایل شما برای بررسی به میزبان ربات منتقل می‌شود و فایل‌های موقت پس از پایان یا خطای پردازش پاک‌سازی می‌شوند. "
             f"فقط گزارش آخرین بررسی (شامل شاهدهایی از متن) تا {REPORT_TTL_MINUTES} دقیقه در حافظهٔ میزبان می‌ماند؛ "
             "پاک‌سازی دوره‌ای حداکثر ۳۰ ثانیه بعد از انقضا انجام می‌شود. "
             "دسترسی به آن به فرستنده و همین گفت‌وگو محدود است؛ /forget آن را حذف می‌کند. "
             "ظرفیت محدود، بررسی جدید و راه‌اندازی مجدد ممکن است گزارش را زودتر حذف کنند. "
             "این ربات متن سند یا شناسه کاربران را در گزارش سلامت منتشر نمی‌کند و سرویس هوش مصنوعی خارجی برای بررسی فراخوانی نمی‌کند. "
             "نگهداری نسخه‌های پیام و فایل در بله و زیرساخت میزبان تابع سیاست‌های آن‌هاست. "
             "پس از توقف ناگهانی سیستم ممکن است پاک‌سازی عادی کامل نشود. نتایج جایگزین بازبینی پژوهشگر نیستند.")
    elif command:
        send(chat_id, "این دستور شناخته نشد. /help راهنما و /settings تنظیمات فعلی را نشان می‌دهد. "
             "Word به‌طور پیش‌فرض کامنت‌دار برمی‌گردد؛ /report گزارش کامل را می‌فرستد.")
    elif msg.get("document"):
        doc = msg["document"]
        name = doc.get("file_name") or ""
        if not name.lower().endswith((".docx", ".pdf")):
            send(chat_id, "لطفاً فایل Word با پسوند .docx یا PDF بفرستید.")
            return
        file_size = doc.get("file_size") or 0
        if not isinstance(file_size, (int, float)) or file_size < 0 or file_size > MAX_MB * 1024 * 1024:
            send(chat_id, f"سقف حجم فایل {MAX_MB} مگابایت است.")
            return
        _submit_job(pool, chat_id, process_document, doc,
                    annotate_requested="/annotate" in (msg.get("caption") or "").split(),
                    mode=modes.get(chat_id, "thesis"), sender_id=sender_id,
                    document_scope=("fragment" if "/part" in (msg.get("caption") or "").split()
                                    else document_scopes.get(chat_id, "fragment" if modes.get(chat_id) in KASHAN_MODE_TITLES else "full")))
    elif text:
        if len(text) > MAX_TEXT_CHARS:
            send(chat_id, f"سقف متن مستقیم {MAX_TEXT_CHARS:,} نویسه است؛ متن طولانی‌تر را در فایل ارسال کنید.")
            return
        _submit_job(pool, chat_id, process_text_message, msg["text"], mode=modes.get(chat_id, "thesis"), sender_id=sender_id)


def _health_payload(readiness=False):
    age = None if _last_poll_monotonic is None else max(0, time.monotonic() - _last_poll_monotonic)
    ready = bool(_bot_status["started"] and not _stop_event.is_set() and
                 age is not None and age <= HEALTH_STALE_SEC and
                 not _bot_status["last_poll_error"])
    revision = os.environ.get("RENDER_GIT_COMMIT", "").lower()
    if len(revision) != 40 or any(char not in "0123456789abcdef" for char in revision):
        revision = None
    return {"status": ("ready" if ready else "not_ready") if readiness else "ok",
            "service": "virastar-bale-bot", "version": PRODUCT_VERSION,
            "revision": revision, "ready": ready,
            "bot_status": {key: _bot_status[key] for key in ("started", "last_poll_time", "total_polls")},
            "queue": jobs.snapshot()}


def _start_health_server():
    """اجرای وب‌سرور سبک روی 0.0.0.0 برای تایید سلامت در هاست‌های ابری (مانند Render)."""
    global _health_server
    if _health_server is not None:
        return _health_server
    port_str = os.environ.get("PORT")
    if not port_str:
        return
    port, error = env_integer(os.environ, "PORT", 10000, 1, 65535)
    if error:
        raise ValueError(error)

    from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
    import json

    class HealthHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            path = self.path.split("?", 1)[0]
            if path not in {"/", "/health", "/ready"}:
                self.send_error(404)
                return
            payload = _health_payload(readiness=path == "/ready")
            self.send_response(503 if path == "/ready" and not payload["ready"] else 200)
            self.send_header("Content-type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))

        def log_message(self, format, *args):
            return

    server = ThreadingHTTPServer(("0.0.0.0", port), HealthHandler)
    _health_server = server
    t = threading.Thread(target=server.serve_forever, daemon=True, name="health")
    t.start()
    log.info("Health server listening on 0.0.0.0:%d", port)
    return server


def _auto_keep_alive():
    """ارسال پینگ خودکار برای جلوگیری از خاموش شدن کانتینر در Render."""
    url = os.environ.get("KEEP_ALIVE_URL", "").strip()
    if not url:
        return
    url = validated_base_url(url, "KEEP_ALIVE_URL")
    while not _stop_event.wait(300):
        try:
            with http.get(url, timeout=(5, 10), allow_redirects=False):
                pass
        except Exception:
            log.warning("Optional keep-alive request failed")


def _acquire_lock_port(port=49250):
    """جلوگیری از اجرای همزمان چند نسخه از ربات با استفاده از پورت لوکال."""
    if os.environ.get("RENDER") or os.environ.get("PORT"):
        return True
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
        return s
    except OSError:
        s.close()
        log.error("خطا: یک نسخه دیگر از ربات در حال حاضر در حال اجرا است! از اجرای همزمان چند ربات جلوگیری شد.")
        return None


def setup_commands():
    """ثبت رسمی دستورات ربات در منوی پیام‌رسان بله."""
    try:
        commands = [
            {"command": "start", "description": "شروع کار با ربات"},
            {"command": "thesis", "description": "بررسی پایان‌نامه و رساله"},
            {"command": "tums", "description": "پروفایل دانشگاه علوم پزشکی تهران"},
            {"command": "kashan", "description": "دانشگاه کاشان؛ انتخاب مقطع"},
            {"command": "kashan_master", "description": "کاشان — کارشناسی ارشد"},
            {"command": "kashan_phd", "description": "کاشان — دکتری"},
            {"command": "kashan_info", "description": "منابع و محدودیت‌های پروفایل کاشان"},
            {"command": "scope", "description": "کاشان؛ سند کامل یا بخشی از سند"},
            {"command": "article", "description": "بررسی مقاله پژوهشی/علمی"},
            {"command": "english", "description": "بررسی متون انگلیسی"},
            {"command": "status", "description": "مشاهده وضعیت جاری"},
            {"command": "settings", "description": "حالت و محدودیت‌های بررسی"},
            {"command": "privacy", "description": "حریم خصوصی و نگهداری فایل"},
            {"command": "report", "description": "گزارش کامل آخرین بررسی شما"},
            {"command": "comments", "description": "راهنمای دیدن کامنت‌های Word"},
            {"command": "forget", "description": "حذف گزارش نگهداری‌شده"},
            {"command": "help", "description": "راهنمای کار با ربات"}
        ]
        call("setMyCommands", commands=commands)
        log.info("Bale bot commands registered successfully")
    except Exception as e:
        log.warning("Could not set bot commands (%s)", type(e).__name__)


def main():
    global API, FILE_API, _last_poll_monotonic, _health_server
    if _config_errors:
        raise SystemExit("Invalid configuration: " + "; ".join(_config_errors))
    if not TOKEN:
        raise SystemExit("متغیر محیطی BALE_TOKEN تنظیم نشده است.")
    try:
        API = validated_base_url(API, "BALE_API_BASE")
        FILE_API = validated_base_url(FILE_API, "BALE_FILE_BASE")
        if os.environ.get("KEEP_ALIVE_URL"):
            validated_base_url(os.environ["KEEP_ALIVE_URL"].strip(), "KEEP_ALIVE_URL")
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc
    lock_sock = _acquire_lock_port()
    if lock_sock is None:
        raise SystemExit("یک نسخه دیگر از ربات در پس‌زمینه در حال اجراست.")
    pool = ThreadPoolExecutor(max_workers=WORKERS, thread_name_prefix="review")
    offset = None
    previous_signals = {}
    if threading.current_thread() is threading.main_thread():
        for sig in (signal.SIGINT, signal.SIGTERM):
            previous_signals[sig] = signal.getsignal(sig)
            signal.signal(sig, lambda *_: _stop_event.set())
    try:
        _start_health_server()
        threading.Thread(target=_report_reaper, daemon=True, name="report-expiry").start()
        if os.environ.get("KEEP_ALIVE_URL"):
            threading.Thread(target=_auto_keep_alive, daemon=True, name="keep-alive").start()
        me = call("getMe")
        if not isinstance(me, dict):
            raise BaleAPIError("getMe: malformed response")
        _bot_status["username"] = me.get("username")
        _bot_status["started"] = True
        _bot_status["start_time"] = time.strftime("%Y-%m-%d %H:%M:%S")
        log.info("Bot authenticated and polling started")
        setup_commands()
        while not _stop_event.is_set():
            try:
                payload = {"timeout": 30}
                if offset is not None:
                    payload["offset"] = offset
                upds = call("getUpdates", **payload)
                if not isinstance(upds, list):
                    raise BaleAPIError("getUpdates: malformed response")
                _last_poll_monotonic = time.monotonic()
                _bot_status["last_poll_time"] = time.strftime("%Y-%m-%d %H:%M:%S")
                _bot_status["last_poll_count"] = len(upds)
                _bot_status["total_polls"] += 1
                _bot_status["last_poll_error"] = None
                for upd in upds:
                    if _stop_event.is_set():
                        break
                    if not isinstance(upd, dict) or not isinstance(upd.get("update_id"), int):
                        log.warning("Malformed update ignored")
                        continue
                    if offset is not None and upd["update_id"] < offset:
                        continue
                    try:
                        handle(upd, pool)
                    except Exception as exc:
                        log.error("Handler failed (%s)", type(exc).__name__)
                    offset = upd["update_id"] + 1
            except KeyboardInterrupt:
                _stop_event.set()
            except Exception as exc:
                _bot_status["last_poll_error"] = type(exc).__name__
                log.warning("Polling error (%s)", type(exc).__name__)
                _stop_event.wait(5)
    except KeyboardInterrupt:
        _stop_event.set()
    except Exception as exc:
        _bot_status["last_poll_error"] = type(exc).__name__
        log.error("Bot startup failed (%s)", type(exc).__name__)
        raise SystemExit("راه‌اندازی ربات موفق نبود؛ تنظیمات و اتصال میزبان را بررسی کنید.") from None
    finally:
        _stop_event.set()
        jobs.stop()
        _bot_status["started"] = False
        pool.shutdown(wait=True, cancel_futures=True)
        reports.close()
        if _health_server is not None:
            _health_server.shutdown()
            _health_server.server_close()
            _health_server = None
        http.close()
        if hasattr(lock_sock, "close"):
            lock_sock.close()
        for sig, handler in previous_signals.items():
            signal.signal(sig, handler)


if __name__ == "__main__":
    main()
