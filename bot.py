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
jobs = JobGate(WORKERS + MAX_QUEUE, COOLDOWN)
busy, last_job, lock = jobs.busy, jobs.last_started, jobs.lock
http = ThreadLocalHTTP(requests.Session)
_stop_event = threading.Event()
_last_poll_monotonic = None
_health_server = None

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
    "• علامت‌گذاری زرد روی خطاهای متنی در فایل Word همراه با کامنت‌های تفصیلی (بدون دستکاری متن اصلی)\n\n"
    "📌 انتخاب نوع سند:\n"
    "🎓 /thesis : پایان‌نامه و رساله دانشگاهی (پیش‌فرض)\n"
    "🏛 /tums : بررسی با پروفایل دانشگاه علوم پزشکی تهران\n"
    "📑 /article : مقاله پژوهشی / علمی\n"
    "🇬🇧 /english : بررسی متون و مقالات انگلیسی\n"
    "ℹ️ /status : مشاهده وضعیت جاری\n"
    "📖 /help : راهنمای کامل\n\n"
    "📎 می‌توانید فایل Word (.docx) یا PDF ارسال کنید یا متن خود را مستقیماً در چت بنویسید."
)

HELP = (
    "📖 راهنمای کار با ربات ویراستار پژوهش:\n\n"
    "۱) انتخاب نوع سند:\n"
    "• دستور /thesis برای پایان‌نامه‌ها و رساله‌ها (پیش‌فرض).\n"
    "• دستور /tums برای پروفایل دانشگاه علوم پزشکی تهران؛ انطباق قطعی نیازمند بازبینی شیوه‌نامه است.\n"
    "• دستور /article برای مقالات پژوهشی و علمی.\n"
    "• دستور /english برای بازبینی قاعده‌محور متون انگلیسی.\n\n"
    "۲) ارسال ورودی:\n"
    "• فایل Word (.docx): خروجی پیش‌فرض گزارش است. برای دریافت کپی حاشیه‌نویسی‌شده، در توضیح فایل /annotate بنویسید. همچنین خلاصه‌ای در چت و فایل متنی گزارش ارسال خواهد شد.\n"
    "• فایل PDF: متن استخراج شده و گزارش جامع ارزیابی صادر می‌شود (فایل‌های تصویری نیازمند OCR هستند).\n"
    "• متن مستقیم: یک یا چند بند از متن خود را بفرستید تا فوراً بررسی گردد.\n\n"
    "۳) /settings تنظیمات و /privacy توضیح حریم خصوصی را نمایش می‌دهند.\n"
    "🔒 فایل‌های موقت این پردازش پاک‌سازی می‌شوند؛ نگهداری داده در بله و میزبان تابع سیاست‌های آن‌هاست.\n"
    "نتایج پیشنهاد بازبینی‌اند؛ بررسی علمی محتوا، صحت منابع و استخراج PDF کامل یا قطعی نیست."
)

MAIN_KEYBOARD = {
    "keyboard": [
        [{"text": "🎓 پایان‌نامه"}, {"text": "📑 مقاله پژوهشی"}],
        [{"text": "🇬🇧 متن انگلیسی"}, {"text": "ℹ️ وضعیت جاری"}],
        [{"text": "📖 راهنما"}, {"text": "🔄 شروع مجدد"}]
    ],
    "resize_keyboard": True
}


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


def send(chat_id, text, reply_markup=MAIN_KEYBOARD):
    parts = checker.split_message(text, 3800)
    for i, part in enumerate(parts):
        payload = {"chat_id": chat_id, "text": part}
        if i == len(parts) - 1 and reply_markup is not None:
            payload["reply_markup"] = reply_markup
        try:
            call("sendMessage", **payload)
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
def process_document(chat_id, doc, annotate_requested=False, mode=None):
    tmp = None
    try:
        name = safe_filename(doc.get("file_name") or "file")
        tmp = tempfile.mkdtemp(prefix="thesis_")
        mode = mode or modes.get(chat_id, "thesis")
        rules = checker.load_rules(mode=mode)
        rules["max_input_bytes"] = MAX_MB * 1024 * 1024
        is_docx = name.lower().endswith(".docx")
        is_pdf = name.lower().endswith(".pdf")

        if not (is_docx or is_pdf):
            send(chat_id, "❌ لطفاً فایل با پسوند .docx (Word) یا .pdf ارسال نمایید (برای فایل‌های .doc، در نرم‌افزار Word گزینه Save As به .docx را انتخاب کنید).")
            return

        if (doc.get("file_size") or 0) > MAX_MB * 1024 * 1024:
            send(chat_id, f"❌ حجم فایل فراتر از سقف مجاز ({MAX_MB} مگابایت) است.")
            return

        send(chat_id, "⏳ فایل دریافت شد؛ فرآیند ارزیابی علمی، نگارشی و ساختاری در حال اجراست…")
        src = os.path.join(tmp, name)
        download(doc["file_id"], src)
        validate_document(src, "pdf" if is_pdf else "docx", MAX_MB * 1024 * 1024)

        with DocumentGuard(src, work_dir=tmp) as guard:
            work_src = guard.isolated_working_copy

            analyze = checker.analyze_pdf if is_pdf else checker.analyze
            d, issues, stats = analyze(work_src, mode, rules)
            guard.verify_original_intact()
            artifacts = write_artifacts(os.path.join(tmp, "reports"), issues, stats, rules,
                                        source_path=work_src,
                                        text_report=checker.build_report(issues, stats, rules, examples=1000))
            if not send(chat_id, checker.build_summary(issues, stats, rules)):
                raise BaleAPIError("Could not deliver summary")
            send_file(chat_id, artifacts["html"], "گزارش قابل جست‌وجو؛ فایل را با مرورگر باز کنید. پردازش این صفحه آفلاین است.")
            send_file(chat_id, artifacts["txt"], "گزارش متنی بازبینی سند")
            if annotate_requested and d is not None:
                out_docx = os.path.join(tmp, "نسخه-بررسی‌شده-" + name)
                if checker.annotate(d, issues, out_docx):
                    send_file(chat_id, out_docx, "کپی Word با کامنت‌های بازبینی؛ متن پیشنهادها خودکار اعمال نشده است.")
                    result = getattr(d, "_last_annotation_summary", {})
                    if result.get("unapplied"):
                        send(chat_id, f"{result['applied']} کامنت درج شد؛ سایر موارد در گزارش موجودند و لزوماً محل امنی برای نشانه‌گذاری ندارند.")
                else:
                    send(chat_id, "برای این سند محدودهٔ قابل‌اطمینانی برای درج کامنت پیدا نشد؛ جزئیات در گزارش موجود است.")

    except BaleAPIError:
        log.warning("Document transport failed")
        send(chat_id, "❌ دریافت فایل یا ارسال نتیجه به‌دلیل مشکل ارتباط با بله تکمیل نشد. "
             "ممکن است بخشی از گزارش رسیده باشد؛ کمی بعد دوباره تلاش کنید.")
    except ValueError as e:
        send(chat_id, f"❌ حجم فایل بیشتر از {MAX_MB} مگابایت است." if str(e) == "too_big"
             else "❌ فایل ارسال‌شده قابل پردازش نبود.")
    except Exception:
        log.error("Document processing failed")
        send(chat_id, "❌ متأسفانه در فرآیند بررسی فایل خطایی رخ داد. از صحت و سلامت فایل اطمینان حاصل کرده و دوباره تلاش فرمایید.")
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)


def process_text_message(chat_id, text, mode=None):
    """بررسی متن مستقیم ارسالی کاربر."""
    try:
        mode = mode or modes.get(chat_id, "thesis")
        rules = checker.load_rules(mode=mode)
        _, issues, stats = checker.analyze_text(text, mode, rules)
        report = checker.build_report(issues, stats, rules, examples=20)
        send(chat_id, report)
    except Exception:
        log.error("Direct text analysis failed")
        send(chat_id, "❌ در بررسی متن مشکلی پیش آمد؛ لطفاً دوباره تلاش کنید.")


# ---------------------------- مدیریت پیام‌ها و دستورات ----------------------------
def _submit_job(pool, chat_id, function, *args, **kwargs):
    token, reason, wait = jobs.reserve(chat_id)
    if reason:
        replies = {
            "busy": "⏳ بررسی قبلی شما در صف یا در حال اجراست؛ لطفاً شکیبا باشید.",
            "cooldown": f"لطفاً {int(wait) + 1} ثانیه دیگر دوباره تلاش کنید.",
            "full": "⏳ ظرفیت پردازش فعلاً پر است؛ کمی بعد دوباره تلاش کنید.",
            "stopping": "سامانه در حال راه‌اندازی مجدد است؛ کمی بعد دوباره تلاش کنید.",
        }
        send(chat_id, replies[reason])
        return False

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
    elif command == "/article" or text in ("📑 مقاله پژوهشی", "مقاله پژوهشی", "مقاله"):
        modes[chat_id] = "article"
        send(chat_id, "📑 الگوی «مقاله علمی/پژوهشی» فعال شد. فایل Word، PDF یا متن خود را ارسال کنید.")
    elif command == "/english" or text in ("🇬🇧 متن انگلیسی", "متن انگلیسی", "انگلیسی"):
        modes[chat_id] = "english"
        send(chat_id, "🇬🇧 الگوی «متن/مقاله انگلیسی» فعال شد. متن یا فایل انگلیسی خود را ارسال فرمایید.")
    elif command in {"/status", "/settings"} or text in ("ℹ️ وضعیت جاری", "وضعیت جاری", "وضعیت"):
        cur_mode = modes.get(chat_id, "thesis")
        title_map = {
            "thesis": "پایان‌نامه دانشگاهی 🎓",
            "tums": "دانشگاه علوم پزشکی تهران 🏛",
            "article": "مقاله پژوهشی 📑",
            "english": "متن انگلیسی 🇬🇧"
        }
        with lock:
            is_busy = chat_id in busy
        send(chat_id, f"📌 حالت فعال شما: {title_map.get(cur_mode, cur_mode)}\n"
             f"وضعیت درخواست: {'در صف یا در حال اجرا' if is_busy else 'آماده دریافت'}\n"
             f"سقف فایل: {MAX_MB} مگابایت؛ سقف متن: {MAX_TEXT_CHARS:,} نویسه\n"
             "حاشیه‌نویسی Word با /annotate در توضیح فایل فعال می‌شود.\n"
             "برای تغییر حالت از /thesis، /tums، /article یا /english استفاده کنید.")
    elif command == "/privacy":
        send(chat_id, "🔒 فایل شما برای بررسی به میزبان ربات منتقل می‌شود و فایل‌های موقت پس از پایان یا خطای پردازش پاک‌سازی می‌شوند. "
             "این ربات متن سند یا شناسه کاربران را در گزارش سلامت منتشر نمی‌کند و سرویس هوش مصنوعی خارجی برای بررسی فراخوانی نمی‌کند. "
             "نگهداری نسخه‌های پیام و فایل در بله و زیرساخت میزبان تابع سیاست‌های آن‌هاست. "
             "پس از توقف ناگهانی سیستم ممکن است پاک‌سازی عادی کامل نشود. نتایج جایگزین بازبینی پژوهشگر نیستند.")
    elif command:
        send(chat_id, "این دستور شناخته نشد. /help راهنما و /settings تنظیمات فعلی را نشان می‌دهد. "
             "برای حاشیه‌نویسی، /annotate را در توضیح فایل Word بنویسید.")
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
                    mode=modes.get(chat_id, "thesis"))
    elif text:
        if len(text) > MAX_TEXT_CHARS:
            send(chat_id, f"سقف متن مستقیم {MAX_TEXT_CHARS:,} نویسه است؛ متن طولانی‌تر را در فایل ارسال کنید.")
            return
        _submit_job(pool, chat_id, process_text_message, msg["text"], mode=modes.get(chat_id, "thesis"))


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
            {"command": "article", "description": "بررسی مقاله پژوهشی/علمی"},
            {"command": "english", "description": "بررسی متون انگلیسی"},
            {"command": "status", "description": "مشاهده وضعیت جاری"},
            {"command": "settings", "description": "حالت و محدودیت‌های بررسی"},
            {"command": "privacy", "description": "حریم خصوصی و نگهداری فایل"},
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
