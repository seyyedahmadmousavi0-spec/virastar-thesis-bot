# -*- coding: utf-8 -*-
"""
بازوی بله «ویراستار پژوهش» برای ارزیابی، ویرایش علمی و بررسی پایان‌نامه و مقاله.
پشتیبانی از فایل‌های Word (.docx)، فایل‌های PDF و متون مستقیم ارسالی در چت.
مجهز به پروفایل تخصصی دانشگاه علوم پزشکی تهران (TUMS) و بررسی متون انگلیسی.
"""
import logging
import os
import shutil
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import requests

import checker

def _load_env_file():
    """اگر فایل .env کنار bot.py باشد، متغیرها را از آن می‌خواند (BALE_TOKEN=...)."""
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

TOKEN = os.environ.get("BALE_TOKEN", "")
API = os.environ.get("BALE_API_BASE", "https://tapi.bale.ai").rstrip("/")
FILE_API = os.environ.get("BALE_FILE_BASE", f"{API}/file").rstrip("/")
MAX_MB = int(os.environ.get("MAX_FILE_MB", "25"))
COOLDOWN = int(os.environ.get("COOLDOWN_SEC", "10"))
WORKERS = int(os.environ.get("WORKERS", "4"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("bot")

modes = {}        # chat_id -> "tums" | "thesis" | "article" | "english"
busy = set()      # چت‌هایی که فایلشان در حال بررسی است
last_job = {}     # chat_id -> زمان آخرین بررسی
lock = threading.Lock()
http = requests.Session()

WELCOME = (
    "سلام 👋\n"
    "به سامانه بازوی «ویراستار پژوهش» خوش آمدید.\n\n"
    "این سامانه جهت بررسی، ویرایش و ارزیابی جامع پایان‌نامه‌ها، رساله‌ها و مقالات علمی طراحی شده است:\n"
    "• املای دقیق و تایپوگرافی فارسی (نیم‌فاصله، ی/ک، ارقام، پیشوندها و پسوندها)\n"
    "• سبک و لحن علمی (حذف عبارات مبالغه‌آمیز، محاوره‌ای و ادعاهای قطعی بدون شاهد)\n"
    "• دستور زبان و ساختار متون فارسی و انگلیسی (Grammar, Articles, Prepositions)\n"
    "• ارزیابی انسجام، پیوستگی و ساختار فصول\n"
    "• ارجاع‌دهی درون‌متنی و تطبیق با فهرست منابع (ونکوور، هاروارد و APA)\n"
    "• صفحه‌آرایی فایل Word (حاشیه‌ها، فونت‌ها، تراز و شماره‌گذاری)\n"
    "• علامت‌گذاری زرد روی خطاهای متنی در فایل Word همراه با کامنت‌های تفصیلی (بدون دستکاری متن اصلی)\n\n"
    "📌 انتخاب نوع سند:\n"
    "🎓 /thesis : پایان‌نامه و رساله دانشگاهی (پیش‌فرض)\n"
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
    "• دستور /article برای مقالات پژوهشی و علمی.\n"
    "• دستور /english برای ویرایش تخصصی متون انگلیسی.\n\n"
    "۲) ارسال ورودی:\n"
    "• فایل Word (.docx): فایل شما بدون هیچ دستکاری در متن تحویل داده می‌شود؛ کلمات دارای ایراد با رنگ زرد هایلایت شده و کامنت‌های راهنما روی آن‌ها درج می‌گردد. همچنین خلاصه‌ای در چت و فایل متنی گزارش ارسال خواهد شد.\n"
    "• فایل PDF: متن استخراج شده و گزارش جامع ارزیابی صادر می‌شود (فایل‌های تصویری نیازمند OCR هستند).\n"
    "• متن مستقیم: یک یا چند بند از متن خود را بفرستید تا فوراً بررسی گردد.\n\n"
    "🔒 فایل‌های ارسالی پس از بررسی بلافاصله از سرور حذف می‌شوند."
)


# ---------------------------- API بله ----------------------------
def call(method, **payload):
    r = http.post(f"{API}/bot{TOKEN}/{method}", json=payload, timeout=45)
    data = r.json()
    if not data.get("ok"):
        raise RuntimeError(f"{method}: {data}")
    return data["result"]


def send(chat_id, text):
    for part in checker.split_message(text, 3800):
        try:
            call("sendMessage", chat_id=chat_id, text=part)
        except Exception as e:
            log.warning("sendMessage failed: %s", e)


def send_file(chat_id, path, caption=""):
    with open(path, "rb") as f:
        r = http.post(f"{API}/bot{TOKEN}/sendDocument",
                      data={"chat_id": chat_id, "caption": caption},
                      files={"document": (os.path.basename(path), f)}, timeout=120)
    if not r.json().get("ok"):
        raise RuntimeError(f"sendDocument: {r.text[:200]}")


def download(file_id, dest):
    info = call("getFile", file_id=file_id)
    url = f"{FILE_API}/bot{TOKEN}/{info['file_path']}"
    with http.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        size = 0
        with open(dest, "wb") as f:
            for chunk in r.iter_content(65536):
                size += len(chunk)
                if size > MAX_MB * 1024 * 1024:
                    raise ValueError("too_big")
                f.write(chunk)


# ---------------------------- پردازش اسناد ----------------------------
def process_document(chat_id, doc):
    name = doc.get("file_name") or "file"
    tmp = tempfile.mkdtemp(prefix="thesis_")
    mode = modes.get(chat_id, "thesis")
    rules = checker.load_rules(mode=mode)

    try:
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

        if is_pdf:
            d, issues, stats = checker.analyze_pdf(src, mode, rules)
            send(chat_id, checker.build_summary(issues, stats, rules))

            report_path = os.path.join(tmp, "گزارش-ارزیابی-سند.txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(checker.build_report(issues, stats, rules, examples=1000))
            send_file(chat_id, report_path, "📋 گزارش تفصیلی ارزیابی فایل PDF")
        else:
            d, issues, stats = checker.analyze(src, mode, rules)
            send(chat_id, checker.build_summary(issues, stats, rules))

            report_path = os.path.join(tmp, "گزارش-ارزیابی-سند.txt")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(checker.build_report(issues, stats, rules, examples=1000))
            send_file(chat_id, report_path, "📋 گزارش تفصیلی ارزیابی فایل Word")

            out_docx = os.path.join(tmp, "نسخه-بررسی‌شده-" + name)
            if checker.annotate(d, issues, out_docx):
                send_file(chat_id, out_docx, "📝 نسخه Word همراه با کامنت‌های تفصیلی (برای مشاهده، در تب Review بخش Comments را باز کنید)")

    except ValueError as e:
        send(chat_id, f"❌ حجم فایل بیشتر از {MAX_MB} مگابایت است." if str(e) == "too_big"
             else "❌ فایل ارسال‌شده قابل پردازش نبود.")
    except Exception:
        log.exception("processing failed")
        send(chat_id, "❌ متأسفانه در فرآیند بررسی فایل خطایی رخ داد. از صحت و سلامت فایل اطمینان حاصل کرده و دوباره تلاش فرمایید.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        with lock:
            busy.discard(chat_id)


def process_text_message(chat_id, text):
    """بررسی متن مستقیم ارسالی کاربر."""
    mode = modes.get(chat_id, "thesis")
    rules = checker.load_rules(mode=mode)
    try:
        _, issues, stats = checker.analyze_text(text, mode, rules)
        report = checker.build_report(issues, stats, rules, examples=20)
        send(chat_id, report)
    except Exception:
        log.exception("direct text analysis failed")
        send(chat_id, "❌ در بررسی متن مشکلی پیش آمد؛ لطفاً دوباره تلاش کنید.")


# ---------------------------- مدیریت پیام‌ها و دستورات ----------------------------
def handle(update, pool):
    msg = update.get("message")
    if not msg:
        return
    chat_id = msg["chat"]["id"]
    text = (msg.get("text") or "").strip()

    if text.startswith("/start"):
        send(chat_id, WELCOME)
    elif text.startswith("/help"):
        send(chat_id, HELP)
    elif text.startswith("/thesis"):
        modes[chat_id] = "thesis"
        send(chat_id, "🎓 الگوی «پایان‌نامه» فعال شد. فایل Word، PDF یا متن خود را ارسال کنید.")
    elif text.startswith("/article"):
        modes[chat_id] = "article"
        send(chat_id, "📑 الگوی «مقاله علمی/پژوهشی» فعال شد. فایل Word، PDF یا متن خود را ارسال کنید.")
    elif text.startswith("/english"):
        modes[chat_id] = "english"
        send(chat_id, "🇬🇧 الگوی «متن/مقاله انگلیسی» فعال شد. متن یا فایل انگلیسی خود را ارسال فرمایید.")
    elif text.startswith("/status"):
        cur_mode = modes.get(chat_id, "thesis")
        title_map = {
            "thesis": "پایان‌نامه دانشگاهی 🎓",
            "article": "مقاله پژوهشی 📑",
            "english": "متن انگلیسی 🇬🇧"
        }
        send(chat_id, f"📌 حالت فعال شما: {title_map.get(cur_mode, cur_mode)}\nبرای تغییر می‌توانید از دستورات /thesis ، /article یا /english استفاده کنید.")
    elif msg.get("document"):
        with lock:
            if chat_id in busy:
                send(chat_id, "⏳ پردازش فایل قبلی شما در حال اجراست؛ لطفاً شکیبا باشید.")
                return
            wait = COOLDOWN - (time.time() - last_job.get(chat_id, 0))
            if wait > 0:
                send(chat_id, f"لطفاً {int(wait) + 1} ثانیه‌ی دیگر شکیبا باشید و سپس فایل را ارسال کنید.")
                return
            busy.add(chat_id)
            last_job[chat_id] = time.time()
        pool.submit(process_document, chat_id, msg["document"])
    elif text:
        # پردازش مستقیم متن ارسالی در صورتی که دستور نباشد
        process_text_message(chat_id, text)


import socket

def _start_health_server():
    """اجرای وب‌سرور سبک روی 0.0.0.0 برای تایید سلامت در هاست‌های ابری (مانند Render)."""
    port_str = os.environ.get("PORT")
    if not port_str:
        return
    try:
        port = int(port_str)
    except ValueError:
        port = 10000

    from http.server import HTTPServer, BaseHTTPRequestHandler

    class HealthHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(b'{"status":"ok","service":"virastar-bale-bot"}')

        def log_message(self, format, *args):
            return

    try:
        server = HTTPServer(("0.0.0.0", port), HealthHandler)
        t = threading.Thread(target=server.serve_forever, daemon=True)
        t.start()
        log.info("Health server listening on 0.0.0.0:%d", port)
    except Exception as e:
        log.warning("Could not start health server on port %d: %s", port, e)


def _acquire_lock_port(port=49250):
    """جلوگیری از اجرای همزمان چند نسخه از ربات با استفاده از پورت لوکال."""
    # روی هاست‌های ابری مانند Render نیازی به قفل پورت لوکال نیست
    if os.environ.get("RENDER") or os.environ.get("PORT"):
        return True
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
        return s
    except OSError:
        log.error("خطا: یک نسخه دیگر از ربات در حال حاضر در حال اجرا است! از اجرای همزمان چند ربات جلوگیری شد.")
        return None


def main():
    _start_health_server()
    lock_sock = _acquire_lock_port()
    if lock_sock is None:
        raise SystemExit("یک نسخه دیگر از ربات در پس‌زمینه در حال اجراست.")
    if not TOKEN:
        raise SystemExit("متغیر محیطی BALE_TOKEN تنظیم نشده است.")
    pool = ThreadPoolExecutor(max_workers=WORKERS)
    offset = None
    log.info("bot started as %s", call("getMe").get("username"))
    while True:
        try:
            payload = {"timeout": 30}
            if offset is not None:
                payload["offset"] = offset
            for upd in call("getUpdates", **payload):
                offset = upd["update_id"] + 1
                try:
                    handle(upd, pool)
                except Exception:
                    log.exception("handler error")
        except KeyboardInterrupt:
            break
        except Exception as e:
            log.warning("polling error: %s", e)
            time.sleep(5)


if __name__ == "__main__":
    main()
