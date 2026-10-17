# -*- coding: utf-8 -*-
"""
رانر ترکیبی (Web + Bot) برای استقرار روی هاست‌های ابری رایگان خارجی (مانند Render و Koyeb).
یک وب‌سرور سبک روی پورت پلتفرم اجرا می‌کند تا سرویس آنلاین بماند و با پینگ مانیتورینگ (UptimeRobot) به خواب نرود.
"""
import os
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import bot


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(b'{"status":"ok","service":"virastar-bale-bot"}')

    def log_message(self, format, *args):
        # جلوگیری از شلوغ شدن لاگ با پینگ‌های مداوم
        return


def run_health_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), HealthHandler)
    server.serve_forever()


if __name__ == "__main__":
    # اجرای وب‌سرور سبک در ترد پس‌زمینه برای هاست‌های خارجی
    t = threading.Thread(target=run_health_server, daemon=True)
    t.start()
    
    # اجرای حلقه اصلی ربات بله
    bot.main()

