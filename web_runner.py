# -*- coding: utf-8 -*-
"""
رانر ترکیبی (Web + Bot) برای استقرار روی هاست‌های ابری رایگان خارجی (مانند Render و Koyeb).
یک وب‌سرور سبک روی پورت پلتفرم اجرا می‌کند تا سرویس آنلاین بماند و با پینگ مانیتورینگ (UptimeRobot) به خواب نرود.
"""
import bot


if __name__ == "__main__":
    # bot.main owns the single health server and its lifecycle.
    bot.main()

