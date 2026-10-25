# -*- coding: utf-8 -*-
"""Student-facing bot interface: credit, menus, feedback and result summary."""
import json
import os
os.environ["BALE_SKIP_DOTENV"] = "1"
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import bot
import checker
from analysis_model import Issue


class NoPool:
    def submit(self, *args, **kwargs):
        raise AssertionError("A menu interaction must not start a background job")


class BotInterfaceTests(unittest.TestCase):
    def setUp(self):
        self.messages = []
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.feedback = Path(self.directory.name) / "feedback.jsonl"
        for name, value in (("modes", {}), ("awaiting_feedback", set()),
                            ("FEEDBACK_PATH", str(self.feedback)), ("ADMIN_CHAT_ID", "")):
            target = patch.object(bot, name, value)
            target.start()
            self.addCleanup(target.stop)
        target = patch.object(bot, "send", side_effect=lambda chat, text, **kw: self.messages.append((chat, text)) or True)
        target.start()
        self.addCleanup(target.stop)

    def say(self, text, chat=5):
        bot.handle({"message": {"chat": {"id": chat, "type": "private"}, "from": {"id": chat}, "text": text}}, NoPool())
        return self.messages[-1][1]

    def test_start_names_the_creator_and_hides_the_medical_profile(self):
        welcome = self.say("/start")
        self.assertIn("سازنده: سیداحمد موسوی — دانشگاه کاشان", welcome)
        self.assertIn("سازنده", self.say("/help"))
        for surface in (welcome, bot.HELP, json.dumps(bot.MAIN_KEYBOARD, ensure_ascii=False)):
            self.assertNotIn("tums", surface.lower())
            self.assertNotIn("علوم پزشکی", surface)

    def test_feedback_is_stored_without_user_identifiers_and_not_analysed(self):
        self.assertIn("/cancel", self.say(bot.FEEDBACK_BUTTON))
        self.assertIn("ممنون", self.say("کامنت spelling روی «کرونباخ» اشتباه بود."))
        record = json.loads(self.feedback.read_text(encoding="utf-8"))
        self.assertEqual("کامنت spelling روی «کرونباخ» اشتباه بود.", record["text"])
        self.assertEqual({"time", "mode", "version", "text"}, set(record))
        # The next text is a normal review request again (NoPool would raise).
        self.assertNotIn(5, bot.awaiting_feedback)

    def test_feedback_can_be_cancelled_and_forwarded_to_the_owner(self):
        self.say("/feedback")
        self.assertIn("لغو", self.say("/cancel"))
        self.assertFalse(self.feedback.exists())
        with patch.object(bot, "ADMIN_CHAT_ID", "99"):
            self.say("/feedback")
            self.say("پیشنهاد: پروفایل دانشگاه تهران اضافه شود.")
        self.assertTrue(any(chat == "99" and "پیشنهاد" in text for chat, text in self.messages))

    def test_mode_confirmation_says_what_to_do_next(self):
        reply = self.say("📑 مقاله پژوهشی")
        self.assertEqual("article", bot.modes[5])
        self.assertIn("بفرستید", reply)

    def test_summary_groups_findings_by_topic(self):
        issues = [Issue("plural_space"), Issue("spelling"), Issue("page_margin"), Issue("cite_missing_ref")]
        summary = checker.build_summary(issues, {"mode": "thesis"}, checker.load_rules(mode="thesis"))
        self.assertIn("4 مورد برای بازبینی", summary)
        self.assertIn("املا و نگارش: 2", summary)
        self.assertIn("ارجاع و منابع: 1", summary)
        self.assertIn("صفحه‌آرایی و قالب: 1", summary)


if __name__ == "__main__":
    unittest.main()
