"""Bot admission, transport, lifecycle and command tests; all networking is mocked."""
import io
import os
os.environ["BALE_SKIP_DOTENV"] = "1"
import threading
import unittest
from concurrent.futures import Future
from unittest.mock import MagicMock, patch

import bot
from bot_runtime import JobGate, ThreadLocalHTTP, env_integer, validated_base_url


class RecordingPool:
    def __init__(self):
        self.calls = []

    def submit(self, function):
        future = Future()
        self.calls.append((function, future))
        return future

    def run(self, index=0):
        function, future = self.calls[index]
        if future.set_running_or_notify_cancel():
            try:
                future.set_result(function())
            except BaseException as exc:
                future.set_exception(exc)
        return future


class BotRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.clock = [100.0]
        self.gate = JobGate(2, 10, clock=lambda: self.clock[0])
        self.pool = RecordingPool()
        for name, value in (("jobs", self.gate), ("busy", self.gate.busy),
                            ("last_job", self.gate.last_started), ("lock", self.gate.lock),
                            ("modes", {}), ("_stop_event", threading.Event())):
            p = patch.object(bot, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.send = patch.object(bot, "send").start()
        self.addCleanup(patch.stopall)

    def update(self, chat=1, text=None, document=None):
        msg = {"chat": {"id": chat}}
        if text is not None:
            msg["text"] = text
        if document is not None:
            msg["document"] = document
        return {"message": msg}

    def test_text_runs_on_worker_and_keeps_original_spaces_and_mode(self):
        bot.modes[1] = "english"
        with patch.object(bot, "process_text_message") as analyze:
            bot.handle(self.update(text="  Research text  "), self.pool)
            analyze.assert_not_called()
            bot.modes[1] = "article"
            self.pool.run()
            analyze.assert_called_once_with(1, "  Research text  ", mode="english")
        self.assertNotIn(1, bot.busy)

    def test_document_captures_mode_and_annotation_flag(self):
        document = {"file_name": "sample.docx", "file_id": "test"}
        update = self.update(document=document)
        update["message"]["caption"] = "/annotate"
        bot.modes[1] = "tums"
        with patch.object(bot, "process_document") as analyze:
            bot.handle(update, self.pool)
            bot.modes[1] = "english"
            self.pool.run()
            analyze.assert_called_once_with(1, document, annotate_requested=True, mode="tums")

    def test_busy_protection_is_shared_by_text_and_document(self):
        with patch.object(bot, "process_text_message"):
            bot.handle(self.update(text="Text"), self.pool)
            bot.handle(self.update(document={"file_name": "x.pdf"}), self.pool)
        self.assertEqual(1, len(self.pool.calls))
        self.assertIn("قبلی", self.send.call_args.args[1])

    def test_global_capacity_blocks_a_third_chat(self):
        for chat in (1, 2, 3):
            bot.handle(self.update(chat, "Text"), self.pool)
        self.assertEqual(2, len(self.pool.calls))
        self.assertEqual(2, self.gate.snapshot()["pending"])
        self.assertIn("ظرفیت", self.send.call_args.args[1])

    def test_cooldown_and_monotonic_time(self):
        with patch.object(bot, "process_text_message"):
            bot.handle(self.update(text="Text"), self.pool)
            self.pool.run()
            bot.handle(self.update(text="Again"), self.pool)
            self.assertEqual(1, len(self.pool.calls))
            self.clock[0] += 11
            bot.handle(self.update(text="Later"), self.pool)
            self.assertEqual(2, len(self.pool.calls))

    def test_submit_failure_refunds_reservation_and_cooldown(self):
        pool = MagicMock()
        pool.submit.side_effect = RuntimeError("executor stopped")
        bot.handle(self.update(text="Text"), pool)
        self.assertEqual(0, self.gate.snapshot()["pending"])
        self.assertNotIn(1, bot.last_job)
        bot.handle(self.update(text="Retry"), self.pool)
        self.assertEqual(1, len(self.pool.calls))

    def test_task_failure_releases_reservation(self):
        with patch.object(bot, "process_text_message", side_effect=RuntimeError("synthetic")):
            bot.handle(self.update(text="Text"), self.pool)
            self.pool.run()
        self.assertFalse(bot.busy)
        self.assertEqual(0, self.gate.snapshot()["pending"])

    def test_cancelled_pending_job_refunds_reservation(self):
        bot.handle(self.update(text="Text"), self.pool)
        self.pool.calls[0][1].cancel()
        self.assertFalse(bot.busy)
        self.assertNotIn(1, bot.last_job)

    def test_busy_reply_occurs_outside_admission_lock(self):
        self.gate.reserve(1)
        def reply(*args, **kwargs):
            unlocked = self.gate.lock.acquire(blocking=False)
            self.assertTrue(unlocked)
            if unlocked:
                self.gate.lock.release()
        self.send.side_effect = reply
        bot.handle(self.update(text="Text"), self.pool)

    def test_commands_are_exact_and_unknown_commands_are_help(self):
        bot.handle(self.update(text="/articlex"), self.pool)
        self.assertNotIn(1, bot.modes)
        self.assertFalse(self.pool.calls)
        bot.handle(self.update(text="/tums"), self.pool)
        self.assertEqual("tums", bot.modes[1])
        bot.handle(self.update(text="/privacy"), self.pool)
        self.assertIn("موقت", self.send.call_args.args[1])

    def test_group_commands_addressed_to_other_bots_are_ignored(self):
        with patch.dict(bot._bot_status, {"username": "ResearchBot"}):
            bot.handle(self.update(text="/english@OtherBot"), self.pool)
            self.assertNotIn(1, bot.modes)
            bot.handle(self.update(text="/english@ResearchBot"), self.pool)
            self.assertEqual("english", bot.modes[1])

    def test_oversized_text_and_unsupported_files_do_not_reserve(self):
        with patch.object(bot, "MAX_TEXT_CHARS", 10):
            bot.handle(self.update(text="x" * 11), self.pool)
        bot.handle(self.update(document={"file_name": "macro.docm"}), self.pool)
        bot.handle(self.update(document={"file_name": "x.pdf", "file_size": "broken"}), self.pool)
        self.assertFalse(self.pool.calls)
        self.assertFalse(bot.busy)

    def test_readiness_requires_recent_successful_poll(self):
        with patch.object(bot, "_last_poll_monotonic", 100), patch.object(bot.time, "monotonic", return_value=101), patch.dict(bot._bot_status, {"started": True, "last_poll_error": None}):
            self.assertTrue(bot._health_payload(True)["ready"])
            bot._stop_event.set()
            self.assertFalse(bot._health_payload(True)["ready"])
        with patch.object(bot, "_last_poll_monotonic", 1), patch.object(bot.time, "monotonic", return_value=10000), patch.dict(bot._bot_status, {"started": True}):
            self.assertEqual("ok", bot._health_payload()["status"])
            self.assertFalse(bot._health_payload(True)["ready"])

    def test_health_identifies_deployed_revision_without_message_or_secrets(self):
        with patch.dict(os.environ, {"RENDER_GIT_COMMIT": "a" * 40,
                                     "BALE_TOKEN": "PRIVATE_TOKEN"}), patch.dict(
                bot._bot_status, {"last_message": {"text": "PRIVATE_MESSAGE"},
                                  "last_reply": {"chat_id": 12345}}):
            payload = bot._health_payload()
            self.assertEqual(bot.PRODUCT_VERSION, payload["version"])
            self.assertEqual("a" * 40, payload["revision"])
            self.assertNotIn("PRIVATE", str(payload))
            self.assertNotIn("12345", str(payload))
        with patch.dict(os.environ, {"RENDER_GIT_COMMIT": "PRIVATE_INVALID_VALUE"}):
            self.assertIsNone(bot._health_payload()["revision"])

    def test_health_routes_are_mocked_and_unknown_paths_404(self):
        with patch.dict(os.environ, {"PORT": "8080"}), patch.object(bot, "_health_server", None), patch("http.server.ThreadingHTTPServer") as server, patch.object(bot.threading, "Thread"):
            bot._start_health_server()
            bot._start_health_server()
            server.assert_called_once()
            handler_class = server.call_args.args[1]
            for path, expected in (("/health", 200), ("/ready", 503), ("/unknown", 404)):
                handler = object.__new__(handler_class)
                handler.path = path
                handler.send_response = MagicMock()
                handler.send_header = MagicMock()
                handler.end_headers = MagicMock()
                handler.send_error = MagicMock()
                handler.wfile = io.BytesIO()
                handler.do_GET()
                if expected == 404:
                    handler.send_error.assert_called_once_with(404)
                else:
                    handler.send_response.assert_called_once_with(expected)

    def test_api_checks_http_json_and_never_exposes_response_body(self):
        response = MagicMock(status_code=429)
        response.text = "PRIVATE_BODY_AND_TOKEN"
        with self.assertRaises(bot.BaleAPIError) as raised:
            bot._api_result(response, "sendMessage")
        self.assertNotIn("PRIVATE", str(raised.exception))
        response.status_code = 200
        response.json.return_value = {"ok": False, "description": "PRIVATE_BODY"}
        with self.assertRaises(bot.BaleAPIError):
            bot._api_result(response, "sendMessage")
        response.json.return_value = {"ok": True, "result": {"id": 1}}
        self.assertEqual({"id": 1}, bot._api_result(response, "sendMessage"))
        response.json.side_effect = ValueError("PRIVATE_BODY")
        with self.assertRaises(bot.BaleAPIError):
            bot._api_result(response, "sendMessage")

    def test_api_has_explicit_timeouts_no_redirects_and_no_retry(self):
        with patch.object(bot, "http") as http:
            response = http.post.return_value.__enter__.return_value
            response.status_code = 200
            response.json.return_value = {"ok": True, "result": True}
            bot.call("sendMessage", chat_id=1, text="synthetic")
            self.assertEqual((10, 45), http.post.call_args.kwargs["timeout"])
            self.assertFalse(http.post.call_args.kwargs["allow_redirects"])
            http.post.assert_called_once()

    def test_download_rejects_parent_paths_before_network(self):
        with patch.object(bot, "call", return_value={"file_path": "../private"}), patch.object(bot, "http") as http:
            with self.assertRaises(bot.BaleAPIError):
                bot.download("synthetic", "unused")
            http.get.assert_not_called()

    def test_rules_failure_is_handled_in_text_worker(self):
        with patch.object(bot.checker, "load_rules", side_effect=ValueError("synthetic")):
            bot.handle(self.update(text="Text"), self.pool)
            self.pool.run()
        self.assertFalse(bot.busy)
        self.assertIn("مشکلی", self.send.call_args.args[1])

    def test_keep_alive_disabled_by_default(self):
        with patch.dict(os.environ, {"KEEP_ALIVE_URL": ""}), patch.object(bot, "http") as http:
            bot._auto_keep_alive()
            http.get.assert_not_called()

    def test_main_gracefully_closes_workers_health_sessions_and_lock(self):
        def api(method, **kwargs):
            if method == "getMe":
                return {"username": "SyntheticBot"}
            bot._stop_event.set()
            return []
        with patch.object(bot, "TOKEN", "synthetic"), patch.object(bot, "_config_errors", []), patch.object(bot, "_acquire_lock_port") as lock_socket, patch.object(bot, "_start_health_server"), patch.object(bot, "_health_server") as health, patch.object(bot, "setup_commands"), patch.object(bot, "call", side_effect=api), patch.object(bot, "http") as http, patch.object(bot, "ThreadPoolExecutor") as executor, patch.object(bot.signal, "signal"), patch.dict(os.environ, {"KEEP_ALIVE_URL": ""}):
            bot.main()
            executor.return_value.shutdown.assert_called_once_with(wait=True, cancel_futures=True)
            lock_socket.return_value.close.assert_called_once()
            health.shutdown.assert_called_once()
            http.close.assert_called_once()
        self.assertFalse(self.gate.snapshot()["accepting"])


class RuntimePrimitiveTests(unittest.TestCase):
    def test_configuration_validation_has_safe_fallbacks(self):
        self.assertEqual((4, None), env_integer({}, "WORKERS", 4, 1, 32))
        for invalid in ("zero", "0", "999"):
            value, error = env_integer({"WORKERS": invalid}, "WORKERS", 4, 1, 32)
            self.assertEqual(4, value)
            self.assertIn("WORKERS", error)
        self.assertEqual("https://example.test", validated_base_url("https://example.test/", "API"))
        self.assertEqual("http://127.0.0.1:8080", validated_base_url("http://127.0.0.1:8080", "API"))
        for invalid in ("http://example.test", "https://secret@example.test", "https://example.test?token=private"):
            with self.assertRaises(ValueError):
                validated_base_url(invalid, "API")

    def test_stale_callback_cannot_release_a_new_reservation(self):
        gate = JobGate(1, 0)
        old, _, _ = gate.reserve(1)
        gate.release(old)
        new, _, _ = gate.reserve(1)
        gate.release(old)
        self.assertEqual(1, gate.snapshot()["pending"])
        gate.release(new)

    def test_sessions_are_thread_local_and_closed_after_use(self):
        created = []
        def factory():
            session = MagicMock()
            created.append(session)
            return session
        http = ThreadLocalHTTP(factory)
        http.get("first")
        http.post("second")
        thread = threading.Thread(target=lambda: http.get("third"))
        thread.start()
        thread.join(timeout=2)
        self.assertFalse(thread.is_alive())
        self.assertEqual(2, len(created))
        created[0].get.assert_called_once_with("first")
        created[0].post.assert_called_once_with("second")
        created[1].get.assert_called_once_with("third")
        http.close()
        for session in created:
            session.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()
