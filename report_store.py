"""Bounded, expiring in-memory reports, scoped to both chat and sender.

No original documents, filesystem paths or credentials are retained. A caller
gets immutable bytes, so expiry/deletion cannot race an in-flight delivery.
"""
from dataclasses import dataclass
import threading
import time


@dataclass(frozen=True)
class _Entry:
    expires: float
    payloads: tuple
    size: int


class ReportStore:
    def __init__(self, ttl_seconds=900, max_bytes=32 * 1024 * 1024,
                 max_entry_bytes=8 * 1024 * 1024, max_entries=64, clock=time.monotonic):
        if any(not isinstance(v, int) or isinstance(v, bool) or v <= 0
               for v in (ttl_seconds, max_bytes, max_entry_bytes, max_entries)):
            raise ValueError("Report limits must be positive integers")
        self.ttl_seconds = ttl_seconds
        self.max_bytes = max_bytes
        self.max_entry_bytes = min(max_entry_bytes, max_bytes)
        self.max_entries = max_entries
        self._clock = clock
        self._entries = {}
        self._bytes = 0
        self._lock = threading.Lock()
        self._closed = False

    def _drop(self, key):
        entry = self._entries.pop(key, None)
        if entry is not None:
            self._bytes -= entry.size

    def _expire(self):
        now = self._clock()
        for key, entry in list(self._entries.items()):
            if entry.expires <= now:
                self._drop(key)

    def put(self, chat_id, sender_id, payloads):
        if set(payloads) != {"html", "txt"} or any(not isinstance(v, bytes) for v in payloads.values()):
            raise ValueError("Only HTML and TXT report bytes may be retained")
        size = sum(map(len, payloads.values()))
        key = (chat_id, sender_id)
        with self._lock:
            self._expire()
            # A new result invalidates the old one even if it cannot be cached.
            self._drop(key)
            if self._closed or size > self.max_entry_bytes:
                return False
            while self._entries and (len(self._entries) >= self.max_entries or self._bytes + size > self.max_bytes):
                self._drop(next(iter(self._entries)))
            self._entries[key] = _Entry(self._clock() + self.ttl_seconds,
                                        tuple(payloads.items()), size)
            self._bytes += size
            return True

    def get(self, chat_id, sender_id):
        with self._lock:
            self._expire()
            entry = self._entries.get((chat_id, sender_id))
            return dict(entry.payloads) if entry is not None else None

    def delete(self, chat_id, sender_id):
        with self._lock:
            self._expire()
            found = (chat_id, sender_id) in self._entries
            self._drop((chat_id, sender_id))
            return found

    def reap(self):
        with self._lock:
            self._expire()

    def close(self):
        with self._lock:
            self._entries.clear()
            self._bytes = 0
            self._closed = True

