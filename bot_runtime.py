"""Small, dependency-free primitives for bounded bot work and safe configuration."""
from dataclasses import dataclass
import threading
import time
from urllib.parse import urlsplit


def env_integer(environ, name, default, minimum, maximum):
    try:
        value = int(environ.get(name, default))
    except (TypeError, ValueError):
        return default, f"{name} must be an integer"
    if not minimum <= value <= maximum:
        return default, f"{name} must be between {minimum} and {maximum}"
    return value, None


def validated_base_url(value, name):
    """Permit HTTPS origins, plus HTTP loopback for local testing/proxies."""
    try:
        parsed = urlsplit(value)
        _ = parsed.port
        valid = (parsed.hostname and not parsed.username and not parsed.password
                 and not parsed.query and not parsed.fragment
                 and (parsed.scheme == "https" or
                      (parsed.scheme == "http" and parsed.hostname in {"localhost", "127.0.0.1", "::1"})))
    except (TypeError, ValueError):
        valid = False
    if not valid:
        raise ValueError(f"{name} must be HTTPS (HTTP is allowed only on loopback)")
    return value.rstrip("/")


@dataclass(frozen=True)
class Reservation:
    chat_id: int
    sequence: int


class JobGate:
    """Bound running plus queued jobs; never hold this lock during user callbacks."""

    def __init__(self, capacity, cooldown, clock=time.monotonic):
        self.capacity = capacity
        self.cooldown = cooldown
        self.clock = clock
        self.lock = threading.Lock()
        self.busy = set()
        self.last_started = {}
        self._tokens = {}
        self._previous = {}
        self._sequence = 0
        self._accepting = True

    def reserve(self, chat_id):
        with self.lock:
            if not self._accepting:
                return None, "stopping", 0
            if chat_id in self.busy:
                return None, "busy", 0
            now = self.clock()
            wait = self.cooldown - (now - self.last_started.get(chat_id, -float("inf")))
            if wait > 0:
                return None, "cooldown", wait
            if len(self._tokens) >= self.capacity:
                return None, "full", 0
            self._sequence += 1
            token = Reservation(chat_id, self._sequence)
            self._previous[token] = self.last_started.get(chat_id)
            self.last_started[chat_id] = now
            self._tokens[chat_id] = token
            self.busy.add(chat_id)
            # Cooldown entries expire naturally; keep idle-user metadata bounded.
            if len(self.last_started) > max(1000, self.capacity * 4):
                expired = [cid for cid, stamp in self.last_started.items()
                           if cid not in self.busy and now - stamp >= self.cooldown]
                for cid in expired:
                    self.last_started.pop(cid, None)
            return token, None, 0

    def release(self, token, refund=False):
        with self.lock:
            if self._tokens.get(token.chat_id) != token:
                return
            self._tokens.pop(token.chat_id, None)
            self.busy.discard(token.chat_id)
            previous = self._previous.pop(token, None)
            if refund:
                if previous is None:
                    self.last_started.pop(token.chat_id, None)
                else:
                    self.last_started[token.chat_id] = previous

    def stop(self):
        with self.lock:
            self._accepting = False

    def snapshot(self):
        with self.lock:
            return {"pending": len(self._tokens), "capacity": self.capacity,
                    "accepting": self._accepting}


class ThreadLocalHTTP:
    """requests sessions are owned by one thread and closed after workers drain."""

    def __init__(self, factory):
        self.factory = factory
        self._local = threading.local()
        self._lock = threading.Lock()
        self._sessions = []

    def _session(self):
        if not hasattr(self._local, "session"):
            session = self.factory()
            self._local.session = session
            with self._lock:
                self._sessions.append(session)
        return self._local.session

    def get(self, *args, **kwargs):
        return self._session().get(*args, **kwargs)

    def post(self, *args, **kwargs):
        return self._session().post(*args, **kwargs)

    def close(self):
        with self._lock:
            sessions, self._sessions = self._sessions, []
        for session in sessions:
            session.close()
        self._local = threading.local()
