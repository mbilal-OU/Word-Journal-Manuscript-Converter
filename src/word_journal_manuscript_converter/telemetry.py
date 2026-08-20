from __future__ import annotations

import json
import os
import platform
import threading
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib import error, request

from . import __version__

# Publishable credentials are appropriate for public clients. Database RLS grants
# anonymous INSERT only and denies anonymous reads/updates/deletes.
SUPABASE_URL = "https://lrptpuiyenilpzrbjpmm.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_ztMbYB5-GWfALAsR4HNCTg_xDp_YgM_"
EVENTS_ENDPOINT = f"{SUPABASE_URL}/rest/v1/wjmc_events"
FEEDBACK_ENDPOINT = f"{SUPABASE_URL}/rest/v1/wjmc_feedback"

_ALLOWED_EVENT_PROPERTIES = {
    "feature",
    "channel",
    "result",
    "release_tag",
    "asset",
    "page",
    "action",
}


def _settings_dir() -> Path:
    if os.name == "nt":
        root = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        return root / "WordJournalManuscriptConverter"
    if platform.system() == "Darwin":
        return Path.home() / "Library" / "Application Support" / "WordJournalManuscriptConverter"
    return Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "word-journal-manuscript-converter"


class SettingsStore:
    def __init__(self) -> None:
        self.path = _settings_dir() / "settings.json"
        self.data: dict[str, Any] = self._load()
        if not self.data.get("install_id"):
            self.data["install_id"] = str(uuid.uuid4())
            self.save()

    def _load(self) -> dict[str, Any]:
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, OSError, json.JSONDecodeError, TypeError):
            return {}

    def save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        except OSError:
            pass

    @property
    def install_id(self) -> str:
        return str(self.data["install_id"])

    @property
    def analytics_consent(self) -> bool | None:
        value = self.data.get("analytics_consent")
        return value if isinstance(value, bool) else None

    def set_analytics_consent(self, value: bool) -> None:
        self.data["analytics_consent"] = bool(value)
        self.save()

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.data[key] = value
        self.save()


def _sanitized_properties(properties: dict[str, Any] | None) -> dict[str, Any]:
    clean: dict[str, Any] = {}
    for key, value in (properties or {}).items():
        if key not in _ALLOWED_EVENT_PROPERTIES:
            continue
        if isinstance(value, (str, int, float, bool)) or value is None:
            clean[key] = value if not isinstance(value, str) else value[:160]
    return clean


def _post_json(url: str, payload: dict[str, Any], timeout: float = 4.0) -> tuple[bool, str | None]:
    body = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    req = request.Request(
        url,
        data=body,
        method="POST",
        headers={
            "apikey": SUPABASE_PUBLISHABLE_KEY,
            "Content-Type": "application/json",
            "Prefer": "return=minimal",
            "User-Agent": f"WordJournalManuscriptConverter/{__version__}",
        },
    )
    try:
        with request.urlopen(req, timeout=timeout) as response:
            return 200 <= int(response.status) < 300, None
    except (error.URLError, error.HTTPError, TimeoutError, OSError) as exc:
        return False, str(exc)


def _background_post(url: str, payload: dict[str, Any]) -> None:
    threading.Thread(target=_post_json, args=(url, payload), daemon=True).start()


@dataclass
class FeedbackResult:
    sent: bool
    error: str | None = None


class TelemetryClient:
    """Privacy-minimized optional product analytics.

    Event properties are allow-listed. Manuscript content, filenames, paths,
    citation text, references, figures, document hashes, and document metadata
    are never accepted by this client.
    """

    def __init__(self, settings: SettingsStore | None = None, *, source: str = "desktop") -> None:
        self.settings = settings or SettingsStore()
        self.source = source
        self.session_id = str(uuid.uuid4())
        self.started_at = time.monotonic()
        self._last_heartbeat = self.started_at

    @property
    def enabled(self) -> bool:
        return self.settings.analytics_consent is True

    def track(
        self,
        event_name: str,
        *,
        properties: dict[str, Any] | None = None,
        duration_seconds: int | None = None,
    ) -> None:
        if not self.enabled:
            return
        payload: dict[str, Any] = {
            "source": self.source,
            "event_name": event_name,
            "anonymous_id": self.settings.install_id,
            "session_id": self.session_id,
            "app_version": __version__,
            "platform": f"{platform.system()} {platform.release()}"[:80],
            "properties": _sanitized_properties(properties),
        }
        if duration_seconds is not None:
            payload["duration_seconds"] = max(0, min(int(duration_seconds), 86400))
        _background_post(EVENTS_ENDPOINT, payload)

    def track_feature(self, feature: str, *, result: str | None = None) -> None:
        props: dict[str, Any] = {"feature": feature[:160]}
        if result:
            props["result"] = result[:160]
        self.track("feature_used", properties=props)

    def heartbeat_if_due(self, interval_seconds: int = 300) -> None:
        if not self.enabled:
            return
        now = time.monotonic()
        if now - self._last_heartbeat >= interval_seconds:
            self._last_heartbeat = now
            self.track("session_heartbeat", duration_seconds=int(now - self.started_at))

    def close(self) -> None:
        if not self.enabled:
            return
        payload: dict[str, Any] = {
            "source": self.source,
            "event_name": "session_end",
            "anonymous_id": self.settings.install_id,
            "session_id": self.session_id,
            "app_version": __version__,
            "platform": f"{platform.system()} {platform.release()}"[:80],
            "duration_seconds": max(0, min(int(time.monotonic() - self.started_at), 86400)),
            "properties": {},
        }
        _post_json(EVENTS_ENDPOINT, payload, timeout=1.5)

    def submit_feedback(
        self,
        *,
        rating: int | None,
        category: str,
        message: str,
        contact_email: str | None = None,
        consent_to_contact: bool = False,
    ) -> FeedbackResult:
        message = message.strip()
        if not message:
            return FeedbackResult(False, "Feedback message is required.")
        payload: dict[str, Any] = {
            "source": self.source,
            "anonymous_id": self.settings.install_id if self.enabled else None,
            "app_version": __version__,
            "platform": f"{platform.system()} {platform.release()}"[:80],
            "rating": rating,
            "category": category,
            "message": message[:4000],
            "contact_email": (contact_email or "").strip()[:320] or None,
            "consent_to_contact": bool(consent_to_contact and contact_email),
        }
        sent, err = _post_json(FEEDBACK_ENDPOINT, payload, timeout=6.0)
        return FeedbackResult(sent, err)
