from __future__ import annotations

import json
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ProfileDescriptor:
    key: str
    journal: str
    article_type: str
    source_url: str | None
    source_urls: list[str] | None
    checked_on: str | None
    notes: str | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "journal": self.journal,
            "article_type": self.article_type,
            "source_url": self.source_url,
            "source_urls": self.source_urls,
            "checked_on": self.checked_on,
            "notes": self.notes,
        }


def _bundled_root():
    return resources.files("word_journal_manuscript_converter").joinpath("bundled_profiles")


def bundled_profile_keys() -> list[str]:
    root = _bundled_root()
    return sorted(p.name[:-5] for p in root.iterdir() if p.name.endswith(".json"))


def _load_bundled_data(key: str) -> dict[str, Any]:
    filename = key if key.endswith(".json") else f"{key}.json"
    target = _bundled_root().joinpath(filename)
    if not target.is_file():
        raise ValueError(f"Unknown bundled profile '{key}'. Run 'word-journal-converter profiles' to list available profiles.")
    return json.loads(target.read_text(encoding="utf-8"))


def load_profile_data(source: str | Path) -> tuple[dict[str, Any], str]:
    path = Path(str(source)).expanduser()
    if path.exists():
        if path.suffix.lower() != ".json":
            raise ValueError("Journal profile must be a JSON file.")
        return json.loads(path.read_text(encoding="utf-8")), str(path.resolve())
    key = str(source).removesuffix(".json")
    return _load_bundled_data(key), f"bundled:{key}"


def validate_profile_data(data: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    if not isinstance(data, dict):
        return ["Profile root must be a JSON object."]
    if not str(data.get("journal", "")).strip():
        issues.append("Missing journal name.")
    if not str(data.get("article_type", "")).strip():
        issues.append("Missing article_type.")
    req = data.get("requirements")
    if not isinstance(req, dict):
        issues.append("requirements must be a JSON object.")
    source = data.get("source_url")
    if source is not None and not str(source).startswith(("https://", "http://")):
        issues.append("source_url must be an http(s) URL when present.")
    sources = data.get("source_urls")
    if sources is not None:
        if not isinstance(sources, list) or not all(str(x).startswith(("https://", "http://")) for x in sources):
            issues.append("source_urls must be a list of http(s) URLs when present.")
    checked = data.get("checked_on")
    if checked is not None:
        from datetime import date
        try:
            date.fromisoformat(str(checked))
        except ValueError:
            issues.append("checked_on must use YYYY-MM-DD.")
    return issues


def list_bundled_profiles() -> list[ProfileDescriptor]:
    result: list[ProfileDescriptor] = []
    for key in bundled_profile_keys():
        data = _load_bundled_data(key)
        result.append(
            ProfileDescriptor(
                key=key,
                journal=str(data.get("journal", key)),
                article_type=str(data.get("article_type", "research-article")),
                source_url=data.get("source_url"),
                source_urls=data.get("source_urls"),
                checked_on=data.get("checked_on"),
                notes=data.get("notes"),
            )
        )
    return result
