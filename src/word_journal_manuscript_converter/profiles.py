from __future__ import annotations

import json
from copy import deepcopy
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any

from .journal_catalog import BUILTIN_PROFILE_CATALOG


# Frozen desktop builds must still start even if a packager drops package-data files.
# Keep the original shipped profile embedded as a fail-safe. The larger source-dated
# catalog lives in journal_catalog.py and is also available in frozen builds.
_EMBEDDED_PROFILES: dict[str, dict[str, Any]] = {
    "generic-review-copy": {
        "journal": "Generic review-copy profile",
        "article_type": "research-article",
        "checked_on": "2026-08-19",
        "notes": "Demonstration profile only. It is not an official journal specification.",
        "requirements": {
            "abstract_required": True,
            "tracked_changes_allowed": False,
            "citations_must_resolve": True,
            "margins_inches": {"top": 1.0, "right": 1.0, "bottom": 1.0, "left": 1.0},
            "body_font": "Times New Roman",
            "body_font_size_pt": 12,
            "line_spacing": 2.0,
            "line_numbering": {"count_by": 1, "restart": "continuous"},
        },
    },
}


def _catalog() -> dict[str, dict[str, Any]]:
    return {**_EMBEDDED_PROFILES, **BUILTIN_PROFILE_CATALOG}


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
    file_keys: set[str] = set()
    try:
        root = _bundled_root()
        file_keys = {p.name[:-5] for p in root.iterdir() if p.name.endswith(".json")}
    except (FileNotFoundError, ModuleNotFoundError, OSError, TypeError):
        pass
    return sorted(file_keys | set(_catalog()))


def _load_bundled_data(key: str) -> dict[str, Any]:
    normalized = key.removesuffix(".json")
    filename = f"{normalized}.json"
    try:
        target = _bundled_root().joinpath(filename)
        if target.is_file():
            return json.loads(target.read_text(encoding="utf-8"))
    except (FileNotFoundError, ModuleNotFoundError, OSError, TypeError):
        pass

    embedded = _catalog().get(normalized)
    if embedded is not None:
        return deepcopy(embedded)
    raise ValueError(
        f"Unknown bundled profile '{key}'. Run 'word-journal-converter profiles' to list available profiles."
    )


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
