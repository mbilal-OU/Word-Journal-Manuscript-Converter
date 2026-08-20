from __future__ import annotations

import json
import os
import platform
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from urllib import error, request

from .branding import RELEASE_TAG, RELEASES_URL

API_RELEASES = "https://api.github.com/repos/mbilal-OU/Word-Journal-Manuscript-Converter/releases?per_page=30"


@dataclass(frozen=True)
class ReleaseAsset:
    name: str
    browser_download_url: str
    size: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "browser_download_url": self.browser_download_url,
            "size": self.size,
        }


@dataclass(frozen=True)
class UpdateInfo:
    tag: str
    name: str
    html_url: str
    prerelease: bool
    published_at: str | None
    assets: tuple[ReleaseAsset, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "tag": self.tag,
            "name": self.name,
            "html_url": self.html_url,
            "prerelease": self.prerelease,
            "published_at": self.published_at,
            "assets": [asset.to_dict() for asset in self.assets],
        }

    def preferred_asset(self) -> ReleaseAsset | None:
        system = platform.system()
        if system == "Windows":
            preferred = [
                "Word-Journal-Manuscript-Converter-Setup.exe",
                "Word-Journal-Manuscript-Converter-Windows-x64.zip",
            ]
        elif system == "Darwin":
            preferred = ["Word-Journal-Manuscript-Converter-macOS.zip"]
        else:
            preferred = ["Word-Journal-Manuscript-Converter-Linux-x64.tar.gz"]
        for wanted in preferred:
            for asset in self.assets:
                if asset.name == wanted:
                    return asset
        return None


def _version_key(tag: str) -> tuple[int, int, int, int, int]:
    text = tag.lower().strip().lstrip("v")
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)(?:[-.]?(alpha|a|beta|b|rc)[.-]?(\d+)?)?$", text)
    if not match:
        return (0, 0, 0, -1, 0)
    major, minor, patch = (int(match.group(i)) for i in (1, 2, 3))
    stage = match.group(4)
    number = int(match.group(5) or 0)
    rank = {None: 4, "rc": 3, "beta": 2, "b": 2, "alpha": 1, "a": 1}[stage]
    return (major, minor, patch, rank, number)


def fetch_releases(timeout: float = 5.0) -> list[dict[str, Any]]:
    req = request.Request(
        API_RELEASES,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "WordJournalManuscriptConverter-UpdateChecker",
        },
    )
    with request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def check_for_update(current_tag: str = RELEASE_TAG, timeout: float = 5.0) -> UpdateInfo | None:
    current = _version_key(current_tag)
    try:
        releases = fetch_releases(timeout=timeout)
    except (error.URLError, error.HTTPError, TimeoutError, OSError, json.JSONDecodeError):
        return None

    candidates = [
        release
        for release in releases
        if not release.get("draft")
        and _version_key(str(release.get("tag_name", ""))) > current
    ]
    if not candidates:
        return None

    newest = max(candidates, key=lambda r: _version_key(str(r.get("tag_name", ""))))
    assets = tuple(
        ReleaseAsset(
            name=str(asset.get("name", "")),
            browser_download_url=str(asset.get("browser_download_url", "")),
            size=int(asset.get("size") or 0),
        )
        for asset in newest.get("assets", [])
        if asset.get("name") and asset.get("browser_download_url")
    )
    return UpdateInfo(
        tag=str(newest.get("tag_name", "")),
        name=str(newest.get("name") or newest.get("tag_name") or "Update"),
        html_url=str(newest.get("html_url") or RELEASES_URL),
        prerelease=bool(newest.get("prerelease")),
        published_at=newest.get("published_at"),
        assets=assets,
    )


def download_asset(
    asset: ReleaseAsset,
    *,
    progress: Callable[[int, int], None] | None = None,
    timeout: float = 30.0,
) -> Path:
    target_dir = Path(tempfile.gettempdir()) / "WordJournalManuscriptConverter" / "updates"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / asset.name
    req = request.Request(
        asset.browser_download_url,
        headers={"User-Agent": "WordJournalManuscriptConverter-Updater"},
    )
    with request.urlopen(req, timeout=timeout) as response, target.open("wb") as handle:
        total = int(response.headers.get("Content-Length") or asset.size or 0)
        downloaded = 0
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            handle.write(chunk)
            downloaded += len(chunk)
            if progress:
                progress(downloaded, total)
    return target


def launch_downloaded_update(path: str | Path) -> bool:
    path = Path(path)
    if not path.exists():
        return False
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
            return True
        if system == "Darwin":
            subprocess.Popen(["open", str(path)])
            return True
        subprocess.Popen(["xdg-open", str(path.parent)])
        return True
    except (OSError, subprocess.SubprocessError):
        return False
