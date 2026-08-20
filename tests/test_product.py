from __future__ import annotations

from word_journal_manuscript_converter.branding import DISPLAY_VERSION, RELEASE_TAG
from word_journal_manuscript_converter.telemetry import _sanitized_properties
from word_journal_manuscript_converter.updates import ReleaseAsset, UpdateInfo, _version_key


def test_public_early_access_branding_is_simple():
    assert DISPLAY_VERSION == "Early Access"
    assert RELEASE_TAG == "v0.5.0-beta.1"


def test_telemetry_properties_are_allowlisted():
    data = _sanitized_properties(
        {
            "feature": "citation_navigation",
            "result": "success",
            "filename": "secret_manuscript.docx",
            "path": "C:/private/secret_manuscript.docx",
            "citation_text": "[1]",
        }
    )
    assert data == {"feature": "citation_navigation", "result": "success"}


def test_update_version_ordering():
    assert _version_key("v0.5.0-beta.2") > _version_key("v0.5.0-beta.1")
    assert _version_key("v0.5.0-rc.1") > _version_key("v0.5.0-beta.9")
    assert _version_key("v1.0.0") > _version_key("v0.5.0-rc.9")


def test_windows_installer_is_preferred_when_present(monkeypatch):
    monkeypatch.setattr("word_journal_manuscript_converter.updates.platform.system", lambda: "Windows")
    info = UpdateInfo(
        tag="v0.5.0-beta.2",
        name="Early Access Update",
        html_url="https://example.invalid/release",
        prerelease=True,
        published_at=None,
        assets=(
            ReleaseAsset("Word-Journal-Manuscript-Converter-Windows-x64.zip", "https://example.invalid/portable"),
            ReleaseAsset("Word-Journal-Manuscript-Converter-Setup.exe", "https://example.invalid/setup"),
        ),
    )
    assert info.preferred_asset() is not None
    assert info.preferred_asset().name == "Word-Journal-Manuscript-Converter-Setup.exe"
