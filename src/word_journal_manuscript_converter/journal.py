from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

from .audit import inspect_docx
from .citations import build_citation_graph
from .profiles import load_profile_data, validate_profile_data
from .structure import extract_structure


@dataclass
class JournalProfile:
    journal: str
    article_type: str = "research-article"
    requirements: dict[str, Any] = field(default_factory=dict)
    source_url: str | None = None
    source_urls: list[str] | None = None
    checked_on: str | None = None
    notes: str | None = None
    profile_ref: str | None = None

    @classmethod
    def from_json(cls, source: str | Path) -> "JournalProfile":
        data, resolved = load_profile_data(source)
        issues = validate_profile_data(data)
        if issues:
            raise ValueError("Invalid journal profile: " + "; ".join(issues))
        profile = cls(**data)
        profile.profile_ref = resolved
        return profile

    def age_days(self, today: date | None = None) -> int | None:
        if not self.checked_on:
            return None
        try:
            checked = date.fromisoformat(self.checked_on)
        except ValueError:
            return None
        return ((today or date.today()) - checked).days


def readiness_check(docx_path: str | Path, profile_path: str | Path) -> dict[str, Any]:
    inventory = inspect_docx(docx_path)
    structure = extract_structure(docx_path)
    citations = build_citation_graph(docx_path)
    profile = JournalProfile.from_json(profile_path)
    req = profile.requirements
    checks: list[dict[str, Any]] = []

    def add(name: str, status: str, detail: str, auto_fixable: bool = False) -> None:
        checks.append({"check": name, "status": status, "detail": detail, "auto_fixable": auto_fixable})

    age = profile.age_days()
    if age is None:
        add("profile_provenance", "warn", "Profile has no valid checked_on date. Verify the official journal instructions before use.")
    elif age > 120:
        add("profile_freshness", "warn", f"Profile was checked {age} days ago. Re-verify its official source before submission.")
    else:
        add("profile_freshness", "pass", f"Profile checked {age} days ago.")

    if req.get("requires_live_citations"):
        n = inventory.citation.total_candidate_fields
        add("live_citations", "pass" if n else "warn", f"Detected {n} live citation-manager fields.")

    if req.get("requires_figures") is True:
        add("figures_present", "pass" if inventory.images else "warn", f"Detected {inventory.images} embedded media files.")

    if req.get("max_comments") is not None:
        limit = int(req["max_comments"])
        add("comments", "pass" if inventory.comments <= limit else "fail", f"Detected {inventory.comments}; allowed {limit}.")

    if req.get("tracked_changes_allowed") is False:
        n = inventory.tracked_insertions + inventory.tracked_deletions
        add("tracked_changes", "pass" if n == 0 else "fail", f"Detected {n} tracked-change elements.")

    if req.get("abstract_required"):
        n = structure.abstract_word_count
        add("abstract_present", "pass" if n else "fail", f"Detected abstract with {n} words.")

    if req.get("abstract_max_words") is not None:
        limit = int(req["abstract_max_words"])
        n = structure.abstract_word_count
        status = "pass" if n and n <= limit else "fail"
        add("abstract_word_limit", status, f"Abstract: {n}/{limit} words.")

    if req.get("abstract_recommended_max_words") is not None:
        limit = int(req["abstract_recommended_max_words"])
        n = structure.abstract_word_count
        status = "pass" if n and n <= limit else "warn"
        add("abstract_recommended_limit", status, f"Abstract: {n}/{limit} words against the journal's recommended target.")

    if req.get("keywords_min") is not None or req.get("keywords_max") is not None:
        n = len(structure.keywords)
        lo = int(req.get("keywords_min", 0))
        hi = int(req.get("keywords_max", 10**9))
        status = "pass" if lo <= n <= hi else "fail"
        add("keywords", status, f"Detected {n} keywords; expected {lo}–{hi}.")

    required_sections = [str(x) for x in req.get("required_sections", [])]
    normalized_headings = {re.sub(r"\s+", " ", h.lower().rstrip(":")) for h in structure.headings}
    methods = {"materials and methods", "materials & methods", "methods", "methodology", "online methods"}
    conflicts = {"competing interests", "conflict of interest", "conflicts of interest", "conflict of interests", "conflicts of interests", "competing interest"}
    aliases = {
        "materials and methods": methods,
        "methods": methods,
        "online methods": methods,
        "data availability": {"data availability", "data availability statement", "availability of data and materials"},
        "author contributions": {"author contributions", "author contribution", "author contribution statement", "contributions", "credit author statement", "credit authorship contribution statement"},
        "competing interests": conflicts,
        "conflicts of interest": conflicts,
        "author summary": {"author summary", "author summary statement"},
        "data summary": {"data summary", "data summary statement"},
        "impact statement": {"impact statement", "impact"},
        "importance": {"importance"},
        "funding information": {"funding information", "funding", "funding statement"},
    }
    for section in required_sections:
        norm = re.sub(r"\s+", " ", section.lower().rstrip(":"))
        accepted = aliases.get(norm, {norm})
        found = bool(normalized_headings.intersection(accepted))
        add(f"section:{section}", "pass" if found else "fail", f"Required section '{section}' {'found' if found else 'not found'}.")

    if req.get("citations_must_resolve"):
        unresolved = citations.unmatched_citations
        live_fields = int(citations.live_field_inventory.get("total_candidate_fields", 0) or 0)
        if unresolved and live_fields:
            add(
                "citation_reference_integrity",
                "warn",
                f"Citation mode={citations.mode}; plain-text cross-check found {len(unresolved)} unresolved keys, but {live_fields} live citation-manager fields are present. Verify in the citation manager before submission.",
            )
        else:
            add("citation_reference_integrity", "pass" if not unresolved else "fail", f"Citation mode={citations.mode}; unresolved citation keys={len(unresolved)}.")

    for key, label in (("margins_inches", "Page margins"), ("line_numbering", "Line numbering"), ("body_font", "Body font"), ("body_font_size_pt", "Body font size"), ("line_spacing", "Line spacing")):
        if key in req:
            add(f"format:{key}", "info", f"{label} target: {req[key]}", auto_fixable=True)

    scored = [x for x in checks if x["status"] in {"pass", "warn", "fail"} and not x["check"].startswith("profile_")]
    if scored:
        points = sum(1 if x["status"] == "pass" else 0.5 if x["status"] == "warn" else 0 for x in scored)
        score = round(100 * points / len(scored))
    else:
        score = 100

    return {
        "journal": profile.journal,
        "article_type": profile.article_type,
        "profile_ref": profile.profile_ref,
        "profile_source_url": profile.source_url,
        "profile_source_urls": profile.source_urls or ([profile.source_url] if profile.source_url else []),
        "profile_checked_on": profile.checked_on,
        "profile_age_days": age,
        "profile_notes": profile.notes,
        "readiness_score": score,
        "checks": checks,
        "inventory": inventory.to_dict(),
        "structure": {"word_count": structure.word_count, "abstract_word_count": structure.abstract_word_count, "keywords": structure.keywords, "headings": structure.headings},
        "citation_graph": citations.to_dict(),
        "note": "Only explicit rules in the selected profile are evaluated. Always verify the journal's current official author instructions before submission.",
    }
