from __future__ import annotations

import re
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .audit import inspect_docx
from .structure import extract_structure


@dataclass
class CitationLink:
    citation: str
    reference_key: str
    matched: bool
    reference_text: str | None = None


@dataclass
class CitationGraphReport:
    mode: str
    in_text_citation_count: int
    unique_citation_keys: int
    reference_count: int
    matched_links: int
    unmatched_citations: list[str]
    uncited_references: list[str]
    links: list[CitationLink]
    live_field_inventory: dict[str, Any]
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "in_text_citation_count": self.in_text_citation_count,
            "unique_citation_keys": self.unique_citation_keys,
            "reference_count": self.reference_count,
            "matched_links": self.matched_links,
            "unmatched_citations": self.unmatched_citations,
            "uncited_references": self.uncited_references,
            "links": [asdict(x) for x in self.links],
            "live_field_inventory": self.live_field_inventory,
            "warnings": self.warnings,
        }


def _expand_numeric_group(group: str) -> list[str]:
    keys: list[str] = []
    for token in re.split(r"\s*,\s*", group):
        token = token.strip()
        m = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)", token)
        if m:
            a, b = map(int, m.groups())
            if a <= b and b - a <= 100:
                keys.extend(str(i) for i in range(a, b + 1))
        elif token.isdigit():
            keys.append(str(int(token)))
    return keys


def _numeric_citations(text: str) -> list[str]:
    keys: list[str] = []
    for m in re.finditer(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]", text):
        keys.extend(_expand_numeric_group(m.group(1)))
    return keys


def _numbered_references(ref_texts: list[str]) -> dict[str, str]:
    refs: dict[str, str] = {}
    for idx, text in enumerate(ref_texts, start=1):
        m = re.match(r"^\s*(?:\[(\d+)\]|(\d+)[.)])\s*(.+)$", text)
        if m:
            key = str(int(m.group(1) or m.group(2)))
            refs[key] = text
        else:
            # Ordered Word lists often omit the visible list number from XML text.
            refs.setdefault(str(idx), text)
    return refs


def _author_year_reference_key(text: str) -> str | None:
    year = re.search(r"\b((?:19|20)\d{2}[a-z]?)\b", text)
    author = re.match(r"\s*([A-Z][A-Za-z'’\-]+)", text)
    if not year or not author:
        return None
    return f"{author.group(1).lower()}:{year.group(1).lower()}"


def _author_year_citations(text: str) -> list[str]:
    keys: list[str] = []
    # Conservative detection: surname + optional et al. + year inside parentheses or narrative text.
    for m in re.finditer(r"\b([A-Z][A-Za-z'’\-]+)(?:\s+et\s+al\.)?[, ]+\(?((?:19|20)\d{2}[a-z]?)\)?", text):
        keys.append(f"{m.group(1).lower()}:{m.group(2).lower()}")
    return keys


def build_citation_graph(path: str | Path) -> CitationGraphReport:
    inventory = inspect_docx(path)
    structure = extract_structure(path)
    body_text = "\n".join(
        p.text for p in structure.paragraphs
        if structure.reference_heading_index is None or p.index < structure.reference_heading_index
    )
    ref_texts = [p.text for p in structure.reference_paragraphs]
    warnings: list[str] = []

    # Live field mode is inventory-first. We do not flatten or reinterpret manager payloads.
    if inventory.citation.total_candidate_fields:
        warnings.append(
            "Live citation-manager fields detected. Word Journal Manuscript Converter preserves these fields and does not rewrite their payloads in automatic mode."
        )

    numeric_keys = _numeric_citations(body_text)
    numeric_refs = _numbered_references(ref_texts)
    if numeric_keys and numeric_refs:
        counts = Counter(numeric_keys)
        links = [
            CitationLink(citation=f"[{k}]", reference_key=k, matched=k in numeric_refs, reference_text=numeric_refs.get(k))
            for k in sorted(counts, key=lambda x: int(x))
        ]
        matched = sum(1 for x in links if x.matched)
        unmatched = [x.reference_key for x in links if not x.matched]
        uncited = [k for k in numeric_refs if k not in counts]
        return CitationGraphReport(
            mode="numbered",
            in_text_citation_count=len(numeric_keys),
            unique_citation_keys=len(counts),
            reference_count=len(numeric_refs),
            matched_links=matched,
            unmatched_citations=unmatched,
            uncited_references=uncited,
            links=links,
            live_field_inventory=asdict(inventory.citation),
            warnings=warnings,
        )

    ay_refs: dict[str, str] = {}
    for text in ref_texts:
        key = _author_year_reference_key(text)
        if key:
            ay_refs[key] = text
    ay_keys = _author_year_citations(body_text)
    if ay_keys and ay_refs:
        counts = Counter(ay_keys)
        links = [
            CitationLink(citation=k, reference_key=k, matched=k in ay_refs, reference_text=ay_refs.get(k))
            for k in sorted(counts)
        ]
        return CitationGraphReport(
            mode="author-year",
            in_text_citation_count=len(ay_keys),
            unique_citation_keys=len(counts),
            reference_count=len(ay_refs),
            matched_links=sum(1 for x in links if x.matched),
            unmatched_citations=[x.reference_key for x in links if not x.matched],
            uncited_references=[k for k in ay_refs if k not in counts],
            links=links,
            live_field_inventory=asdict(inventory.citation),
            warnings=warnings,
        )

    warnings.append("A reliable plain-text citation/reference pattern was not detected.")
    return CitationGraphReport(
        mode="live-fields" if inventory.citation.total_candidate_fields else "undetermined",
        in_text_citation_count=inventory.citation.total_candidate_fields,
        unique_citation_keys=inventory.citation.total_candidate_fields,
        reference_count=len(ref_texts),
        matched_links=0,
        unmatched_citations=[],
        uncited_references=[],
        links=[],
        live_field_inventory=asdict(inventory.citation),
        warnings=warnings,
    )
