from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .audit import inspect_docx
from .structure import extract_structure


_SUPERSCRIPT_TRANS = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
_SUPERSCRIPT_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹"


@dataclass
class CitationLink:
    citation: str
    reference_key: str
    matched: bool
    reference_text: str | None = None
    ambiguous: bool = False


@dataclass
class CitationGraphReport:
    mode: str
    citation_style: str
    detection_confidence: int
    in_text_citation_count: int
    unique_citation_keys: int
    reference_count: int
    matched_links: int
    unmatched_citations: list[str]
    ambiguous_citations: list[str]
    uncited_references: list[str]
    links: list[CitationLink]
    live_field_inventory: dict[str, Any]
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {
            "mode": self.mode,
            "citation_style": self.citation_style,
            "detection_confidence": self.detection_confidence,
            "in_text_citation_count": self.in_text_citation_count,
            "unique_citation_keys": self.unique_citation_keys,
            "reference_count": self.reference_count,
            "matched_links": self.matched_links,
            "unmatched_citations": self.unmatched_citations,
            "ambiguous_citations": self.ambiguous_citations,
            "uncited_references": self.uncited_references,
            "links": [asdict(x) for x in self.links],
            "live_field_inventory": self.live_field_inventory,
            "warnings": self.warnings,
        }


def _expand_numeric_group(group: str) -> list[str]:
    keys: list[str] = []
    for token in re.split(r"\s*[,;]\s*", group):
        token = token.strip()
        m = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", token)
        if m:
            a, b = map(int, m.groups())
            if a <= b and b - a <= 100:
                keys.extend(str(i) for i in range(a, b + 1))
        elif token.isdigit():
            keys.append(str(int(token)))
    return keys


def _numeric_bracket_citations(text: str) -> list[str]:
    keys: list[str] = []
    for m in re.finditer(r"\[(\d+(?:\s*(?:[-–—,;])\s*\d+)*)\]", text):
        keys.extend(_expand_numeric_group(m.group(1)))
    return keys


def _numeric_parenthetical_citations(text: str) -> list[str]:
    keys: list[str] = []
    for m in re.finditer(r"\((\d+(?:\s*(?:[-–—,;])\s*\d+)*)\)", text):
        group = m.group(1)
        if re.fullmatch(r"(?:19|20)\d{2}", group.strip()):
            continue
        keys.extend(_expand_numeric_group(group))
    return keys


def _superscript_to_ascii(token: str) -> str:
    return token.translate(_SUPERSCRIPT_TRANS)


def _numeric_superscript_citations(text: str) -> list[str]:
    keys: list[str] = []
    pattern = rf"([{_SUPERSCRIPT_DIGITS}]+(?:\s*(?:[-–—,;])\s*[{_SUPERSCRIPT_DIGITS}]+)*)"
    for m in re.finditer(pattern, text):
        keys.extend(_expand_numeric_group(_superscript_to_ascii(m.group(1))))
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
    year = re.search(r"\b((?:18|19|20)\d{2}[a-z]?)\b", text, re.IGNORECASE)
    author = re.match(r"\s*([A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", text)
    if not year or not author:
        return None
    return f"{author.group(1).casefold()}:{year.group(1).casefold()}"


def _author_year_reference_map(ref_texts: list[str]) -> dict[str, list[str]]:
    refs: dict[str, list[str]] = defaultdict(list)
    for text in ref_texts:
        key = _author_year_reference_key(text)
        if key:
            refs[key].append(text)
    return dict(refs)


def _author_year_citations(text: str) -> list[str]:
    keys: list[str] = []
    surname = r"[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
    year = r"(?:18|19|20)\d{2}[a-z]?"
    author_tail = rf"(?:\s+et\s+al\.)?(?:\s+(?:&|and)\s+{surname})?"

    # Parenthetical/grouped forms, including (Smith, 2024),
    # (Smith & Jones, 2024), and (Smith et al., 2024; Jones, 2023).
    for m in re.finditer(r"\(([^()]{3,300})\)", text):
        inner = m.group(1)
        for hit in re.finditer(rf"\b({surname}){author_tail}\s*,?\s*({year})\b", inner):
            keys.append(f"{hit.group(1).casefold()}:{hit.group(2).casefold()}")

    # Narrative forms such as Smith (2024), Smith and Jones (2024),
    # and Smith et al. (2024).
    for hit in re.finditer(rf"\b({surname}){author_tail}\s*\(\s*({year})\s*\)", text):
        keys.append(f"{hit.group(1).casefold()}:{hit.group(2).casefold()}")

    # Less punctuated forms sometimes used in prose or imported text: Smith, 2024.
    for hit in re.finditer(rf"\b({surname}){author_tail}\s*,\s*({year})\b", text):
        key = f"{hit.group(1).casefold()}:{hit.group(2).casefold()}"
        if key not in keys:
            keys.append(key)
    return keys


def _numeric_confidence(keys: list[str], refs: dict[str, str], style: str) -> int:
    if not keys or not refs:
        return 0
    matched = sum(key in refs for key in keys)
    ratio = matched / len(keys)
    evidence = min(12, len(keys) * 2)
    base = {
        "numeric-brackets": 70,
        "numeric-parentheses": 65,
        "numeric-superscript": 50,
    }.get(style, 55)
    score = round(base + 25 * ratio + evidence / 2)
    if style == "numeric-superscript" and len(keys) < 2:
        score = min(score, 60)
    return max(40, min(99, score))


def _author_year_confidence(keys: list[str], refs: dict[str, list[str]]) -> int:
    if not keys or not refs:
        return 0
    unique_matches = sum(len(refs.get(key, [])) == 1 for key in keys)
    ratio = unique_matches / len(keys)
    evidence = min(20, len(keys) * 2)
    return max(45, min(98, round(60 + 30 * ratio + evidence / 4)))


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
            "Live citation-manager fields detected. The master manuscript remains protected and citation-manager payloads are not rewritten automatically."
        )
        warnings.append(
            "Plain-text citation/reference mapping is only a secondary cross-check for this document."
        )

    numeric_refs = _numbered_references(ref_texts)
    numeric_candidates = {
        "numeric-brackets": _numeric_bracket_citations(body_text),
        "numeric-parentheses": _numeric_parenthetical_citations(body_text),
        "numeric-superscript": _numeric_superscript_citations(body_text),
    }
    numeric_style, numeric_keys = max(numeric_candidates.items(), key=lambda item: len(item[1]))
    numeric_conf = _numeric_confidence(numeric_keys, numeric_refs, numeric_style)

    ay_refs = _author_year_reference_map(ref_texts)
    ay_keys = _author_year_citations(body_text)
    ay_conf = _author_year_confidence(ay_keys, ay_refs)

    if numeric_keys and numeric_refs and numeric_conf >= max(65, ay_conf):
        counts = Counter(numeric_keys)
        links = [
            CitationLink(citation=k, reference_key=k, matched=k in numeric_refs, reference_text=numeric_refs.get(k))
            for k in sorted(counts, key=lambda x: int(x))
        ]
        unmatched = [x.reference_key for x in links if not x.matched]
        uncited = [k for k in numeric_refs if k not in counts]
        if unmatched and numeric_refs:
            max_cited = max((int(k) for k in counts), default=0)
            max_ref = max((int(k) for k in numeric_refs if k.isdigit()), default=0)
            if max_cited > max_ref:
                warnings.append(
                    f"Visible citations reach {max_cited} but extracted references reach only {max_ref}. "
                    "This can indicate inconsistent Word styles or an incomplete bibliography extraction."
                )
        if numeric_style == "numeric-superscript":
            warnings.append(
                "Superscript citation detection is intentionally conservative because scientific exponents can look similar. Review the citation map before creating a linked copy."
            )
        return CitationGraphReport(
            mode="numbered",
            citation_style=numeric_style,
            detection_confidence=numeric_conf,
            in_text_citation_count=len(numeric_keys),
            unique_citation_keys=len(counts),
            reference_count=len(numeric_refs),
            matched_links=sum(1 for x in links if x.matched),
            unmatched_citations=unmatched,
            ambiguous_citations=[],
            uncited_references=uncited,
            links=links,
            live_field_inventory=asdict(inventory.citation),
            warnings=warnings,
        )

    if ay_keys and ay_refs:
        counts = Counter(ay_keys)
        ambiguous = sorted(key for key in counts if len(ay_refs.get(key, [])) > 1)
        links: list[CitationLink] = []
        for key in sorted(counts):
            candidates = ay_refs.get(key, [])
            matched = len(candidates) == 1
            surname, year = key.split(":", 1)
            links.append(
                CitationLink(
                    citation=f"{surname.title()} {year}",
                    reference_key=key,
                    matched=matched,
                    reference_text=candidates[0] if matched else None,
                    ambiguous=len(candidates) > 1,
                )
            )
        if ambiguous:
            warnings.append(
                "Some author-year citations match more than one bibliography entry with the same first-author surname and year. These are left unresolved rather than guessed."
            )
        return CitationGraphReport(
            mode="author-year",
            citation_style="author-year",
            detection_confidence=ay_conf,
            in_text_citation_count=len(ay_keys),
            unique_citation_keys=len(counts),
            reference_count=len(ref_texts),
            matched_links=sum(1 for x in links if x.matched),
            unmatched_citations=[x.reference_key for x in links if not x.matched and not x.ambiguous],
            ambiguous_citations=ambiguous,
            uncited_references=[key for key in ay_refs if key not in counts],
            links=links,
            live_field_inventory=asdict(inventory.citation),
            warnings=warnings,
        )

    warnings.append("A reliable plain-text citation/reference pattern was not detected.")
    return CitationGraphReport(
        mode="live-fields" if inventory.citation.total_candidate_fields else "undetermined",
        citation_style="live-fields" if inventory.citation.total_candidate_fields else "undetermined",
        detection_confidence=100 if inventory.citation.total_candidate_fields else 0,
        in_text_citation_count=inventory.citation.total_candidate_fields,
        unique_citation_keys=inventory.citation.total_candidate_fields,
        reference_count=len(ref_texts),
        matched_links=0,
        unmatched_citations=[],
        ambiguous_citations=[],
        uncited_references=[],
        links=[],
        live_field_inventory=asdict(inventory.citation),
        warnings=warnings,
    )
