from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

from .docx_package import DocxPackage, NS, W_NS


@dataclass
class ParagraphRecord:
    index: int
    text: str
    style: str | None = None
    is_heading: bool = False


@dataclass
class ManuscriptStructure:
    paragraphs: list[ParagraphRecord]
    headings: list[str]
    abstract_text: str
    keywords: list[str]
    reference_heading_index: int | None
    reference_paragraphs: list[ParagraphRecord]

    @property
    def word_count(self) -> int:
        return len(re.findall(r"\b[\w'’-]+\b", " ".join(p.text for p in self.paragraphs)))

    @property
    def abstract_word_count(self) -> int:
        return len(re.findall(r"\b[\w'’-]+\b", self.abstract_text))

    def to_dict(self) -> dict[str, Any]:
        return {
            "paragraphs": [asdict(p) for p in self.paragraphs],
            "headings": self.headings,
            "abstract_text": self.abstract_text,
            "abstract_word_count": self.abstract_word_count,
            "keywords": self.keywords,
            "reference_heading_index": self.reference_heading_index,
            "reference_paragraphs": [asdict(p) for p in self.reference_paragraphs],
            "word_count": self.word_count,
        }


def _paragraph_text(p) -> str:
    chunks: list[str] = []
    for node in p.iter():
        if node.tag == f"{{{W_NS}}}t" and node.text:
            chunks.append(node.text)
        elif node.tag == f"{{{W_NS}}}tab":
            chunks.append("\t")
        elif node.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
            chunks.append("\n")
    return "".join(chunks).strip()


def _style_id(p) -> str | None:
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.attrib.get(f"{{{W_NS}}}val")


def _looks_heading(text: str, style: str | None) -> bool:
    if style and re.match(r"(?i)^heading\s*\d*$", style.replace("_", " ")):
        return True
    # Conservative fallback: short, non-sentence manuscript headings.
    normalized = text.strip()
    if not normalized or len(normalized) > 100:
        return False
    common = {
        "abstract", "introduction", "background", "methods", "materials and methods",
        "results", "discussion", "conclusion", "conclusions", "references", "bibliography",
        "acknowledgments", "acknowledgements", "funding", "data availability",
        "author contributions", "competing interests", "conflict of interest",
        "conflicts of interest", "ethics approval", "supplementary material",
    }
    return normalized.lower().rstrip(":") in common


def extract_structure(path: str | Path) -> ManuscriptStructure:
    package = DocxPackage(path)
    root = package.xml("word/document.xml")
    records: list[ParagraphRecord] = []
    headings: list[str] = []

    for i, p in enumerate(root.findall(".//w:body/w:p", NS)):
        text = _paragraph_text(p)
        style = _style_id(p)
        is_heading = _looks_heading(text, style)
        rec = ParagraphRecord(index=i, text=text, style=style, is_heading=is_heading)
        records.append(rec)
        if text and is_heading:
            headings.append(text)

    # Abstract extraction: start at an Abstract heading and stop at next heading.
    abstract_text = ""
    abstract_idx = next((r.index for r in records if r.text.lower().rstrip(":") == "abstract"), None)
    if abstract_idx is not None:
        parts: list[str] = []
        for rec in records:
            if rec.index <= abstract_idx:
                continue
            if rec.is_heading:
                break
            if rec.text:
                parts.append(rec.text)
        abstract_text = " ".join(parts).strip()
    else:
        # Common journal template variant: a paragraph beginning with "Abstract".
        for rec in records:
            m = re.match(r"(?is)^abstract\s*[:—-]\s*(.+)$", rec.text)
            if m:
                abstract_text = m.group(1).strip()
                break

    keywords: list[str] = []
    for rec in records:
        m = re.match(r"(?is)^key\s*words?\s*[:—-]\s*(.+)$", rec.text)
        if m:
            keywords = [x.strip() for x in re.split(r"[;,]", m.group(1)) if x.strip()]
            break

    ref_idx = next(
        (r.index for r in records if r.text.lower().rstrip(":") in {"references", "bibliography"}),
        None,
    )
    refs: list[ParagraphRecord] = []
    if ref_idx is not None:
        for rec in records:
            if rec.index <= ref_idx:
                continue
            # Stop only on a later recognized heading, not a styled reference entry.
            if rec.is_heading and rec.text.lower().rstrip(":") not in {"references", "bibliography"}:
                break
            if rec.text:
                refs.append(rec)

    return ManuscriptStructure(
        paragraphs=records,
        headings=headings,
        abstract_text=abstract_text,
        keywords=keywords,
        reference_heading_index=ref_idx,
        reference_paragraphs=refs,
    )
