from __future__ import annotations

import copy
import hashlib
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import inspect_docx, validate_docx_structure, verify_preservation
from .citations import _author_year_reference_key, build_citation_graph
from .docx_package import DocxPackage, NS, W_NS, serialize_xml_preserving_namespaces
from .ooxml_order import ensure_child
from .structure import extract_structure

ET.register_namespace("w", W_NS)

_SUPERSCRIPT_TRANS = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
_SUPERSCRIPT_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹"


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


@dataclass
class LinkResult:
    input: str
    output: str
    mode: str
    citation_style: str
    detection_confidence: int
    links_added: int
    reverse_links_added: int
    citation_bookmarks_added: int
    references_bookmarked: int
    skipped_complex_citations: int
    unresolved_citations: list[str]
    preservation: dict
    structural_validation: dict
    warnings: list[str]

    @property
    def passed(self) -> bool:
        return bool(self.preservation.get("passed")) and bool(self.structural_validation.get("passed"))

    def to_dict(self) -> dict:
        data = asdict(self)
        data["passed"] = self.passed
        return data


def _paragraph_text(p) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def _reference_key(text: str, fallback: int) -> str:
    m = re.match(r"^\s*(?:\[(\d+)\]|(\d+)[.)])\s*", text)
    return str(int(m.group(1) or m.group(2))) if m else str(fallback)


def _bookmark_token(prefix: str, key: str, suffix: str = "") -> str:
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:10]
    return f"_{prefix}{digest}{suffix}"[:39]


def _max_bookmark_id(root) -> int:
    ids = []
    for b in root.findall(".//w:bookmarkStart", NS):
        try:
            ids.append(int(b.attrib.get(_q("id"), "0")))
        except ValueError:
            pass
    return max(ids, default=0)


def _add_bookmark(p, bookmark_id: int, name: str) -> None:
    start = ET.Element(_q("bookmarkStart"), {_q("id"): str(bookmark_id), _q("name"): name})
    end = ET.Element(_q("bookmarkEnd"), {_q("id"): str(bookmark_id)})
    children = list(p)
    insert_at = 1 if children and children[0].tag == _q("pPr") else 0
    p.insert(insert_at, start)
    p.append(end)


def _make_run(fragment: str, source_run: ET.Element, *, hyperlink_style: bool = False) -> ET.Element:
    r = ET.Element(_q("r"))
    rpr = source_run.find("w:rPr", NS)
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    if hyperlink_style:
        hrpr = r.find("w:rPr", NS)
        if hrpr is None:
            hrpr = ET.Element(_q("rPr"))
            r.insert(0, hrpr)
        color = ensure_child(hrpr, "color")
        color.set(_q("val"), "0563C1")
        underline = ensure_child(hrpr, "u")
        underline.set(_q("val"), "single")
    t = ET.SubElement(r, _q("t"))
    if fragment.startswith(" ") or fragment.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = fragment
    return r


def _simple_run_text(run: ET.Element) -> str | None:
    text_nodes = run.findall("w:t", NS)
    if len(text_nodes) != 1:
        return None
    if run.find("w:drawing", NS) is not None or run.find("w:fldChar", NS) is not None or run.find("w:instrText", NS) is not None:
        return None
    return text_nodes[0].text or ""


def _numeric_spans(text: str, style: str) -> list[tuple[int, int, str]]:
    spans: list[tuple[int, int, str]] = []
    if style == "numeric-brackets":
        groups = re.finditer(r"\[(\d+(?:\s*(?:[-–—,;])\s*\d+)*)\]", text)
        for group in groups:
            inner_start = group.start(1)
            for hit in re.finditer(r"\d+", group.group(1)):
                spans.append((inner_start + hit.start(), inner_start + hit.end(), str(int(hit.group(0)))))
    elif style == "numeric-parentheses":
        groups = re.finditer(r"\((\d+(?:\s*(?:[-–—,;])\s*\d+)*)\)", text)
        for group in groups:
            if re.fullmatch(r"(?:19|20)\d{2}", group.group(1).strip()):
                continue
            inner_start = group.start(1)
            for hit in re.finditer(r"\d+", group.group(1)):
                spans.append((inner_start + hit.start(), inner_start + hit.end(), str(int(hit.group(0)))))
    elif style == "numeric-superscript":
        pattern = rf"[{_SUPERSCRIPT_DIGITS}]+(?:\s*(?:[-–—,;])\s*[{_SUPERSCRIPT_DIGITS}]+)*"
        for group in re.finditer(pattern, text):
            for hit in re.finditer(rf"[{_SUPERSCRIPT_DIGITS}]+", group.group(0)):
                raw = hit.group(0).translate(_SUPERSCRIPT_TRANS)
                spans.append((group.start() + hit.start(), group.start() + hit.end(), str(int(raw))))
    return spans


def _author_year_spans(text: str) -> list[tuple[int, int, str]]:
    surname = r"[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
    year = r"(?:18|19|20)\d{2}[a-z]?"
    candidates: list[tuple[int, int, str]] = []

    for hit in re.finditer(rf"\b({surname})(?:\s+et\s+al\.)?\s*\(\s*({year})\s*\)", text):
        candidates.append((hit.start(), hit.end(), f"{hit.group(1).casefold()}:{hit.group(2).casefold()}"))
    for hit in re.finditer(rf"\b({surname})(?:\s+et\s+al\.)?\s*,\s*({year})\b", text):
        candidates.append((hit.start(), hit.end(), f"{hit.group(1).casefold()}:{hit.group(2).casefold()}"))

    candidates.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    accepted: list[tuple[int, int, str]] = []
    end = -1
    for item in candidates:
        if item[0] >= end:
            accepted.append(item)
            end = item[1]
    return accepted


def _citation_spans(text: str, mode: str, style: str) -> list[tuple[int, int, str]]:
    if mode == "numbered":
        return _numeric_spans(text, style)
    if mode == "author-year":
        return _author_year_spans(text)
    return []


def _replace_run_with_navigation(
    p: ET.Element,
    run_index: int,
    run: ET.Element,
    ref_names: dict[str, str],
    *,
    mode: str,
    style: str,
    bookmark_id: int,
    occurrence_counts: dict[str, int],
    first_citation_names: dict[str, str],
) -> tuple[int, int, int, int]:
    text = _simple_run_text(run)
    if text is None:
        return 0, 1, 0, bookmark_id
    spans = [span for span in _citation_spans(text, mode, style) if span[2] in ref_names]
    if not spans:
        return 0, 0, 0, bookmark_id

    fragments: list[ET.Element] = []
    cursor = 0
    links = 0
    bookmarks = 0
    for start, end, key in spans:
        if start < cursor:
            continue
        if start > cursor:
            fragments.append(_make_run(text[cursor:start], run))

        occurrence_counts[key] = occurrence_counts.get(key, 0) + 1
        citation_name = _bookmark_token("WJMCCite", key, f"_{occurrence_counts[key]}")
        first_citation_names.setdefault(key, citation_name)
        start_node = ET.Element(_q("bookmarkStart"), {_q("id"): str(bookmark_id), _q("name"): citation_name})
        end_node = ET.Element(_q("bookmarkEnd"), {_q("id"): str(bookmark_id)})
        bookmark_id += 1
        bookmarks += 1

        hyperlink = ET.Element(_q("hyperlink"), {_q("anchor"): ref_names[key], _q("history"): "1"})
        hyperlink.append(_make_run(text[start:end], run, hyperlink_style=True))
        fragments.extend([start_node, hyperlink, end_node])
        links += 1
        cursor = end

    if cursor < len(text):
        fragments.append(_make_run(text[cursor:], run))

    p.remove(run)
    for offset, fragment in enumerate(fragments):
        p.insert(run_index + offset, fragment)
    return links, 0, bookmarks, bookmark_id


def _reverse_target_span(text: str, mode: str) -> tuple[int, int] | None:
    if not text:
        return None
    if mode == "numbered":
        m = re.match(r"^(\s*(?:\[\d+\]|\d+[.)]))", text)
        if m:
            return m.span(1)
    m = re.search(r"\b[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\b", text)
    return m.span(0) if m else None


def _add_reverse_link(p: ET.Element, anchor: str, mode: str) -> bool:
    for child in list(p):
        if child.tag != _q("r"):
            continue
        text = _simple_run_text(child)
        if text is None:
            continue
        span = _reverse_target_span(text, mode)
        if span is None:
            continue
        start, end = span
        children = list(p)
        run_index = children.index(child)
        fragments: list[ET.Element] = []
        if start:
            fragments.append(_make_run(text[:start], child))
        hyperlink = ET.Element(_q("hyperlink"), {_q("anchor"): anchor, _q("history"): "1"})
        hyperlink.append(_make_run(text[start:end], child, hyperlink_style=True))
        fragments.append(hyperlink)
        if end < len(text):
            fragments.append(_make_run(text[end:], child))
        p.remove(child)
        for offset, fragment in enumerate(fragments):
            p.insert(run_index + offset, fragment)
        return True
    return False


def link_plain_citations(input_path: str | Path, output_path: str | Path) -> LinkResult:
    src = Path(input_path)
    dst = Path(output_path)
    if src.resolve() == dst.resolve():
        raise ValueError("The navigable copy must use a different output path from the original manuscript.")

    package = DocxPackage(src)
    inventory = inspect_docx(src)
    warnings: list[str] = []
    if inventory.citation.total_candidate_fields:
        raise ValueError(
            "Live citation-manager fields were detected. The master manuscript will not be rewritten. Create a separate static linked review copy instead."
        )

    structure = extract_structure(src)
    if structure.reference_heading_index is None or not structure.reference_paragraphs:
        raise ValueError("Could not identify a References/Bibliography section with reference entries.")

    graph = build_citation_graph(src)
    if graph.mode not in {"numbered", "author-year"}:
        raise ValueError("A reliably linkable plain-text citation system was not detected.")

    allowed_keys = {item.reference_key for item in graph.links if item.matched and not item.ambiguous}
    if not allowed_keys:
        raise ValueError("No unambiguous citation-to-reference matches were available for linking.")

    root = package.xml("word/document.xml")
    body_paragraphs = root.findall(".//w:body/w:p", NS)
    original_document = package.read("word/document.xml")

    ref_names: dict[str, str] = {}
    ref_paragraphs: dict[str, ET.Element] = {}
    bookmark_id = _max_bookmark_id(root) + 1
    ref_count = 0
    seen_author_year: set[str] = set()

    for ordinal, rec in enumerate(structure.reference_paragraphs, start=1):
        if rec.index >= len(body_paragraphs):
            continue
        p = body_paragraphs[rec.index]
        if graph.mode == "numbered":
            key = _reference_key(rec.text, ordinal)
        else:
            key = _author_year_reference_key(rec.text)
            if key is None or key in seen_author_year:
                continue
            seen_author_year.add(key)
        if key not in allowed_keys:
            continue
        name = _bookmark_token("WJMCRef", key)
        ref_names[key] = name
        ref_paragraphs[key] = p
        if not any(b.attrib.get(_q("name")) == name for b in p.findall("w:bookmarkStart", NS)):
            _add_bookmark(p, bookmark_id, name)
            bookmark_id += 1
            ref_count += 1

    links_added = 0
    citation_bookmarks = 0
    skipped = 0
    occurrence_counts: dict[str, int] = {}
    first_citation_names: dict[str, str] = {}

    for idx, p in enumerate(body_paragraphs):
        if idx >= structure.reference_heading_index:
            break
        if p.find(".//w:instrText", NS) is not None or p.find(".//w:fldChar", NS) is not None:
            skipped += 1
            continue
        for child in list(p):
            if child.tag != _q("r"):
                continue
            current_children = list(p)
            if child not in current_children:
                continue
            run_index = current_children.index(child)
            added, sk, bookmarks, bookmark_id = _replace_run_with_navigation(
                p,
                run_index,
                child,
                ref_names,
                mode=graph.mode,
                style=graph.citation_style,
                bookmark_id=bookmark_id,
                occurrence_counts=occurrence_counts,
                first_citation_names=first_citation_names,
            )
            links_added += added
            skipped += sk
            citation_bookmarks += bookmarks

    reverse_links = 0
    for key, p in ref_paragraphs.items():
        target = first_citation_names.get(key)
        if target and _add_reverse_link(p, target, graph.mode):
            reverse_links += 1

    dst.parent.mkdir(parents=True, exist_ok=True)
    new_document = serialize_xml_preserving_namespaces(root, original_document)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_document if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)

    preservation = verify_preservation(src, dst).to_dict()
    structural = validate_docx_structure(dst) if preservation.get("passed") else {
        "passed": False,
        "checks": [],
        "failures": ["Structural validation was not run because preservation failed."],
        "warnings": [],
    }
    if not preservation.get("passed"):
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because the preservation audit failed.")
    elif not structural.get("passed"):
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because defensive OOXML structural validation failed.")

    unresolved = sorted(set(graph.unmatched_citations + graph.ambiguous_citations))
    if unresolved:
        warnings.append(
            f"{len(unresolved)} citation key(s) were left unlinked because they were unmatched or ambiguous. No guess was made."
        )
    if links_added == 0:
        warnings.append("No safe citation-to-reference hyperlinks were added.")
    if reverse_links:
        warnings.append("Reference backlinks return to the first linked in-text occurrence. Additional occurrences remain available through forward links and the citation map.")

    return LinkResult(
        input=str(src),
        output=str(dst),
        mode=graph.mode,
        citation_style=graph.citation_style,
        detection_confidence=graph.detection_confidence,
        links_added=links_added,
        reverse_links_added=reverse_links,
        citation_bookmarks_added=citation_bookmarks,
        references_bookmarked=ref_count,
        skipped_complex_citations=skipped,
        unresolved_citations=unresolved,
        preservation=preservation,
        structural_validation=structural,
        warnings=warnings,
    )


def link_plain_numbered_citations(input_path: str | Path, output_path: str | Path) -> LinkResult:
    """Backward-compatible entry point.

    The implementation now auto-detects supported plain-text citation styles,
    including numbered and author-year systems.
    """
    return link_plain_citations(input_path, output_path)
