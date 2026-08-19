from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import inspect_docx, verify_preservation
from .docx_package import DocxPackage, NS, W_NS
from .structure import extract_structure

ET.register_namespace("w", W_NS)


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


@dataclass
class LinkResult:
    input: str
    output: str
    links_added: int
    references_bookmarked: int
    skipped_complex_citations: int
    preservation: dict
    warnings: list[str]

    @property
    def passed(self) -> bool:
        return bool(self.preservation.get("passed"))

    def to_dict(self) -> dict:
        data = asdict(self)
        data["passed"] = self.passed
        return data


def _paragraph_text(p) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def _reference_key(text: str, fallback: int) -> str:
    m = re.match(r"^\s*(?:\[(\d+)\]|(\d+)[.)])\s*", text)
    return str(int(m.group(1) or m.group(2))) if m else str(fallback)


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
    p.insert(0, start)
    p.append(end)


def _replace_run_with_hyperlinks(p, run_index: int, run, ref_names: dict[str, str]) -> tuple[int, int]:
    """Replace a simple text run with text/hyperlink fragments.

    Only simple runs are touched. Run properties are cloned so visible formatting
    is preserved. Complex fields and drawings are deliberately left alone.
    """
    text_nodes = run.findall("w:t", NS)
    if len(text_nodes) != 1:
        return 0, 0
    if run.find("w:drawing", NS) is not None or run.find("w:fldChar", NS) is not None or run.find("w:instrText", NS) is not None:
        return 0, 1
    text = text_nodes[0].text or ""
    matches = list(re.finditer(r"\[(\d+)\]", text))
    actionable = [m for m in matches if m.group(1) in ref_names]
    if not actionable:
        return 0, 0

    fragments: list[ET.Element] = []
    cursor = 0
    rpr = run.find("w:rPr", NS)

    def make_run(fragment: str) -> ET.Element:
        r = ET.Element(_q("r"))
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = ET.SubElement(r, _q("t"))
        if fragment.startswith(" ") or fragment.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = fragment
        return r

    links = 0
    for m in actionable:
        if m.start() > cursor:
            fragments.append(make_run(text[cursor:m.start()]))
        hyperlink = ET.Element(_q("hyperlink"), {_q("anchor"): ref_names[m.group(1)], _q("history"): "1"})
        hr = make_run(m.group(0))
        hrpr = hr.find("w:rPr", NS)
        if hrpr is None:
            hrpr = ET.Element(_q("rPr"))
            hr.insert(0, hrpr)
        color = ET.SubElement(hrpr, _q("color"))
        color.set(_q("val"), "0563C1")
        underline = ET.SubElement(hrpr, _q("u"))
        underline.set(_q("val"), "single")
        hyperlink.append(hr)
        fragments.append(hyperlink)
        links += 1
        cursor = m.end()
    if cursor < len(text):
        fragments.append(make_run(text[cursor:]))

    p.remove(run)
    for offset, fragment in enumerate(fragments):
        p.insert(run_index + offset, fragment)
    return links, 0


def link_plain_numbered_citations(input_path: str | Path, output_path: str | Path) -> LinkResult:
    src = Path(input_path)
    dst = Path(output_path)
    package = DocxPackage(src)
    inventory = inspect_docx(src)
    warnings: list[str] = []
    if inventory.citation.total_candidate_fields:
        raise ValueError(
            "Live citation-manager fields were detected. Word Journal Manuscript Converter will not overlay plain-text links on live EndNote/Zotero/CSL citations."
        )

    structure = extract_structure(src)
    if structure.reference_heading_index is None or not structure.reference_paragraphs:
        raise ValueError("Could not identify a References/Bibliography section with reference entries.")

    root = package.xml("word/document.xml")
    body_paragraphs = root.findall(".//w:body/w:p", NS)

    ref_names: dict[str, str] = {}
    bookmark_id = _max_bookmark_id(root) + 1
    ref_count = 0
    for ordinal, rec in enumerate(structure.reference_paragraphs, start=1):
        if rec.index >= len(body_paragraphs):
            continue
        p = body_paragraphs[rec.index]
        key = _reference_key(rec.text, ordinal)
        name = f"_WJMCRef{key}"
        ref_names[key] = name
        if not any(b.attrib.get(_q("name")) == name for b in p.findall("w:bookmarkStart", NS)):
            _add_bookmark(p, bookmark_id, name)
            bookmark_id += 1
            ref_count += 1

    links_added = 0
    skipped = 0
    for idx, p in enumerate(body_paragraphs):
        if idx >= structure.reference_heading_index:
            break
        if p.find(".//w:instrText", NS) is not None or p.find(".//w:fldChar", NS) is not None:
            continue

        # Iterate over a snapshot of original children. Recalculate each run's
        # current index before replacing it so inserting one hyperlink cannot
        # skip a later citation run in the same paragraph.
        for child in list(p):
            if child.tag != _q("r"):
                continue
            current_children = list(p)
            if child not in current_children:
                continue
            run_index = current_children.index(child)
            added, sk = _replace_run_with_hyperlinks(p, run_index, child, ref_names)
            links_added += added
            skipped += sk

    dst.parent.mkdir(parents=True, exist_ok=True)
    new_document = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_document if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)

    preservation = verify_preservation(src, dst).to_dict()
    if not preservation["passed"]:
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because the preservation audit failed.")
    if links_added == 0:
        warnings.append("No simple [N] citation tokens were linkable. Complex ranges/groups are intentionally left unchanged in this release.")

    return LinkResult(
        input=str(src), output=str(dst), links_added=links_added,
        references_bookmarked=ref_count, skipped_complex_citations=skipped,
        preservation=preservation, warnings=warnings,
    )
