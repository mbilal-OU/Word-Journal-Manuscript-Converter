from __future__ import annotations

import copy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .docx_package import DocxPackage, W_NS


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


# Known WordprocessingML child order for containers edited by this project.
# Unknown extension elements are deliberately ignored by the ordering checker.
_ORDER: dict[str, tuple[str, ...]] = {
    "sectPr": (
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
        "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
        "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
        "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
    ),
    "style": (
        "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
        "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
        "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
        "tblPr", "trPr", "tcPr", "tblStylePr",
    ),
    "pPr": (
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
        "pPrChange",
    ),
    "rPr": (
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
        "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
        "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
        "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath", "rPrChange",
    ),
}

_LEADING_PROPERTY = {
    "p": "pPr",
    "tbl": "tblPr",
    "tr": "trPr",
    "tc": "tcPr",
}


def _rank_map(parent: ET.Element) -> dict[str, int] | None:
    order = _ORDER.get(_local_name(parent.tag))
    if not order:
        return None
    return {name: index for index, name in enumerate(order)}


def insert_ordered(parent: ET.Element, child: ET.Element) -> ET.Element:
    """Insert a known child without violating the known WordprocessingML order."""
    parent_name = _local_name(parent.tag)
    child_name = _local_name(child.tag)

    if _LEADING_PROPERTY.get(parent_name) == child_name:
        parent.insert(0, child)
        return child

    ranks = _rank_map(parent)
    if not ranks or child_name not in ranks:
        parent.append(child)
        return child

    target_rank = ranks[child_name]
    children = list(parent)
    insert_at = len(children)
    for index, existing in enumerate(children):
        existing_rank = ranks.get(_local_name(existing.tag))
        if existing_rank is not None and existing_rank > target_rank:
            insert_at = index
            break
    parent.insert(insert_at, child)
    return child


def ensure_child(parent: ET.Element, local: str) -> ET.Element:
    existing = parent.find(f"w:{local}", {"w": W_NS})
    if existing is not None:
        return existing
    return insert_ordered(parent, ET.Element(_q(local)))


def replace_child_ordered(parent: ET.Element, local: str, source: ET.Element | None) -> bool:
    existing = parent.find(f"w:{local}", {"w": W_NS})
    if existing is None and source is None:
        return False
    if existing is not None:
        parent.remove(existing)
    if source is not None:
        insert_ordered(parent, copy.deepcopy(source))
    return True


def known_order_violations(root: ET.Element) -> list[dict[str, Any]]:
    violations: list[dict[str, Any]] = []
    for parent in root.iter():
        parent_name = _local_name(parent.tag)
        expected_leading = _LEADING_PROPERTY.get(parent_name)
        if expected_leading:
            prop = next((child for child in list(parent) if _local_name(child.tag) == expected_leading), None)
            if prop is not None and list(parent) and list(parent)[0] is not prop:
                violations.append(
                    {
                        "container": parent_name,
                        "previous": _local_name(list(parent)[0].tag),
                        "current": expected_leading,
                        "reason": "required leading property element is not first",
                    }
                )

        ranks = _rank_map(parent)
        if not ranks:
            continue
        previous_rank = -1
        previous_name: str | None = None
        for child in list(parent):
            name = _local_name(child.tag)
            rank = ranks.get(name)
            if rank is None:
                continue
            if rank < previous_rank:
                violations.append(
                    {
                        "container": parent_name,
                        "previous": previous_name,
                        "current": name,
                        "reason": "known child appears out of schema order",
                    }
                )
            else:
                previous_rank = rank
                previous_name = name
    return violations


def scan_docx_known_order(path: str | Path) -> dict[str, Any]:
    """Check known edited WordprocessingML containers across Word XML parts."""
    package = DocxPackage(path)
    violations: list[dict[str, Any]] = []
    for part in package.word_xml_parts():
        root = package.xml(part)
        for item in known_order_violations(root):
            item = dict(item)
            item["part"] = part
            violations.append(item)
    return {
        "passed": not violations,
        "violation_count": len(violations),
        "violations": violations,
        "scope": "known WordprocessingML child-order checks for containers edited by this project",
    }
