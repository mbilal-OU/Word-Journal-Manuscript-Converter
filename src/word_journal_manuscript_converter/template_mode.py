from __future__ import annotations

import copy
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import verify_preservation
from .docx_package import DocxPackage, NS, W_NS

ET.register_namespace("w", W_NS)

SAFE_STYLE_IDS = (
    "Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Heading4",
    "Caption", "Quote", "IntenseQuote", "ListParagraph", "Bibliography",
)
TEMPLATE_SUFFIXES = {".docx", ".dotx"}


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _validate_template(path: str | Path) -> Path:
    p = Path(path)
    if not p.exists():
        raise ValueError(f"Template file not found: {p}")
    if p.suffix.lower() not in TEMPLATE_SUFFIXES:
        raise ValueError("Template Mode accepts .docx or .dotx Word template files.")
    if not zipfile.is_zipfile(p):
        raise ValueError("Template is not a valid ZIP-based Microsoft Word package.")
    with zipfile.ZipFile(p) as zf:
        if "word/document.xml" not in set(zf.namelist()):
            raise ValueError("Template package is missing word/document.xml.")
    return p


def _read_xml(path: Path, part: str) -> ET.Element | None:
    with zipfile.ZipFile(path) as zf:
        if part not in set(zf.namelist()):
            return None
        return ET.fromstring(zf.read(part))


def _twips_to_inches(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return round(int(value) / 1440.0, 3)
    except (TypeError, ValueError):
        return None


def _section_snapshot(root: ET.Element) -> dict:
    sect = root.find(".//w:sectPr", NS)
    if sect is None:
        return {}
    pg_sz = sect.find("w:pgSz", NS)
    pg_mar = sect.find("w:pgMar", NS)
    cols = sect.find("w:cols", NS)
    lines = sect.find("w:lnNumType", NS)

    def attr(node: ET.Element | None, name: str):
        return None if node is None else node.attrib.get(_q(name))

    margins = {}
    if pg_mar is not None:
        for side in ("top", "right", "bottom", "left"):
            raw = attr(pg_mar, side)
            margins[side] = {
                "twips": int(raw) if raw and raw.lstrip("-").isdigit() else raw,
                "inches": _twips_to_inches(raw),
            }
    page_size = {}
    if pg_sz is not None:
        for name in ("w", "h", "orient"):
            value = attr(pg_sz, name)
            if value is not None:
                page_size[name] = value
    column_info = {}
    if cols is not None:
        for name in ("num", "space", "equalWidth"):
            value = attr(cols, name)
            if value is not None:
                column_info[name] = value
    line_info = {"enabled": lines is not None}
    if lines is not None:
        for name in ("countBy", "restart", "distance", "start"):
            value = attr(lines, name)
            if value is not None:
                line_info[name] = value
    return {"page_size": page_size, "margins": margins, "columns": column_info, "line_numbering": line_info}


def _style_snapshot(root: ET.Element | None) -> list[dict]:
    if root is None:
        return []
    result: list[dict] = []
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(_q("styleId"))
        if style_id not in SAFE_STYLE_IDS:
            continue
        name = style.find("w:name", NS)
        result.append({
            "style_id": style_id,
            "name": name.attrib.get(_q("val"), style_id) if name is not None else style_id,
            "paragraph_properties": style.find("w:pPr", NS) is not None,
            "run_properties": style.find("w:rPr", NS) is not None,
        })
    return result


def inspect_template(template_path: str | Path) -> dict:
    template = _validate_template(template_path)
    document = _read_xml(template, "word/document.xml")
    styles = _read_xml(template, "word/styles.xml")
    assert document is not None
    section = _section_snapshot(document)
    style_rows = _style_snapshot(styles)
    warnings: list[str] = []
    if not section:
        warnings.append("No section properties were found in the template, so page-level formatting cannot be transferred.")
    if not style_rows:
        warnings.append("No supported standard styles were found in the template.")
    warnings.append("Template Mode transfers safe formatting primitives only. It does not copy template body text, headers, footers, macros, citations, figures, or instructional placeholders.")
    warnings.append("Direct formatting already applied to individual manuscript runs or paragraphs can override transferred style definitions and may require manual review.")
    return {
        "workflow": "Template Mode",
        "template": template.name,
        "template_type": template.suffix.lower(),
        "page_format": section,
        "transferable_styles": style_rows,
        "transferable_style_count": len(style_rows),
        "supported_section_properties": ["page size/orientation", "margins", "columns", "line numbering"],
        "warnings": warnings,
    }


def _replace_child(parent: ET.Element, tag_local: str, source: ET.Element | None) -> bool:
    existing = parent.find(f"w:{tag_local}", NS)
    if source is None:
        if existing is not None:
            parent.remove(existing)
            return True
        return False
    clone = copy.deepcopy(source)
    if existing is None:
        parent.append(clone)
    else:
        index = list(parent).index(existing)
        parent.remove(existing)
        parent.insert(index, clone)
    return True


def _apply_section_format(source_data: bytes, template_root: ET.Element, applied: list[str]) -> bytes:
    root = ET.fromstring(source_data)
    template_sect = template_root.find(".//w:sectPr", NS)
    if template_sect is None:
        return source_data
    source_sections = root.findall(".//w:sectPr", NS)
    if not source_sections:
        return source_data
    for prop, label in (("pgSz", "page size/orientation"), ("pgMar", "page margins"), ("cols", "column layout"), ("lnNumType", "line numbering")):
        template_prop = template_sect.find(f"w:{prop}", NS)
        changed = False
        for source_sect in source_sections:
            changed = _replace_child(source_sect, prop, template_prop) or changed
        if changed:
            applied.append(label)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _style_map(root: ET.Element) -> dict[str, ET.Element]:
    result = {}
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(_q("styleId"))
        if style_id:
            result[style_id] = style
    return result


def _apply_style_format(source_data: bytes, template_styles: ET.Element | None, applied: list[str]) -> bytes:
    if template_styles is None:
        return source_data
    root = ET.fromstring(source_data)
    source_map = _style_map(root)
    template_map = _style_map(template_styles)
    touched: list[str] = []
    for style_id in SAFE_STYLE_IDS:
        source_style = source_map.get(style_id)
        template_style = template_map.get(style_id)
        if source_style is None or template_style is None:
            continue
        changed = False
        for prop in ("pPr", "rPr"):
            template_prop = template_style.find(f"w:{prop}", NS)
            if template_prop is not None:
                changed = _replace_child(source_style, prop, template_prop) or changed
        if changed:
            touched.append(style_id)
    if not touched:
        return source_data
    applied.append("styles: " + ", ".join(touched))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


@dataclass
class TemplateRetargetResult:
    input: str
    template: str
    output: str
    applied: list[str]
    warnings: list[str]
    preservation: dict

    @property
    def passed(self) -> bool:
        return bool(self.preservation.get("passed"))

    def to_dict(self) -> dict:
        data = asdict(self)
        data["passed"] = self.passed
        return data


def retarget_from_template(input_path: str | Path, output_path: str | Path, template_path: str | Path) -> TemplateRetargetResult:
    src = Path(input_path)
    dst = Path(output_path)
    template = _validate_template(template_path)
    DocxPackage(src)
    template_document = _read_xml(template, "word/document.xml")
    template_styles = _read_xml(template, "word/styles.xml")
    assert template_document is not None
    applied: list[str] = []
    warnings = [
        "The journal template body, headers, footers, relationships, macros, and placeholder text were not copied.",
        "Review the output in Microsoft Word because direct formatting can override style-level template formatting.",
    ]
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = _apply_section_format(data, template_document, applied)
            elif item.filename == "word/styles.xml":
                data = _apply_style_format(data, template_styles, applied)
            zout.writestr(item, data)
    preservation = verify_preservation(src, dst).to_dict()
    if not preservation["passed"]:
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because the standard manuscript preservation gate failed.")
    elif not applied:
        warnings.append("No supported transferable template properties matched the manuscript; the output is effectively an unchanged copy.")
    return TemplateRetargetResult(
        input=str(src), template=str(template), output=str(dst), applied=applied, warnings=warnings, preservation=preservation
    )
