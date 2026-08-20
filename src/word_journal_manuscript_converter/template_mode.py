from __future__ import annotations

import copy
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import validate_docx_structure, verify_preservation
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


def _template_feature_inventory(template: Path, styles: ET.Element | None) -> dict:
    with zipfile.ZipFile(template) as zf:
        names = set(zf.namelist())
    custom_styles = 0
    if styles is not None:
        for style in styles.findall("w:style", NS):
            style_id = style.attrib.get(_q("styleId"))
            if style_id and style_id not in SAFE_STYLE_IDS:
                custom_styles += 1
    return {
        "headers": sum(name.startswith("word/header") and name.endswith(".xml") for name in names),
        "footers": sum(name.startswith("word/footer") and name.endswith(".xml") for name in names),
        "numbering_part": "word/numbering.xml" in names,
        "theme_part": any(name.startswith("word/theme/") and name.endswith(".xml") for name in names),
        "custom_style_count": custom_styles,
    }


def inspect_template(template_path: str | Path) -> dict:
    template = _validate_template(template_path)
    document = _read_xml(template, "word/document.xml")
    styles = _read_xml(template, "word/styles.xml")
    assert document is not None
    section = _section_snapshot(document)
    style_rows = _style_snapshot(styles)
    features = _template_feature_inventory(template, styles)
    warnings: list[str] = []
    if not section:
        warnings.append("No section properties were found in the template, so page-level formatting cannot be transferred.")
    if not style_rows:
        warnings.append("No supported standard styles were found in the template.")
    warnings.append(
        "Template Mode transfers and verifies a conservative subset of formatting. It does not claim an exact visual clone of the publisher template."
    )
    warnings.append(
        "Template body text, placeholders, headers, footers, macros, citations, figures, unrelated relationships, custom style systems, themes, and numbering definitions are not imported automatically."
    )
    warnings.append(
        "Direct formatting already applied to individual manuscript runs or paragraphs can override transferred style definitions and may require manual review."
    )
    return {
        "workflow": "Template Mode",
        "template": template.name,
        "template_type": template.suffix.lower(),
        "page_format": section,
        "transferable_styles": style_rows,
        "transferable_style_count": len(style_rows),
        "template_features": features,
        "supported_section_properties": ["page size/orientation", "margins", "columns", "line numbering"],
        "exact_visual_match_guaranteed": False,
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


def _node_signature(node: ET.Element | None):
    if node is None:
        return None
    return (
        node.tag,
        tuple(sorted(node.attrib.items())),
        (node.text or "").strip(),
        tuple(_node_signature(child) for child in list(node)),
    )


def _direct_formatting_counts(path: str | Path) -> dict[str, int]:
    package = DocxPackage(path)
    paragraphs = 0
    runs = 0
    for part in package.story_parts():
        root = package.xml(part)
        for ppr in root.findall(".//w:p/w:pPr", NS):
            if any(child.tag != _q("pStyle") for child in list(ppr)):
                paragraphs += 1
        for rpr in root.findall(".//w:r/w:rPr", NS):
            if list(rpr):
                runs += 1
    return {"paragraphs": paragraphs, "runs": runs}


def verify_template_fidelity(output_path: str | Path, template_path: str | Path) -> dict:
    """Verify the safe subset of template formatting after conversion.

    Two scores are intentionally reported: supported fidelity says whether the
    formatting primitives the engine attempted to transfer actually match the
    template; template coverage says how much of the discovered template
    formatting surface is within the current safe-transfer scope.
    """
    output = Path(output_path)
    template = _validate_template(template_path)
    out_doc = _read_xml(output, "word/document.xml")
    tpl_doc = _read_xml(template, "word/document.xml")
    out_styles = _read_xml(output, "word/styles.xml")
    tpl_styles = _read_xml(template, "word/styles.xml")
    assert out_doc is not None and tpl_doc is not None

    checks: list[dict] = []
    tpl_sect = tpl_doc.find(".//w:sectPr", NS)
    out_sections = out_doc.findall(".//w:sectPr", NS)
    for prop, label in (("pgSz", "page size/orientation"), ("pgMar", "page margins"), ("cols", "column layout"), ("lnNumType", "line numbering")):
        expected = None if tpl_sect is None else tpl_sect.find(f"w:{prop}", NS)
        if expected is None:
            continue
        actuals = [sect.find(f"w:{prop}", NS) for sect in out_sections]
        ok = bool(actuals) and all(_node_signature(actual) == _node_signature(expected) for actual in actuals)
        checks.append({
            "check": f"section:{prop}",
            "status": "pass" if ok else "fail",
            "detail": f"Output {label} {'matches' if ok else 'does not match'} the template in every manuscript section.",
            "machine_verifiable": True,
        })

    out_map = _style_map(out_styles) if out_styles is not None else {}
    tpl_map = _style_map(tpl_styles) if tpl_styles is not None else {}
    for style_id in SAFE_STYLE_IDS:
        tpl_style = tpl_map.get(style_id)
        if tpl_style is None:
            continue
        out_style = out_map.get(style_id)
        for prop in ("pPr", "rPr"):
            expected = tpl_style.find(f"w:{prop}", NS)
            if expected is None:
                continue
            if out_style is None:
                checks.append({
                    "check": f"style:{style_id}:{prop}",
                    "status": "unsupported",
                    "detail": f"Template style {style_id} has {prop} formatting, but the manuscript does not contain that standard style ID for safe replacement.",
                    "machine_verifiable": False,
                })
                continue
            actual = out_style.find(f"w:{prop}", NS)
            ok = _node_signature(actual) == _node_signature(expected)
            checks.append({
                "check": f"style:{style_id}:{prop}",
                "status": "pass" if ok else "fail",
                "detail": f"{style_id} {prop} {'matches' if ok else 'does not match'} the template.",
                "machine_verifiable": True,
            })

    features = _template_feature_inventory(template, tpl_styles)
    feature_labels = {
        "headers": "template headers",
        "footers": "template footers",
        "numbering_part": "template numbering definitions",
        "theme_part": "template theme",
        "custom_style_count": "template custom styles",
    }
    for key, label in feature_labels.items():
        value = features[key]
        present = bool(value)
        if present:
            checks.append({
                "check": f"coverage:{key}",
                "status": "unsupported",
                "detail": f"Detected {label} ({value}). This feature is intentionally not imported by safe Template Mode.",
                "machine_verifiable": False,
            })

    direct = _direct_formatting_counts(output)
    if direct["paragraphs"] or direct["runs"]:
        checks.append({
            "check": "manual:direct_formatting",
            "status": "manual",
            "detail": f"Detected direct formatting in {direct['paragraphs']} paragraphs and {direct['runs']} runs. It can override transferred styles and requires visual review in Word.",
            "machine_verifiable": False,
        })

    machine = [row for row in checks if row["status"] in {"pass", "fail"}]
    supported_score = round(100 * sum(row["status"] == "pass" for row in machine) / len(machine)) if machine else 100
    coverage_denominator = len(machine) + sum(row["status"] == "unsupported" for row in checks)
    coverage_score = round(100 * sum(row["status"] == "pass" for row in machine) / coverage_denominator) if coverage_denominator else 100
    blocking = sum(row["status"] == "fail" for row in checks)
    manual = sum(row["status"] in {"manual", "unsupported"} for row in checks)

    return {
        "supported_fidelity_score": supported_score,
        "template_coverage_score": coverage_score,
        "blocking_failures": blocking,
        "manual_review_items": manual,
        "exact_visual_match_guaranteed": False,
        "verdict": "SUPPORTED TEMPLATE FORMAT VERIFIED" if not blocking else "SUPPORTED TEMPLATE FORMAT MISMATCH",
        "checks": checks,
        "note": (
            "A 100% supported-fidelity score means the safe formatting primitives transferred by this engine match the template. "
            "It does not mean the entire template was reproduced. Template coverage and manual-review items show the remaining gap."
        ),
    }


@dataclass
class TemplateRetargetResult:
    input: str
    template: str
    output: str
    applied: list[str]
    warnings: list[str]
    preservation: dict
    structural_validation: dict
    fidelity: dict

    @property
    def passed(self) -> bool:
        return (
            bool(self.preservation.get("passed"))
            and bool(self.structural_validation.get("passed"))
            and int(self.fidelity.get("blocking_failures", 0)) == 0
        )

    def to_dict(self) -> dict:
        data = asdict(self)
        data["passed"] = self.passed
        data["supported_fidelity_score"] = self.fidelity.get("supported_fidelity_score")
        data["template_coverage_score"] = self.fidelity.get("template_coverage_score")
        data["verdict"] = self.fidelity.get("verdict")
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
        "Template Mode intentionally does not claim an exact visual clone of the supplied template.",
        "The journal template body, headers, footers, relationships, macros, placeholder text, custom style systems, themes, and numbering definitions are not copied automatically.",
        "The fidelity report distinguishes verified safe-transfer formatting from unsupported or manual-review template features.",
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
    structural = validate_docx_structure(dst) if preservation.get("passed") else {
        "passed": False,
        "checks": [],
        "failures": ["Structural validation was not run because preservation failed."],
        "warnings": [],
    }
    fidelity = verify_template_fidelity(dst, template) if preservation.get("passed") and structural.get("passed") else {
        "supported_fidelity_score": 0,
        "template_coverage_score": 0,
        "blocking_failures": 1,
        "manual_review_items": 0,
        "exact_visual_match_guaranteed": False,
        "verdict": "OUTPUT WITHHELD",
        "checks": [],
        "note": "Fidelity verification was not run because an earlier safety gate failed.",
    }

    if not preservation.get("passed"):
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because the standard manuscript preservation gate failed.")
    elif not structural.get("passed"):
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because defensive OOXML structural validation failed.")
    elif fidelity.get("blocking_failures"):
        dst.unlink(missing_ok=True)
        warnings.append("Output removed because the formatting that Template Mode claimed to transfer did not verify against the supplied template.")
    elif not applied:
        warnings.append("No supported transferable template properties matched the manuscript; the output is effectively an unchanged copy.")

    return TemplateRetargetResult(
        input=str(src),
        template=str(template),
        output=str(dst),
        applied=applied,
        warnings=warnings,
        preservation=preservation,
        structural_validation=structural,
        fidelity=fidelity,
    )
