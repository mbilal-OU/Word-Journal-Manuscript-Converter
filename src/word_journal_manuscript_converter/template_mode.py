from __future__ import annotations

import copy
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import validate_docx_structure, verify_preservation
from .docx_package import DocxPackage, NS, W_NS
from .ooxml_order import ensure_child, replace_child_ordered

ET.register_namespace("w", W_NS)

CORE_STYLE_IDS = (
    "Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Heading4",
    "Caption", "Quote", "IntenseQuote", "ListParagraph", "Bibliography",
)
TEMPLATE_SUFFIXES = {".docx", ".dotx"}
_SHARED_FORMAT_PARTS = (
    "word/theme/theme1.xml",
    "word/fontTable.xml",
    "word/stylesWithEffects.xml",
)


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


def _style_name(style: ET.Element) -> str:
    name = style.find("w:name", NS)
    return (name.attrib.get(_q("val"), "") if name is not None else "").strip()


def _style_snapshot(root: ET.Element | None) -> list[dict]:
    if root is None:
        return []
    result: list[dict] = []
    for style in root.findall("w:style", NS):
        if style.attrib.get(_q("type")) not in {None, "paragraph"}:
            continue
        style_id = style.attrib.get(_q("styleId"))
        if not style_id:
            continue
        result.append({
            "style_id": style_id,
            "name": _style_name(style) or style_id,
            "class": "core" if style_id in CORE_STYLE_IDS else "custom",
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
            if style.attrib.get(_q("type")) in {None, "paragraph"} and style_id and style_id not in CORE_STYLE_IDS:
                custom_styles += 1
    return {
        "headers": sum(name.startswith("word/header") and name.endswith(".xml") for name in names),
        "footers": sum(name.startswith("word/footer") and name.endswith(".xml") for name in names),
        "numbering_part": "word/numbering.xml" in names,
        "theme_part": "word/theme/theme1.xml" in names,
        "font_table_part": "word/fontTable.xml" in names,
        "styles_with_effects_part": "word/stylesWithEffects.xml" in names,
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
        warnings.append("No paragraph styles were found in the template.")
    warnings.append(
        "Template Mode transfers and verifies a conservative formatting surface. It does not claim an exact visual clone of every publisher template."
    )
    warnings.append(
        "Matching paragraph styles, document defaults, page layout, and compatible theme/font parts can be transferred. Headers, footers, macros, placeholders, and numbering definitions remain protected from automatic import."
    )
    warnings.append(
        "Direct formatting already applied to manuscript runs or paragraphs can override template styles and may require manual review."
    )
    return {
        "workflow": "Template Mode",
        "template": template.name,
        "template_type": template.suffix.lower(),
        "page_format": section,
        "transferable_styles": style_rows,
        "transferable_style_count": len(style_rows),
        "template_features": features,
        "supported_section_properties": [
            "page size/orientation", "margins", "columns", "line numbering",
            "matching paragraph styles", "document defaults", "compatible theme/font parts",
        ],
        "exact_visual_match_guaranteed": False,
        "warnings": warnings,
    }


def _replace_child(parent: ET.Element, tag_local: str, source: ET.Element | None) -> bool:
    return replace_child_ordered(parent, tag_local, source)


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


def _style_map(root: ET.Element | None) -> dict[str, ET.Element]:
    result: dict[str, ET.Element] = {}
    if root is None:
        return result
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(_q("styleId"))
        if style_id:
            result[style_id] = style
    return result


def _style_name_map(root: ET.Element | None) -> dict[str, ET.Element]:
    result: dict[str, ET.Element] = {}
    if root is None:
        return result
    for style in root.findall("w:style", NS):
        name = _style_name(style).casefold()
        if name:
            result[name] = style
    return result


def _copy_doc_defaults(source_root: ET.Element, template_root: ET.Element, touched: list[str]) -> None:
    template_defaults = template_root.find("w:docDefaults", NS)
    if template_defaults is None:
        return
    source_defaults = source_root.find("w:docDefaults", NS)
    clone = copy.deepcopy(template_defaults)
    if source_defaults is None:
        source_root.insert(0, clone)
    else:
        index = list(source_root).index(source_defaults)
        source_root.remove(source_defaults)
        source_root.insert(index, clone)
    touched.append("document defaults")


def _apply_style_format(source_data: bytes, template_styles: ET.Element | None, applied: list[str]) -> bytes:
    if template_styles is None:
        return source_data
    root = ET.fromstring(source_data)
    source_map = _style_map(root)
    template_map = _style_map(template_styles)
    template_names = _style_name_map(template_styles)
    touched: list[str] = []

    _copy_doc_defaults(root, template_styles, touched)

    for source_id, source_style in source_map.items():
        if source_style.attrib.get(_q("type")) not in {None, "paragraph"}:
            continue
        template_style = template_map.get(source_id)
        if template_style is None:
            source_name = _style_name(source_style).casefold()
            template_style = template_names.get(source_name) if source_name else None
        if template_style is None:
            continue
        changed = False
        for prop in ("pPr", "rPr"):
            template_prop = template_style.find(f"w:{prop}", NS)
            if template_prop is not None:
                changed = _replace_child(source_style, prop, template_prop) or changed
        if changed:
            touched.append(source_id)

    if not touched:
        return source_data
    applied.append("styles: " + ", ".join(touched))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def _semantic_style_ids(template_styles: ET.Element | None) -> dict[str, str]:
    if template_styles is None:
        return {}
    by_id = _style_map(template_styles)
    by_name = _style_name_map(template_styles)
    result: dict[str, str] = {}
    canonical = {
        "normal": "Normal",
        "title": "Title",
        "subtitle": "Subtitle",
        "heading1": "Heading1",
        "heading2": "Heading2",
        "heading3": "Heading3",
        "heading4": "Heading4",
        "caption": "Caption",
        "bibliography": "Bibliography",
    }
    for role, style_id in canonical.items():
        if style_id in by_id:
            result[role] = style_id
    name_aliases = {
        "normal": ("normal", "body text"),
        "title": ("title", "article title"),
        "subtitle": ("subtitle",),
        "heading1": ("heading 1", "heading1"),
        "heading2": ("heading 2", "heading2"),
        "heading3": ("heading 3", "heading3"),
        "heading4": ("heading 4", "heading4"),
        "caption": ("caption", "figure caption"),
        "bibliography": ("bibliography", "references", "reference"),
    }
    for role, aliases in name_aliases.items():
        if role in result:
            continue
        for alias in aliases:
            style = by_name.get(alias.casefold())
            if style is not None and style.attrib.get(_q("styleId")):
                result[role] = style.attrib[_q("styleId")]
                break
    return result


def _apply_semantic_paragraph_styles(source_data: bytes, template_styles: ET.Element | None, applied: list[str]) -> bytes:
    mapping = _semantic_style_ids(template_styles)
    if not mapping:
        return source_data
    root = ET.fromstring(source_data)
    paragraphs = root.findall(".//w:body/w:p", NS)
    ref_heading_index: int | None = None
    for i, p in enumerate(paragraphs):
        if _paragraph_text(p).casefold().rstrip(":") in {"references", "bibliography"}:
            ref_heading_index = i
            break

    changed = 0
    common_headings = {
        "abstract", "introduction", "background", "methods", "materials and methods",
        "results", "discussion", "conclusion", "conclusions", "references", "bibliography",
        "acknowledgments", "acknowledgements", "funding", "data availability",
        "author contributions", "competing interests", "conflict of interest",
    }

    for i, p in enumerate(paragraphs):
        text = _paragraph_text(p)
        ppr = p.find("w:pPr", NS)
        pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
        current = pstyle.attrib.get(_q("val")) if pstyle is not None else None
        role: str | None = None
        current_folded = (current or "").replace(" ", "").casefold()
        for candidate in ("title", "subtitle", "heading1", "heading2", "heading3", "heading4", "caption", "bibliography"):
            if current_folded == candidate:
                role = candidate
                break
        if role is None and (current is None or current_folded == "normal"):
            role = "normal"
        normalized = text.casefold().rstrip(":").strip()
        if normalized in common_headings and "heading1" in mapping:
            role = "heading1"
        if ref_heading_index is not None and i > ref_heading_index and "bibliography" in mapping and text:
            role = "bibliography"
        if text.casefold().startswith(("figure ", "table ")) and "caption" in mapping:
            role = "caption"
        target = mapping.get(role or "")
        if not target or target == current:
            continue
        ppr = ensure_child(p, "pPr")
        pstyle = ensure_child(ppr, "pStyle")
        pstyle.set(_q("val"), target)
        changed += 1

    if not changed:
        return source_data
    applied.append(f"semantic paragraph style assignments: {changed}")
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


def _part_equal(output: Path, template: Path, part: str) -> tuple[bool | None, str]:
    with zipfile.ZipFile(output) as oz, zipfile.ZipFile(template) as tz:
        onames = set(oz.namelist())
        tnames = set(tz.namelist())
        if part not in tnames:
            return None, "Template does not contain this part."
        if part not in onames:
            return False, "Template contains this part but the manuscript package does not have a compatible target part."
        return oz.read(part) == tz.read(part), "Compared output and template package parts directly."


def verify_template_fidelity(output_path: str | Path, template_path: str | Path) -> dict:
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

    out_map = _style_map(out_styles)
    out_names = _style_name_map(out_styles)
    tpl_map = _style_map(tpl_styles)
    checked_styles = 0
    unmatched_custom = 0
    for style_id, tpl_style in tpl_map.items():
        if tpl_style.attrib.get(_q("type")) not in {None, "paragraph"}:
            continue
        out_style = out_map.get(style_id)
        if out_style is None:
            name = _style_name(tpl_style).casefold()
            out_style = out_names.get(name) if name else None
        has_format = any(tpl_style.find(f"w:{prop}", NS) is not None for prop in ("pPr", "rPr"))
        if out_style is None:
            if style_id not in CORE_STYLE_IDS and has_format:
                unmatched_custom += 1
            continue
        for prop in ("pPr", "rPr"):
            expected = tpl_style.find(f"w:{prop}", NS)
            if expected is None:
                continue
            actual = out_style.find(f"w:{prop}", NS)
            ok = _node_signature(actual) == _node_signature(expected)
            checks.append({
                "check": f"style:{style_id}:{prop}",
                "status": "pass" if ok else "fail",
                "detail": f"{style_id} {prop} {'matches' if ok else 'does not match'} the template.",
                "machine_verifiable": True,
            })
            checked_styles += 1

    if unmatched_custom:
        checks.append({
            "check": "coverage:unmatched_custom_styles",
            "status": "unsupported",
            "detail": f"The template contains {unmatched_custom} formatted custom paragraph style(s) that do not map to a manuscript style by ID or name.",
            "machine_verifiable": False,
        })

    for part, label in (("word/theme/theme1.xml", "theme"), ("word/fontTable.xml", "font table"), ("word/stylesWithEffects.xml", "style effects")):
        equal, detail = _part_equal(output, template, part)
        if equal is None:
            continue
        if equal:
            checks.append({"check": f"part:{label}", "status": "pass", "detail": f"Output {label} matches the template.", "machine_verifiable": True})
        else:
            checks.append({"check": f"coverage:{label}", "status": "unsupported", "detail": detail, "machine_verifiable": False})

    features = _template_feature_inventory(template, tpl_styles)
    for key, label in (("headers", "template headers"), ("footers", "template footers"), ("numbering_part", "template numbering definitions")):
        value = features[key]
        if bool(value):
            checks.append({
                "check": f"coverage:{key}",
                "status": "unsupported",
                "detail": f"Detected {label} ({value}). This feature is not imported automatically because it can change manuscript content or relationship structure.",
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
        "style_checks": checked_styles,
        "exact_visual_match_guaranteed": False,
        "verdict": "SUPPORTED TEMPLATE FORMAT VERIFIED" if not blocking else "SUPPORTED TEMPLATE FORMAT MISMATCH",
        "checks": checks,
        "note": (
            "A 100% supported-fidelity score means the formatting primitives this engine transferred match the template. "
            "It does not mean the entire publisher template was reproduced. Template coverage and manual-review items show the remaining gap."
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
        "Template Mode does not claim an exact visual clone of every supplied template.",
        "The engine adapts compatible page layout, matching paragraph styles, document defaults, and selected package-level formatting while protecting manuscript content.",
        "Headers, footers, macros, placeholder text, and numbering definitions are not imported automatically because they can alter content or relationship structure.",
        "The fidelity report separates verified formatting from unsupported or manual-review template features.",
    ]
    dst.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(template, "r") as tzin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        template_names = set(tzin.namelist())
        source_names = set(zin.namelist())
        shared_replacements = {
            part: tzin.read(part)
            for part in _SHARED_FORMAT_PARTS
            if part in template_names and part in source_names
        }
        for part in shared_replacements:
            applied.append("template package part: " + part)

        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = _apply_section_format(data, template_document, applied)
                data = _apply_semantic_paragraph_styles(data, template_styles, applied)
            elif item.filename == "word/styles.xml":
                data = _apply_style_format(data, template_styles, applied)
            elif item.filename in shared_replacements:
                data = shared_replacements[item.filename]
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
        warnings.append("Output removed because formatting that Template Mode claimed to transfer did not verify against the supplied template.")
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
