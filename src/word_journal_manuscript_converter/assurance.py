from __future__ import annotations

from pathlib import Path
from typing import Any

from .audit import inspect_docx, validate_docx_structure, verify_preservation
from .docx_package import DocxPackage, NS, W_NS
from .journal import JournalProfile, readiness_check


_FORMAT_REQUIREMENTS = {
    "margins_inches",
    "line_numbering",
    "body_font",
    "body_font_size_pt",
    "line_spacing",
}
_READINESS_REQUIREMENTS = {
    "requires_live_citations",
    "requires_figures",
    "max_comments",
    "tracked_changes_allowed",
    "abstract_required",
    "abstract_max_words",
    "abstract_recommended_max_words",
    "keywords_min",
    "keywords_max",
    "required_sections",
    "citations_must_resolve",
}


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _twips_to_inches(value: str | None) -> float | None:
    if value is None:
        return None
    try:
        return int(value) / 1440.0
    except (TypeError, ValueError):
        return None


def _section_formats(path: str | Path) -> list[dict[str, Any]]:
    package = DocxPackage(path)
    root = package.xml("word/document.xml")
    rows: list[dict[str, Any]] = []
    for sect in root.findall(".//w:sectPr", NS):
        pgmar = sect.find("w:pgMar", NS)
        lines = sect.find("w:lnNumType", NS)
        margins: dict[str, float | None] = {}
        if pgmar is not None:
            for side in ("top", "right", "bottom", "left"):
                margins[side] = _twips_to_inches(pgmar.attrib.get(_q(side)))
        line_numbering: dict[str, Any] = {"enabled": lines is not None}
        if lines is not None:
            for attr in ("countBy", "restart", "distance", "start"):
                value = lines.attrib.get(_q(attr))
                if value is not None:
                    line_numbering[attr] = value
        rows.append({"margins_inches": margins, "line_numbering": line_numbering})
    return rows


def _normal_style_format(path: str | Path) -> dict[str, Any]:
    package = DocxPackage(path)
    if not package.has("word/styles.xml"):
        return {}
    root = package.xml("word/styles.xml")
    normal = None
    for style in root.findall("w:style", NS):
        if style.attrib.get(_q("styleId")) == "Normal":
            normal = style
            break
    if normal is None:
        return {}
    result: dict[str, Any] = {}
    fonts = normal.find("w:rPr/w:rFonts", NS)
    if fonts is not None:
        result["font"] = fonts.attrib.get(_q("ascii")) or fonts.attrib.get(_q("hAnsi"))
    size = normal.find("w:rPr/w:sz", NS)
    if size is not None:
        try:
            result["font_size_pt"] = int(size.attrib.get(_q("val"), "0")) / 2.0
        except ValueError:
            pass
    spacing = normal.find("w:pPr/w:spacing", NS)
    if spacing is not None and spacing.attrib.get(_q("line")):
        try:
            line = int(spacing.attrib[_q("line")])
            if spacing.attrib.get(_q("lineRule"), "auto") == "auto":
                result["line_spacing"] = line / 240.0
            else:
                result["line_spacing_twips"] = line
        except ValueError:
            pass
    return result


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


def _check(
    name: str,
    category: str,
    status: str,
    detail: str,
    *,
    expected: Any = None,
    actual: Any = None,
    machine_verifiable: bool = True,
) -> dict[str, Any]:
    return {
        "check": name,
        "category": category,
        "status": status,
        "detail": detail,
        "expected": expected,
        "actual": actual,
        "machine_verifiable": machine_verifiable,
    }


def _verify_profile_formatting(output_path: str | Path, req: dict[str, Any]) -> list[dict[str, Any]]:
    checks: list[dict[str, Any]] = []
    sections = _section_formats(output_path)
    normal = _normal_style_format(output_path)

    margins = req.get("margins_inches")
    if isinstance(margins, dict):
        for side in ("top", "right", "bottom", "left"):
            if side not in margins:
                continue
            expected = float(margins[side])
            actuals = [row.get("margins_inches", {}).get(side) for row in sections]
            ok = bool(actuals) and all(x is not None and abs(float(x) - expected) <= (1.0 / 1440.0) for x in actuals)
            checks.append(_check(
                f"format:margins:{side}", "formatting", "pass" if ok else "fail",
                f"{side.title()} margin {'matches' if ok else 'does not match'} the profile target in every section.",
                expected=expected, actual=actuals,
            ))

    line_numbers = req.get("line_numbering")
    if line_numbers:
        enabled = [bool(row.get("line_numbering", {}).get("enabled")) for row in sections]
        ok = bool(enabled) and all(enabled)
        detail = "Line numbering is enabled in every section." if ok else "Line numbering is not enabled in every section."
        checks.append(_check("format:line_numbering", "formatting", "pass" if ok else "fail", detail, expected=True, actual=enabled))
        if isinstance(line_numbers, dict):
            mapping = {"count_by": "countBy", "restart": "restart", "distance_twips": "distance", "start": "start"}
            for source_key, xml_key in mapping.items():
                if source_key not in line_numbers:
                    continue
                expected = str(line_numbers[source_key])
                actuals = [row.get("line_numbering", {}).get(xml_key) for row in sections]
                ok = bool(actuals) and all(str(x) == expected for x in actuals)
                checks.append(_check(
                    f"format:line_numbering:{source_key}", "formatting", "pass" if ok else "fail",
                    f"Line-numbering {source_key.replace('_', ' ')} {'matches' if ok else 'does not match'} the profile target.",
                    expected=expected, actual=actuals,
                ))

    font = req.get("body_font")
    if font:
        actual = normal.get("font")
        ok = isinstance(actual, str) and actual.casefold() == str(font).casefold()
        checks.append(_check(
            "format:body_font", "formatting", "pass" if ok else "fail",
            f"Normal style font is {actual or 'not detected'}; target is {font}.", expected=font, actual=actual,
        ))

    size = req.get("body_font_size_pt")
    if size is not None:
        expected = float(size)
        actual = normal.get("font_size_pt")
        ok = actual is not None and abs(float(actual) - expected) < 0.01
        checks.append(_check(
            "format:body_font_size", "formatting", "pass" if ok else "fail",
            f"Normal style font size is {actual if actual is not None else 'not detected'} pt; target is {expected:g} pt.",
            expected=expected, actual=actual,
        ))

    spacing = req.get("line_spacing")
    if spacing is not None:
        expected = float(spacing)
        actual = normal.get("line_spacing")
        ok = actual is not None and abs(float(actual) - expected) < 0.01
        checks.append(_check(
            "format:line_spacing", "formatting", "pass" if ok else "fail",
            f"Normal style line spacing is {actual if actual is not None else 'not detected'}; target is {expected:g}.",
            expected=expected, actual=actual,
        ))

    direct = _direct_formatting_counts(output_path)
    if any(k in req for k in ("body_font", "body_font_size_pt", "line_spacing")) and (direct["paragraphs"] or direct["runs"]):
        checks.append(_check(
            "format:direct_formatting_review", "formatting", "manual",
            f"Detected direct formatting in {direct['paragraphs']} paragraphs and {direct['runs']} runs. Direct formatting can override style-level appearance, so visually review the output in Word.",
            expected="No unintended overrides", actual=direct, machine_verifiable=False,
        ))
    return checks


def assess_profile_conversion(
    input_path: str | Path,
    output_path: str | Path,
    profile_path: str | Path,
) -> dict[str, Any]:
    """Independently verify a journal-profile conversion after the output is written.

    The assurance pass re-opens the saved DOCX instead of trusting transformation
    messages. It separates machine-verifiable formatting, manuscript requirements,
    document integrity, and manual/unsupported requirements.
    """
    src = Path(input_path)
    dst = Path(output_path)
    profile = JournalProfile.from_json(profile_path)
    req = profile.requirements

    checks: list[dict[str, Any]] = []
    formatting_checks = _verify_profile_formatting(dst, req)
    checks.extend(formatting_checks)

    readiness = readiness_check(dst, profile_path)
    for row in readiness.get("checks", []):
        if row.get("check", "").startswith("format:"):
            continue
        status = str(row.get("status", "info"))
        if status == "info":
            continue
        checks.append(_check(
            str(row.get("check", "requirement")),
            "manuscript_requirement" if not str(row.get("check", "")).startswith("profile_") else "profile_provenance",
            status,
            str(row.get("detail", "")),
            machine_verifiable=status in {"pass", "fail", "warn"},
        ))

    preservation = verify_preservation(src, dst).to_dict()
    structural = validate_docx_structure(dst)
    before_inventory = inspect_docx(src).to_dict()
    after_inventory = inspect_docx(dst).to_dict()

    checks.append(_check(
        "integrity:preservation_gate", "document_integrity", "pass" if preservation.get("passed") else "fail",
        "Protected manuscript content is unchanged after conversion." if preservation.get("passed") else "One or more protected manuscript-content checks failed.",
        expected="PASS", actual="PASS" if preservation.get("passed") else "FAIL",
    ))
    checks.append(_check(
        "integrity:ooxml_structure", "document_integrity", "pass" if structural.get("passed") else "fail",
        "Defensive OOXML structural checks passed." if structural.get("passed") else "Defensive OOXML structural checks found a problem.",
        expected="PASS", actual="PASS" if structural.get("passed") else "FAIL",
    ))
    eq_before = {
        "total": before_inventory.get("equations", 0),
        "native_omml": before_inventory.get("native_equations", 0),
        "embedded_equation_objects": before_inventory.get("embedded_equation_objects", 0),
    }
    eq_after = {
        "total": after_inventory.get("equations", 0),
        "native_omml": after_inventory.get("native_equations", 0),
        "embedded_equation_objects": after_inventory.get("embedded_equation_objects", 0),
    }
    eq_ok = eq_before == eq_after
    checks.append(_check(
        "integrity:equations", "document_integrity", "pass" if eq_ok else "fail",
        f"Equation inventory {'is preserved' if eq_ok else 'changed'}: before={eq_before}, after={eq_after}.",
        expected=eq_before, actual=eq_after,
    ))

    handled = _FORMAT_REQUIREMENTS | _READINESS_REQUIREMENTS
    unsupported = sorted(str(k) for k in req if k not in handled)
    for key in unsupported:
        checks.append(_check(
            f"unsupported:{key}", "unsupported", "unsupported",
            f"Requirement '{key}' is present in the journal profile but is not yet independently machine-verified by the assurance engine.",
            expected=req.get(key), actual=None, machine_verifiable=False,
        ))

    format_machine = [c for c in checks if c["category"] == "formatting" and c["status"] in {"pass", "fail"}]
    formatting_score = round(100 * sum(c["status"] == "pass" for c in format_machine) / len(format_machine)) if format_machine else 100

    blocking = [c for c in checks if c["status"] == "fail"]
    manual = [c for c in checks if c["status"] in {"warn", "manual"}]
    unsupported_checks = [c for c in checks if c["status"] == "unsupported"]
    verdict = "BLOCKED - FIX REQUIRED" if blocking else "READY FOR FINAL AUTHOR REVIEW"

    return {
        "workflow": "Conversion Assurance",
        "journal": profile.journal,
        "article_type": profile.article_type,
        "profile_source_url": profile.source_url,
        "profile_source_urls": profile.source_urls or ([profile.source_url] if profile.source_url else []),
        "profile_checked_on": profile.checked_on,
        "formatting_compliance_score": formatting_score,
        "manuscript_requirement_score": readiness.get("readiness_score", 100),
        "document_integrity": "PASS" if preservation.get("passed") else "FAIL",
        "structural_sanity": "PASS" if structural.get("passed") else "FAIL",
        "equation_inventory_before": eq_before,
        "equation_inventory_after": eq_after,
        "blocking_failures": len(blocking),
        "manual_review_items": len(manual),
        "unsupported_requirements": len(unsupported_checks),
        "verdict": verdict,
        "checks": checks,
        "preservation": preservation,
        "structural_validation": structural,
        "note": (
            "Assurance covers only rules that can be safely inspected from the DOCX and the selected source-dated profile. "
            "It does not guarantee editorial acceptance or replace a final review of the journal's current official instructions."
        ),
    }
