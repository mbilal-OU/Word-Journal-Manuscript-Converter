from __future__ import annotations

import json
import shutil
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import verify_preservation
from .docx_package import DocxPackage, NS, W_NS
from .journal import JournalProfile

ET.register_namespace("w", W_NS)


@dataclass
class Transformation:
    name: str
    status: str
    detail: str


@dataclass
class RetargetResult:
    input: str
    output: str
    profile: str
    transformations: list[Transformation]
    preservation: dict

    @property
    def passed(self) -> bool:
        return bool(self.preservation.get("passed"))

    def to_dict(self) -> dict:
        return {
            "input": self.input,
            "output": self.output,
            "profile": self.profile,
            "transformations": [asdict(t) for t in self.transformations],
            "preservation": self.preservation,
            "passed": self.passed,
        }


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _ensure(parent, tag: str):
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = ET.SubElement(parent, _q(tag))
    return node


def _set_attr(node, name: str, value: str) -> None:
    node.set(_q(name), value)


def _edit_document_xml(data: bytes, req: dict, transformations: list[Transformation]) -> bytes:
    root = ET.fromstring(data)
    changed = False

    margins = req.get("margins_inches")
    line_numbers = req.get("line_numbering")
    for sect in root.findall(".//w:sectPr", NS):
        if isinstance(margins, dict):
            pgmar = _ensure(sect, "pgMar")
            for side in ("top", "right", "bottom", "left"):
                if side in margins:
                    _set_attr(pgmar, side, str(round(float(margins[side]) * 1440)))
                    changed = True
        if line_numbers:
            ln = _ensure(sect, "lnNumType")
            if isinstance(line_numbers, dict):
                if "count_by" in line_numbers:
                    _set_attr(ln, "countBy", str(int(line_numbers["count_by"])))
                if "restart" in line_numbers:
                    _set_attr(ln, "restart", str(line_numbers["restart"]))
                if "distance_twips" in line_numbers:
                    _set_attr(ln, "distance", str(int(line_numbers["distance_twips"])))
            else:
                _set_attr(ln, "countBy", "1")
            changed = True

    if isinstance(margins, dict):
        transformations.append(Transformation("page_margins", "applied", f"Set margins to {margins}."))
    if line_numbers:
        transformations.append(Transformation("line_numbering", "applied", "Enabled Word line-numbering properties."))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else data


def _edit_styles_xml(data: bytes, req: dict, transformations: list[Transformation]) -> bytes:
    root = ET.fromstring(data)
    changed = False
    normal = None
    for style in root.findall("w:style", NS):
        if style.attrib.get(_q("styleId")) == "Normal":
            normal = style
            break
    if normal is None:
        return data

    font = req.get("body_font")
    size_pt = req.get("body_font_size_pt")
    line_spacing = req.get("line_spacing")

    if font or size_pt:
        rpr = _ensure(normal, "rPr")
        if font:
            rf = _ensure(rpr, "rFonts")
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                _set_attr(rf, attr, str(font))
            changed = True
        if size_pt:
            half_points = str(round(float(size_pt) * 2))
            _set_attr(_ensure(rpr, "sz"), "val", half_points)
            _set_attr(_ensure(rpr, "szCs"), "val", half_points)
            changed = True
        transformations.append(
            Transformation("normal_style", "applied", f"Updated Normal style font={font or 'unchanged'}, size={size_pt or 'unchanged'} pt.")
        )

    if line_spacing:
        ppr = _ensure(normal, "pPr")
        spacing = _ensure(ppr, "spacing")
        # Word auto line spacing uses 240 = single, 360 = 1.5, 480 = double.
        _set_attr(spacing, "line", str(round(float(line_spacing) * 240)))
        _set_attr(spacing, "lineRule", "auto")
        changed = True
        transformations.append(Transformation("line_spacing", "applied", f"Set Normal style line spacing to {line_spacing}."))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else data


def retarget_docx(input_path: str | Path, output_path: str | Path, profile_path: str | Path) -> RetargetResult:
    src = Path(input_path)
    dst = Path(output_path)
    DocxPackage(src)  # validation
    profile = JournalProfile.from_json(profile_path)
    req = profile.requirements
    transformations: list[Transformation] = []

    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = _edit_document_xml(data, req, transformations)
            elif item.filename == "word/styles.xml":
                data = _edit_styles_xml(data, req, transformations)
            zout.writestr(item, data)

    if not transformations:
        transformations.append(Transformation("package_copy", "no-op", "No auto-fixable formatting rules were present; package was copied unchanged."))

    preservation = verify_preservation(src, dst).to_dict()
    if not preservation["passed"]:
        # Fail closed: remove unsafe output rather than leaving a deceptively usable file.
        dst.unlink(missing_ok=True)
        transformations.append(Transformation("preservation_gate", "failed", "Output was removed because the preservation audit failed."))
    else:
        transformations.append(Transformation("preservation_gate", "passed", "Protected manuscript content passed the post-transform audit."))

    return RetargetResult(
        input=str(src),
        output=str(dst),
        profile=profile.journal,
        transformations=transformations,
        preservation=preservation,
    )
