from __future__ import annotations

import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from tests.test_template_mode import _write_template
from tests.test_workflow import CONTENT_TYPES, DOC, RELS, STYLES, make_docx
from word_journal_manuscript_converter.audit import inspect_docx, validate_docx_structure
from word_journal_manuscript_converter.docx_package import NS, W_NS
from word_journal_manuscript_converter.linking import link_plain_numbered_citations
from word_journal_manuscript_converter.retarget import retarget_docx
from word_journal_manuscript_converter.template_mode import retarget_from_template


def _write_doc(path: Path, document: str, *, styles: str = STYLES) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", styles)


def _locals(nodes) -> list[str]:
    return [node.tag.rsplit("}", 1)[-1] for node in nodes]


def test_equation_inventory_counts_native_display_inline_and_equation_ole(tmp_path: Path):
    path = tmp_path / "equations.docx"
    document = DOC.replace(
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:o="urn:schemas-microsoft-com:office:office"',
    ).replace(
        "<w:sectPr/>",
        """
        <w:p><m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>
        <w:p><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p>
        <w:p><w:object><o:OLEObject ProgID="Equation.3"/></w:object></w:p>
        <w:sectPr/>
        """,
    )
    _write_doc(path, document)
    inv = inspect_docx(path)
    assert inv.native_equations == 2
    assert inv.embedded_equation_objects == 1
    assert inv.equations == 3
    assert inv.equation_story_parts["word/document.xml"] == 3


def test_linked_reference_bookmark_keeps_paragraph_properties_first(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "linked.docx"
    document = DOC.replace(
        '<w:p><w:r><w:t>1. Smith A. Example study. 2024.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>1. Smith A. Example study. 2024.</w:t></w:r></w:p>',
    )
    _write_doc(src, document)
    result = link_plain_numbered_citations(src, out)
    assert result.passed
    assert validate_docx_structure(out)["passed"] is True
    with zipfile.ZipFile(out) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    ref = next(
        p for p in root.findall(".//w:body/w:p", NS)
        if "1. Smith A." in "".join(t.text or "" for t in p.findall(".//w:t", NS))
    )
    assert list(ref)[0].tag == f"{{{W_NS}}}pPr"


def test_link_hyperlink_run_properties_keep_known_order(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "linked.docx"
    document = DOC.replace(
        '<w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
        '<w:p><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
    )
    _write_doc(src, document)
    result = link_plain_numbered_citations(src, out)
    assert result.passed
    with zipfile.ZipFile(out) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    linked_rpr = root.find(".//w:hyperlink/w:r/w:rPr", NS)
    assert linked_rpr is not None
    names = _locals(list(linked_rpr))
    assert names.index("color") < names.index("sz") < names.index("u")


def test_structural_validator_rejects_property_element_after_bookmark(tmp_path: Path):
    path = tmp_path / "bad-order.docx"
    document = DOC.replace(
        '<w:p><w:r><w:t>1. Smith A. Example study. 2024.</w:t></w:r></w:p>',
        '<w:p><w:bookmarkStart w:id="9" w:name="Bad"/><w:pPr><w:keepNext/></w:pPr><w:r><w:t>1. Smith A. Example study. 2024.</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>',
    )
    _write_doc(path, document)
    report = validate_docx_structure(path)
    assert report["passed"] is False
    assert any(row["check"] in {"wordprocessingml_property_position", "wordprocessingml_child_order"} and row["status"] == "fail" for row in report["checks"])


def test_structural_validator_rejects_sectpr_child_order(tmp_path: Path):
    path = tmp_path / "bad-section-order.docx"
    document = DOC.replace(
        "<w:sectPr/>",
        '<w:sectPr><w:cols w:num="1"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    )
    _write_doc(path, document)
    report = validate_docx_structure(path)
    assert report["passed"] is False
    row = next(x for x in report["checks"] if x["check"] == "wordprocessingml_child_order")
    assert row["status"] == "fail"


def test_profile_conversion_inserts_word_properties_in_schema_order(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "retargeted.docx"
    profile = tmp_path / "profile.json"
    document = DOC.replace(
        "<w:sectPr/>",
        '<w:sectPr><w:cols w:num="1"/><w:docGrid w:linePitch="360"/></w:sectPr>',
    )
    styles = STYLES.replace(
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>',
    )
    _write_doc(src, document, styles=styles)
    profile.write_text(
        json.dumps(
            {
                "journal": "Order Test",
                "article_type": "article",
                "requirements": {
                    "margins_inches": {"top": 1, "right": 1, "bottom": 1, "left": 1},
                    "body_font": "Times New Roman",
                    "body_font_size_pt": 12,
                    "line_spacing": 2.0,
                    "line_numbering": {"count_by": 1, "restart": "continuous"},
                },
            }
        ),
        encoding="utf-8",
    )
    result = retarget_docx(src, out, profile)
    assert result.passed
    assert out.exists()
    assert validate_docx_structure(out)["passed"] is True
    with zipfile.ZipFile(out) as zf:
        document_root = ET.fromstring(zf.read("word/document.xml"))
        styles_root = ET.fromstring(zf.read("word/styles.xml"))
    sect = document_root.find(".//w:sectPr", NS)
    assert sect is not None
    sect_names = _locals(list(sect))
    assert sect_names.index("pgMar") < sect_names.index("lnNumType") < sect_names.index("cols") < sect_names.index("docGrid")
    normal = next(s for s in styles_root.findall("w:style", NS) if s.attrib.get(f"{{{W_NS}}}styleId") == "Normal")
    ppr = normal.find("w:pPr", NS)
    rpr = normal.find("w:rPr", NS)
    assert ppr is not None and rpr is not None
    assert _locals(list(ppr)).index("spacing") < _locals(list(ppr)).index("jc")
    rnames = _locals(list(rpr))
    assert rnames.index("rFonts") < rnames.index("sz") < rnames.index("szCs")


def test_profile_conversion_runs_independent_assurance(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "retargeted.docx"
    profile = tmp_path / "profile.json"
    make_docx(src)
    profile.write_text(
        json.dumps(
            {
                "journal": "Assurance Test Journal",
                "article_type": "research-article",
                "checked_on": "2026-08-20",
                "requirements": {
                    "margins_inches": {"top": 1, "right": 1, "bottom": 1, "left": 1},
                    "body_font": "Times New Roman",
                    "body_font_size_pt": 12,
                    "line_spacing": 2.0,
                    "line_numbering": {"count_by": 1, "restart": "continuous"},
                    "abstract_required": True,
                    "keywords_min": 3,
                    "keywords_max": 5,
                    "required_sections": ["Introduction"],
                    "citations_must_resolve": True,
                },
            }
        ),
        encoding="utf-8",
    )
    result = retarget_docx(src, out, profile)
    assert result.passed
    assert out.exists()
    assert result.assurance["formatting_compliance_score"] == 100
    assert result.assurance["document_integrity"] == "PASS"
    assert result.assurance["structural_sanity"] == "PASS"
    assert result.assurance["blocking_failures"] == 0
    assert result.assurance["verdict"] == "READY FOR FINAL AUTHOR REVIEW"


def test_template_conversion_reports_supported_fidelity_and_coverage(tmp_path: Path):
    src = tmp_path / "paper.docx"
    template = tmp_path / "journal.dotx"
    out = tmp_path / "paper_template_retargeted.docx"
    make_docx(src)
    _write_template(template)
    result = retarget_from_template(src, out, template)
    assert result.passed
    assert out.exists()
    assert result.fidelity["supported_fidelity_score"] == 100
    assert result.fidelity["blocking_failures"] == 0
    assert result.fidelity["exact_visual_match_guaranteed"] is False
    assert result.structural_validation["passed"] is True
