from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from tests.test_workflow import CONTENT_TYPES, DOC, RELS, STYLES, make_docx
from word_journal_manuscript_converter.docx_package import DocxPackage, NS, W_NS
from word_journal_manuscript_converter.template_mode import inspect_template, retarget_from_template


def _write_template(path: Path) -> None:
    document = DOC.replace(
        "<w:sectPr/>",
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/><w:cols w:num="1" w:space="720"/><w:lnNumType w:countBy="1" w:restart="continuous"/></w:sectPr>',
    )
    styles = STYLES.replace(
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>',
    ).replace(
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>',
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>',
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES); zf.writestr("_rels/.rels", RELS); zf.writestr("word/document.xml", document); zf.writestr("word/styles.xml", styles)


def test_template_inspection_accepts_dotx_and_reports_transferable_format(tmp_path: Path):
    template=tmp_path/"journal.dotx"; _write_template(template); report=inspect_template(template)
    assert report["workflow"]=="Template Mode"; assert report["template_type"]==".dotx"; assert report["transferable_style_count"]>=2; assert report["page_format"]["margins"]["left"]["inches"]==0.75; assert report["page_format"]["line_numbering"]["enabled"] is True


def test_template_retarget_preserves_content_and_applies_safe_format(tmp_path: Path):
    src=tmp_path/"paper.docx"; template=tmp_path/"journal.dotx"; out=tmp_path/"paper_template_retargeted.docx"; make_docx(src); _write_template(template)
    before=DocxPackage(src); before_text=before.visible_text(); before_fields=before.field_instructions(); result=retarget_from_template(src,out,template)
    assert result.passed; assert out.exists(); assert before_text==DocxPackage(out).visible_text(); assert before_fields==DocxPackage(out).field_instructions(); assert "page margins" in result.applied; assert any(x.startswith("styles:") for x in result.applied)
    with zipfile.ZipFile(out) as zf:
        document=ET.fromstring(zf.read("word/document.xml")); styles=ET.fromstring(zf.read("word/styles.xml"))
    pgmar=document.find(".//w:sectPr/w:pgMar",NS); assert pgmar is not None; assert pgmar.attrib[f"{{{W_NS}}}left"]=="1080"; assert pgmar.attrib[f"{{{W_NS}}}right"]=="1080"
    normal=next(s for s in styles.findall("w:style",NS) if s.attrib.get(f"{{{W_NS}}}styleId")=="Normal"); fonts=normal.find("w:rPr/w:rFonts",NS); assert fonts is not None; assert fonts.attrib[f"{{{W_NS}}}ascii"]=="Times New Roman"
