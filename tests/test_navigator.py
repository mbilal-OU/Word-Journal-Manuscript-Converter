from __future__ import annotations

import zipfile
from pathlib import Path

from word_journal_manuscript_converter.audit import inspect_docx
from word_journal_manuscript_converter.docx_package import DocxPackage
from word_journal_manuscript_converter.navigator import (
    analyze_citation_navigation,
    make_navigable_copy,
    render_navigation_html,
)

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Test manuscript</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
  <w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>References</w:t></w:r></w:p>
  <w:p><w:r><w:t>1. Smith A. Example study. 2024.</w:t></w:r></w:p>
  <w:p><w:r><w:t>2. Jones B. Another study. 2023.</w:t></w:r></w:p>
  <w:sectPr/>
 </w:body>
</w:document>'''
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
 <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
</w:styles>'''


def _write_docx(path: Path, document_xml: str = DOC) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", STYLES)


def _simple_endnote_doc() -> str:
    return DOC.replace(
        '<w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
        '<w:p><w:r><w:t>Prior work supports this result </w:t></w:r>'
        '<w:fldSimple w:instr="ADDIN EN.CITE DATA"><w:r><w:t>[1]</w:t></w:r></w:fldSimple>'
        '<w:r><w:t>. A second study agrees [2].</w:t></w:r></w:p>',
    )


def _complex_endnote_doc() -> str:
    return DOC.replace(
        '<w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
        '<w:p><w:r><w:t>Prior work supports this result </w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> ADDIN EN.CITE DATA </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>[1]</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t>. A second study agrees [2].</w:t></w:r></w:p>',
    )


def test_plain_numbered_navigation_supports_clickable_copy(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "paper_navigable.docx"
    _write_docx(src)

    report = analyze_citation_navigation(src)
    assert report["citation_manager"] == "None detected"
    assert report["live_fields"] is False
    assert report["navigation_strategy"] == "clickable-docx-export"
    assert report["citation_graph"]["matched_links"] == 2

    result = make_navigable_copy(src, out)
    assert result["created"] is True
    assert out.exists()


def test_live_endnote_document_requires_explicit_static_review_opt_in(tmp_path: Path):
    src = tmp_path / "endnote.docx"
    out = tmp_path / "endnote_navigable.docx"
    _write_docx(src, _simple_endnote_doc())

    report = analyze_citation_navigation(src)
    assert report["citation_manager"] == "EndNote"
    assert report["live_fields"] is True
    assert report["navigation_strategy"] == "live-safe-or-static-review"

    result = make_navigable_copy(src, out)
    assert result["created"] is False
    assert result["mode"] == "live-safe-word-navigation"
    assert not out.exists()


def test_live_endnote_can_create_separate_static_linked_review_copy(tmp_path: Path):
    src = tmp_path / "endnote.docx"
    out = tmp_path / "endnote_linked_review_copy.docx"
    _write_docx(src, _simple_endnote_doc())

    before_text = DocxPackage(src).visible_text()
    result = make_navigable_copy(src, out, static_review_copy=True)

    assert result["created"] is True
    assert result["mode"] == "linked-review-copy"
    assert result["citation_manager_master_preserved"] is True
    assert result["links_added"] >= 2
    assert out.exists()
    assert DocxPackage(src).visible_text() == before_text
    assert DocxPackage(out).visible_text() == before_text
    assert inspect_docx(src).citation.endnote_fields == 1
    assert inspect_docx(out).citation.endnote_fields == 0
    assert result["preservation"]["passed"] is True


def test_complex_endnote_field_is_flattened_only_in_review_copy(tmp_path: Path):
    src = tmp_path / "endnote_complex.docx"
    out = tmp_path / "endnote_complex_review.docx"
    _write_docx(src, _complex_endnote_doc())

    before_text = DocxPackage(src).visible_text()
    result = make_navigable_copy(src, out, static_review_copy=True)

    assert result["created"] is True
    assert out.exists()
    assert DocxPackage(out).visible_text() == before_text
    assert inspect_docx(src).citation.endnote_fields == 1
    assert inspect_docx(out).citation.endnote_fields == 0
    assert result["flattening"]["complex_fields_flattened"] == 1


def test_navigation_html_has_internal_reference_links(tmp_path: Path):
    src = tmp_path / "paper.docx"
    _write_docx(src)
    report = analyze_citation_navigation(src)
    html = render_navigation_html(report)
    assert 'href="#ref-1"' in html
    assert 'id="ref-1"' in html
    assert "Citation Navigator" in html


def test_navigation_html_explains_static_review_option_for_live_fields(tmp_path: Path):
    src = tmp_path / "endnote.docx"
    _write_docx(src, _simple_endnote_doc())
    report = analyze_citation_navigation(src)
    html = render_navigation_html(report)
    assert "Live citation manager detected" in html
    assert "static linked review copy" in html
