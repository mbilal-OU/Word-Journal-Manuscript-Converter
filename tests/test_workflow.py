from __future__ import annotations

import json
import zipfile
from pathlib import Path

from word_journal_manuscript_converter.citations import build_citation_graph
from word_journal_manuscript_converter.linking import link_plain_numbered_citations
from word_journal_manuscript_converter.retarget import retarget_docx
from word_journal_manuscript_converter.structure import extract_structure

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
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Abstract</w:t></w:r></w:p>
  <w:p><w:r><w:t>This is a short abstract with 12.5 units.</w:t></w:r></w:p>
  <w:p><w:r><w:t>Keywords: alpha; beta; gamma</w:t></w:r></w:p>
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


def make_docx(path: Path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", DOC)
        zf.writestr("word/styles.xml", STYLES)


def test_structure_and_citation_graph(tmp_path: Path):
    p = tmp_path / "paper.docx"
    make_docx(p)
    s = extract_structure(p)
    assert s.abstract_word_count > 0
    assert s.keywords == ["alpha", "beta", "gamma"]
    assert "References" in s.headings
    g = build_citation_graph(p)
    assert g.mode == "numbered"
    assert g.in_text_citation_count == 2
    assert g.matched_links == 2
    assert not g.unmatched_citations


def test_link_plain_numbered_citations_preserves_content(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "linked.docx"
    make_docx(src)
    result = link_plain_numbered_citations(src, out)
    assert result.passed
    assert result.links_added == 2
    assert result.references_bookmarked == 2
    assert out.exists()


def test_safe_retarget_preserves_content(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "retargeted.docx"
    make_docx(src)
    profile = tmp_path / "profile.json"
    profile.write_text(json.dumps({
        "journal": "Test",
        "article_type": "research-article",
        "requirements": {
            "margins_inches": {"top": 1, "right": 1, "bottom": 1, "left": 1},
            "body_font": "Times New Roman",
            "body_font_size_pt": 12,
            "line_spacing": 2.0,
            "line_numbering": {"count_by": 1, "restart": "continuous"}
        }
    }), encoding="utf-8")
    result = retarget_docx(src, out, profile)
    assert result.passed
    assert out.exists()
    assert any(t.name == "page_margins" for t in result.transformations)
    assert any(t.name == "normal_style" for t in result.transformations)


def test_bundled_profiles_are_valid():
    from word_journal_manuscript_converter.profiles import list_bundled_profiles, load_profile_data, validate_profile_data

    profiles = list_bundled_profiles()
    keys = {p.key for p in profiles}
    assert "plos-one-research-article" in keys
    assert "scientific-reports-article" in keys
    assert "frontiers-microbiology-original-research" in keys
    for desc in profiles:
        data, resolved = load_profile_data(desc.key)
        assert resolved.startswith("bundled:")
        assert not validate_profile_data(data)


def test_full_analysis_and_html_report(tmp_path: Path):
    from word_journal_manuscript_converter.reporting import analyze_manuscript, render_html_report

    p = tmp_path / "paper.docx"
    make_docx(p)
    report = analyze_manuscript(p, "generic-review-copy")
    assert report["version"] == "0.3.2"
    assert report["structure"]["reference_count"] == 2
    assert report["citation_graph"]["matched_links"] == 2
    assert report["readiness"]["journal"] == "Generic review-copy profile"
    html = render_html_report(report)
    assert "Word Journal Manuscript Converter" in html
    assert "Journal readiness" in html
    assert "Test manuscript" not in html


def test_bundled_profile_ref_is_accepted_by_retarget(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "retargeted.docx"
    make_docx(src)
    result = retarget_docx(src, out, "generic-review-copy")
    assert result.passed
    assert out.exists()


def test_bundled_profiles_fall_back_when_resource_directory_missing(tmp_path, monkeypatch):
    import word_journal_manuscript_converter.profiles as profiles

    monkeypatch.setattr(profiles, "_bundled_root", lambda: tmp_path / "missing-bundled-profiles")
    keys = profiles.bundled_profile_keys()
    assert "scientific-reports-article" in keys
    data, resolved = profiles.load_profile_data("scientific-reports-article")
    assert data["journal"] == "Scientific Reports"
    assert resolved == "bundled:scientific-reports-article"


def test_reference_extraction_does_not_stop_on_accidental_heading_style(tmp_path: Path):
    p = tmp_path / "styled-reference.docx"
    doc = DOC.replace(
        '<w:p><w:r><w:t>2. Jones B. Another study. 2023.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>2. Jones B. Another study. 2023.</w:t></w:r></w:p>'
    )
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/styles.xml", STYLES)
    structure = extract_structure(p)
    assert len(structure.reference_paragraphs) == 2
    graph = build_citation_graph(p)
    assert graph.reference_count == 2
    assert not graph.unmatched_citations


def test_long_prose_with_heading_style_is_not_structural_heading(tmp_path: Path):
    p = tmp_path / "long-heading-style.docx"
    prose = "This is ordinary manuscript prose " + "with several words " * 35
    doc = DOC.replace(
        '<w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{prose}</w:t></w:r></w:p>'
    )
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/styles.xml", STYLES)
    structure = extract_structure(p)
    assert prose not in structure.headings
