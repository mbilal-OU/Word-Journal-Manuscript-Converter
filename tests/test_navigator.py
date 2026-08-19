from __future__ import annotations

import zipfile
from pathlib import Path

from word_journal_manuscript_converter.navigator import (
    analyze_citation_navigation,
    make_navigable_copy,
    render_navigation_html,
)
from tests.test_workflow import CONTENT_TYPES, DOC, RELS, STYLES, make_docx


def _write_docx(path: Path, document_xml: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", STYLES)


def test_plain_numbered_navigation_supports_clickable_copy(tmp_path: Path):
    src = tmp_path / "paper.docx"
    out = tmp_path / "paper_navigable.docx"
    make_docx(src)

    report = analyze_citation_navigation(src)
    assert report["citation_manager"] == "None detected"
    assert report["live_fields"] is False
    assert report["navigation_strategy"] == "clickable-docx-export"
    assert report["citation_graph"]["matched_links"] == 2

    result = make_navigable_copy(src, out)
    assert result["created"] is True
    assert out.exists()


def test_live_endnote_document_is_not_rewritten(tmp_path: Path):
    src = tmp_path / "endnote.docx"
    out = tmp_path / "endnote_navigable.docx"
    live_doc = DOC.replace(
        '<w:p><w:r><w:t>Prior work supports this result [1]. A second study agrees [2].</w:t></w:r></w:p>',
        '<w:p><w:fldSimple w:instr="ADDIN EN.CITE DATA"><w:r><w:t>[1]</w:t></w:r></w:fldSimple><w:r><w:t> and supporting work [2].</w:t></w:r></w:p>'
    )
    _write_docx(src, live_doc)

    report = analyze_citation_navigation(src)
    assert report["citation_manager"] == "EndNote"
    assert report["live_fields"] is True
    assert report["navigation_strategy"] == "live-safe-word-navigation"

    result = make_navigable_copy(src, out)
    assert result["created"] is False
    assert result["mode"] == "live-safe-word-navigation"
    assert not out.exists()


def test_navigation_html_has_internal_reference_links(tmp_path: Path):
    src = tmp_path / "paper.docx"
    make_docx(src)
    report = analyze_citation_navigation(src)
    html = render_navigation_html(report)
    assert 'href="#ref-1"' in html
    assert 'id="ref-1"' in html
    assert "Citation Navigator" in html
