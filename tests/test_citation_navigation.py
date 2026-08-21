from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from tests.test_workflow import CONTENT_TYPES, DOC, RELS, STYLES
from word_journal_manuscript_converter.citations import build_citation_graph
from word_journal_manuscript_converter.docx_package import NS, W_NS
from word_journal_manuscript_converter.linking import link_plain_citations


def _write_docx(path: Path, document: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", STYLES)


def test_numbered_parentheses_are_detected_and_linked_bidirectionally(tmp_path: Path):
    src = tmp_path / "parenthetical.docx"
    out = tmp_path / "parenthetical_linked.docx"
    document = DOC.replace("[1]", "(1)").replace("[2]", "(2)")
    _write_docx(src, document)

    graph = build_citation_graph(src)
    assert graph.mode == "numbered"
    assert graph.citation_style == "numeric-parentheses"
    assert graph.detection_confidence >= 80

    result = link_plain_citations(src, out)
    assert result.passed
    assert result.links_added == 2
    assert result.reverse_links_added == 2
    assert result.citation_bookmarks_added == 2
    assert result.references_bookmarked == 2
    assert out.exists()

    with zipfile.ZipFile(out) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    anchors = [node.attrib.get(f"{{{W_NS}}}anchor") for node in root.findall(".//w:hyperlink", NS)]
    assert any(anchor and "WJMCRef" in anchor for anchor in anchors)
    assert any(anchor and "WJMCCite" in anchor for anchor in anchors)


def test_author_year_is_detected_and_linked_without_guessing(tmp_path: Path):
    src = tmp_path / "author_year.docx"
    out = tmp_path / "author_year_linked.docx"
    document = DOC.replace(
        "Prior work supports this result [1]. A second study agrees [2].",
        "Prior work supports this result (Smith, 2024; Jones, 2023).",
    ).replace(
        "1. Smith A. Example study. 2024.",
        "Smith A. Example study. 2024.",
    ).replace(
        "2. Jones B. Another study. 2023.",
        "Jones B. Another study. 2023.",
    )
    _write_docx(src, document)

    graph = build_citation_graph(src)
    assert graph.mode == "author-year"
    assert graph.citation_style == "author-year"
    assert graph.matched_links == 2
    assert not graph.ambiguous_citations

    result = link_plain_citations(src, out)
    assert result.passed
    assert result.citation_style == "author-year"
    assert result.links_added == 2
    assert result.reverse_links_added == 2
    assert result.references_bookmarked == 2
    assert not result.unresolved_citations


def test_ambiguous_author_year_reference_is_left_unlinked(tmp_path: Path):
    src = tmp_path / "ambiguous.docx"
    document = DOC.replace(
        "Prior work supports this result [1]. A second study agrees [2].",
        "Prior work supports this result (Smith, 2024).",
    ).replace(
        "1. Smith A. Example study. 2024.",
        "Smith A. Example study. 2024.",
    ).replace(
        "2. Jones B. Another study. 2023.",
        "Smith B. Different study. 2024.",
    )
    _write_docx(src, document)

    graph = build_citation_graph(src)
    assert graph.mode == "author-year"
    assert "smith:2024" in graph.ambiguous_citations
    assert graph.matched_links == 0


def test_numeric_group_preserves_visible_text_and_links_explicit_numbers(tmp_path: Path):
    src = tmp_path / "grouped.docx"
    out = tmp_path / "grouped_linked.docx"
    document = DOC.replace(
        "Prior work supports this result [1]. A second study agrees [2].",
        "Prior work supports this result [1, 2].",
    )
    _write_docx(src, document)

    result = link_plain_citations(src, out)
    assert result.passed
    assert result.links_added == 2
    assert result.reverse_links_added == 2

    with zipfile.ZipFile(src) as zf:
        before = ET.fromstring(zf.read("word/document.xml"))
    with zipfile.ZipFile(out) as zf:
        after = ET.fromstring(zf.read("word/document.xml"))
    before_text = "".join(t.text or "" for t in before.findall(".//w:t", NS))
    after_text = "".join(t.text or "" for t in after.findall(".//w:t", NS))
    assert before_text == after_text
