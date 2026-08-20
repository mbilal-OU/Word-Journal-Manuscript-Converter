from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from word_journal_manuscript_converter.docx_package import (
    DocxError,
    DocxPackage,
    missing_compatibility_prefixes,
    parse_xml_preserving_namespaces,
    serialize_xml_preserving_namespaces,
)

CONTENT_TYPES = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

RELS = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''


def _document(*, declare_w14: bool) -> bytes:
    extra = ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' if declare_w14 else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        f'{extra} mc:Ignorable="w14">'
        '<w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p><w:sectPr/></w:body>'
        '</w:document>'
    ).encode("utf-8")


def _write_docx(path: Path, document: bytes) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document)


def test_roundtrip_keeps_ignorable_namespace_declaration() -> None:
    original = _document(declare_w14=True)
    assert missing_compatibility_prefixes(original) == []
    root = parse_xml_preserving_namespaces(original)
    rewritten = serialize_xml_preserving_namespaces(root, original)
    assert b'mc:Ignorable="w14"' in rewritten
    assert b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' in rewritten
    assert missing_compatibility_prefixes(rewritten) == []


def test_docx_package_rejects_dangling_ignorable_prefix(tmp_path: Path) -> None:
    path = tmp_path / "dangling.docx"
    _write_docx(path, _document(declare_w14=False))
    with pytest.raises(DocxError, match="undeclared markup-compatibility namespace prefix"):
        DocxPackage(path)


def test_docx_package_accepts_declared_ignorable_prefix(tmp_path: Path) -> None:
    path = tmp_path / "valid.docx"
    _write_docx(path, _document(declare_w14=True))
    package = DocxPackage(path)
    assert package.visible_text().strip() == "Test"
