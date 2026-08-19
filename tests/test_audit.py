from __future__ import annotations

import json
import zipfile
from pathlib import Path

from word_journal_manuscript_converter.audit import inspect_docx, verify_preservation
from word_journal_manuscript_converter.journal import readiness_check


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
 <w:body>
  <w:p><w:r><w:t>Result 12.5 was reported</w:t></w:r></w:p>
  <w:p>
   <w:fldSimple w:instr=" ADDIN EN.CITE &lt;EndNote&gt; "><w:r><w:t>(Smith, 2024)</w:t></w:r></w:fldSimple>
  </w:p>
  <w:p>
   <w:bookmarkStart w:id="0" w:name="ref_smith_2024"/>
   <w:hyperlink w:anchor="ref_smith_2024"><w:r><w:t>Smith 2024</w:t></w:r></w:hyperlink>
   <w:bookmarkEnd w:id="0"/>
  </w:p>
  <w:p><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>
  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:sectPr/>
 </w:body>
</w:document>'''


def make_docx(path: Path, doc_xml: str = DOC, media: bytes | None = b"image-bytes") -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", doc_xml)
        if media is not None:
            zf.writestr("word/media/image1.png", media)


def test_inspect_detects_sensitive_features(tmp_path: Path) -> None:
    p = tmp_path / "paper.docx"
    make_docx(p)
    inv = inspect_docx(p)
    assert inv.valid_docx
    assert inv.tables == 1
    assert inv.sections == 1
    assert inv.images == 1
    assert inv.equations >= 1
    assert inv.citation.endnote_fields == 1
    assert inv.citation.total_candidate_fields == 1
    assert inv.citation.bookmarks == 1
    assert inv.citation.internal_hyperlinks == 1


def test_verify_identical_copy_passes(tmp_path: Path) -> None:
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    make_docx(a)
    make_docx(b)
    report = verify_preservation(a, b)
    assert report.passed


def test_verify_detects_scientific_number_change(tmp_path: Path) -> None:
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    make_docx(a)
    make_docx(b, DOC.replace("12.5", "13.5"))
    report = verify_preservation(a, b)
    assert not report.passed
    assert not report.text_identical
    assert not report.numeric_tokens_identical


def test_readiness_profile(tmp_path: Path) -> None:
    p = tmp_path / "paper.docx"
    make_docx(p)
    profile = tmp_path / "journal.json"
    profile.write_text(json.dumps({
        "journal": "Test Journal",
        "article_type": "research-article",
        "requirements": {
            "requires_live_citations": True,
            "tracked_changes_allowed": False,
            "max_comments": 0
        }
    }), encoding="utf-8")
    result = readiness_check(p, profile)
    statuses = {c["check"]: c["status"] for c in result["checks"]}
    assert statuses["live_citations"] == "pass"
    assert statuses["tracked_changes"] == "pass"
    assert statuses["comments"] == "pass"
