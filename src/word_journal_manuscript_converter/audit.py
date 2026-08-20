from __future__ import annotations

import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from .docx_package import DocxPackage, NS, W_NS
from .models import CitationInventory, DocxInventory, PreservationReport


def _count_nodes(package: DocxPackage, xpath: str) -> int:
    total = 0
    for part in package.story_parts():
        total += len(package.xml(part).findall(xpath, NS))
    return total


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _citation_inventory(package: DocxPackage) -> CitationInventory:
    instructions = package.field_instructions()
    upper = [x.upper() for x in instructions]
    citation_like = [
        x for x in upper
        if any(k in x for k in ("EN.CITE", "ZOTERO_ITEM", "CSL_CITATION", "MENDELEY CITATION"))
    ]
    bibliography_like = [
        x for x in upper
        if any(k in x for k in ("EN.REFLIST", "ZOTERO_BIBL", "CSL_BIBLIOGRAPHY", "BIBLIOGRAPHY"))
    ]
    ref_fields = [x for x in upper if re.match(r"^(REF|PAGEREF)\b", x)]

    bookmarks = _count_nodes(package, ".//w:bookmarkStart")
    internal_hyperlinks = 0
    for part in package.story_parts():
        root = package.xml(part)
        for h in root.findall(".//w:hyperlink", NS):
            if f"{{{W_NS}}}anchor" in h.attrib:
                internal_hyperlinks += 1

    return CitationInventory(
        total_candidate_fields=len(citation_like),
        endnote_fields=sum("EN.CITE" in x for x in upper),
        zotero_fields=sum("ZOTERO_ITEM" in x for x in upper),
        csl_fields=sum("CSL_CITATION" in x or "MENDELEY CITATION" in x for x in upper),
        bibliography_fields=len(bibliography_like),
        internal_hyperlinks=internal_hyperlinks,
        bookmarks=bookmarks,
        ref_fields=len(ref_fields),
    )


def _equation_inventory(package: DocxPackage) -> dict:
    """Count native OMML equations and known equation-editor OLE objects.

    Native display equations are represented by m:oMathPara and can contain one
    or more nested m:oMath nodes. They are counted as one displayed equation,
    while standalone m:oMath nodes are counted as inline equations. Known legacy
    Equation Editor/MathType OLE objects are counted separately.
    """
    native = 0
    embedded_equations = 0
    parts: dict[str, int] = {}

    for part in package.story_parts():
        root = package.xml(part)
        math_paras = root.findall(".//m:oMathPara", NS)
        nested_math_ids = {
            id(node)
            for para in math_paras
            for node in para.findall(".//m:oMath", NS)
        }
        all_math = root.findall(".//m:oMath", NS)
        standalone_math = sum(id(node) not in nested_math_ids for node in all_math)
        part_native = len(math_paras) + standalone_math

        part_embedded = 0
        for node in root.iter():
            if _local_name(node.tag).lower() != "oleobject":
                continue
            metadata = " ".join(str(value) for value in node.attrib.values()).casefold()
            if any(marker in metadata for marker in ("equation", "mathtype", "mathcad")):
                part_embedded += 1

        if part_native or part_embedded:
            parts[part] = part_native + part_embedded
        native += part_native
        embedded_equations += part_embedded

    embedded_objects = sum(
        1 for name in package.names()
        if name.startswith("word/embeddings/") and not name.endswith("/")
    )
    return {
        "native_equations": native,
        "embedded_equation_objects": embedded_equations,
        "equations": native + embedded_equations,
        "equation_story_parts": parts,
        "embedded_objects": embedded_objects,
    }


def validate_docx_structure(path: str | Path) -> dict:
    """Run defensive OOXML package checks after a transformation.

    This is not a complete ECMA-376/XSD validator and does not claim that Word
    will accept every possible feature combination. It catches common package
    damage that our transformations can create: malformed XML, misplaced
    property elements, and unbalanced/duplicate bookmarks.
    """
    checks: list[dict] = []
    failures: list[str] = []
    warnings: list[str] = []

    try:
        package = DocxPackage(path)
    except Exception as exc:
        return {
            "passed": False,
            "checks": [{"check": "docx_package", "status": "fail", "detail": str(exc)}],
            "failures": [str(exc)],
            "warnings": warnings,
            "scope": "defensive OOXML structural checks; not full Microsoft Word validation",
        }

    names = package.names()
    required = {"[Content_Types].xml", "word/document.xml"}
    missing = sorted(required.difference(names))
    if missing:
        failures.append("Missing required package parts: " + ", ".join(missing))
    checks.append({
        "check": "required_parts",
        "status": "fail" if missing else "pass",
        "detail": "Required DOCX package parts are present." if not missing else failures[-1],
    })

    malformed: list[str] = []
    with zipfile.ZipFile(package.path) as zf:
        for name in names:
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                ET.fromstring(zf.read(name))
            except (ET.ParseError, KeyError) as exc:
                malformed.append(f"{name}: {exc}")
    if malformed:
        failures.extend("Malformed XML: " + item for item in malformed)
    checks.append({
        "check": "xml_parseability",
        "status": "fail" if malformed else "pass",
        "detail": "All XML and relationship parts parse successfully." if not malformed else f"Malformed XML parts: {len(malformed)}.",
    })

    ordering_violations = 0
    bookmark_mismatches = 0
    duplicate_bookmark_names = 0
    q_id = f"{{{W_NS}}}id"
    q_name = f"{{{W_NS}}}name"

    for part in package.story_parts():
        root = package.xml(part)
        structural_pairs = (
            (".//w:p", "pPr"),
            (".//w:tbl", "tblPr"),
            (".//w:tr", "trPr"),
            (".//w:tc", "tcPr"),
        )
        for xpath, property_name in structural_pairs:
            for container in root.findall(xpath, NS):
                prop = container.find(f"w:{property_name}", NS)
                children = list(container)
                if prop is not None and children and children[0] is not prop:
                    ordering_violations += 1

        starts = root.findall(".//w:bookmarkStart", NS)
        ends = root.findall(".//w:bookmarkEnd", NS)
        start_ids = Counter(node.attrib.get(q_id, "") for node in starts)
        end_ids = Counter(node.attrib.get(q_id, "") for node in ends)
        if start_ids != end_ids:
            bookmark_mismatches += 1
        names_in_part = [node.attrib.get(q_name, "") for node in starts if node.attrib.get(q_name)]
        duplicate_bookmark_names += sum(count - 1 for count in Counter(names_in_part).values() if count > 1)

    if ordering_violations:
        failures.append(f"Detected {ordering_violations} OOXML property-order violations.")
    checks.append({
        "check": "wordprocessingml_property_order",
        "status": "fail" if ordering_violations else "pass",
        "detail": "Paragraph/table property elements are in a valid leading position." if not ordering_violations else failures[-1],
    })

    if bookmark_mismatches:
        failures.append(f"Detected unbalanced bookmark start/end identifiers in {bookmark_mismatches} story parts.")
    if duplicate_bookmark_names:
        failures.append(f"Detected {duplicate_bookmark_names} duplicate bookmark names.")
    checks.append({
        "check": "bookmark_integrity",
        "status": "fail" if bookmark_mismatches or duplicate_bookmark_names else "pass",
        "detail": "Bookmark starts/ends are balanced and names are unique." if not bookmark_mismatches and not duplicate_bookmark_names else "Bookmark integrity checks failed.",
    })

    warnings.append(
        "These are defensive package checks, not a complete Microsoft Word rendering or schema-conformance test. Final visual review in Word is still required."
    )
    return {
        "passed": not failures,
        "checks": checks,
        "failures": failures,
        "warnings": warnings,
        "scope": "defensive OOXML structural checks; not full Microsoft Word validation",
    }


def inspect_docx(path: str | Path) -> DocxInventory:
    package = DocxPackage(path)
    names = package.names()
    document = package.xml("word/document.xml")

    paragraphs = len(document.findall(".//w:p", NS))
    tables = len(document.findall(".//w:tbl", NS))
    sections = len(document.findall(".//w:sectPr", NS))
    images = sum(1 for n in names if n.startswith("word/media/") and not n.endswith("/"))
    equation_inventory = _equation_inventory(package)
    comments = 0
    if package.has("word/comments.xml"):
        comments = len(package.xml("word/comments.xml").findall(".//w:comment", NS))
    footnotes = 0
    if package.has("word/footnotes.xml"):
        footnotes = len(package.xml("word/footnotes.xml").findall(".//w:footnote", NS))
    endnotes = 0
    if package.has("word/endnotes.xml"):
        endnotes = len(package.xml("word/endnotes.xml").findall(".//w:endnote", NS))

    tracked_insertions = _count_nodes(package, ".//w:ins")
    tracked_deletions = _count_nodes(package, ".//w:del")
    fields = len(package.field_instructions())
    hyperlinks = _count_nodes(package, ".//w:hyperlink")
    custom_xml_parts = sum(1 for n in names if n.startswith("customXml/") and n.endswith(".xml"))
    citation = _citation_inventory(package)

    warnings: list[str] = []
    if citation.total_candidate_fields == 0:
        warnings.append("No live EndNote/Zotero/CSL citation fields were detected; citations may be plain text or use an unsupported field encoding.")
    if tracked_insertions or tracked_deletions:
        warnings.append("Tracked changes are present. Retargeting must preserve review markup unless the user explicitly requests otherwise.")
    if fields and citation.total_candidate_fields == 0:
        warnings.append("Word fields are present but none matched known citation-manager signatures.")
    if equation_inventory["embedded_equation_objects"]:
        warnings.append(
            f"Detected {equation_inventory['embedded_equation_objects']} legacy/embedded equation-editor object(s) in addition to native Word equations."
        )
    if equation_inventory["equations"] == 0 and images:
        warnings.append(
            "No native or known embedded equation objects were detected. Equations inserted as ordinary pictures cannot be distinguished reliably from figures by the current structural audit."
        )

    return DocxInventory(
        path=str(Path(path)),
        valid_docx=True,
        paragraphs=paragraphs,
        tables=tables,
        sections=sections,
        images=images,
        equations=equation_inventory["equations"],
        native_equations=equation_inventory["native_equations"],
        embedded_equation_objects=equation_inventory["embedded_equation_objects"],
        equation_story_parts=equation_inventory["equation_story_parts"],
        embedded_objects=equation_inventory["embedded_objects"],
        comments=comments,
        footnotes=footnotes,
        endnotes=endnotes,
        tracked_insertions=tracked_insertions,
        tracked_deletions=tracked_deletions,
        fields=fields,
        hyperlinks=hyperlinks,
        custom_xml_parts=custom_xml_parts,
        citation=citation,
        warnings=warnings,
    )


def verify_preservation(before: str | Path, after: str | Path) -> PreservationReport:
    a = DocxPackage(before)
    b = DocxPackage(after)
    ia = inspect_docx(before)
    ib = inspect_docx(after)

    text_a = a.visible_text()
    text_b = b.visible_text()
    fields_a = Counter(a.field_instructions())
    fields_b = Counter(b.field_instructions())
    media_a = Counter(a.media_hashes().values())
    media_b = Counter(b.media_hashes().values())
    custom_a = Counter(a.custom_xml_hashes().values())
    custom_b = Counter(b.custom_xml_hashes().values())
    rels_a = a.relationship_hashes()
    rels_b = b.relationship_hashes()

    return PreservationReport(
        before=str(before),
        after=str(after),
        text_identical=text_a == text_b,
        numeric_tokens_identical=a.numeric_tokens() == b.numeric_tokens(),
        field_instructions_identical=fields_a == fields_b,
        citation_field_count_identical=ia.citation.total_candidate_fields == ib.citation.total_candidate_fields,
        media_hashes_identical=media_a == media_b,
        custom_xml_hashes_identical=custom_a == custom_b,
        relationship_hashes_identical=rels_a == rels_b,
        content_types_identical=a.content_types_hash() == b.content_types_hash(),
        tracked_changes_identical=(ia.tracked_insertions, ia.tracked_deletions) == (ib.tracked_insertions, ib.tracked_deletions),
        comments_identical=ia.comments == ib.comments,
        notes_identical=(ia.footnotes, ia.endnotes) == (ib.footnotes, ib.endnotes),
        equations_identical=(
            ia.equations,
            ia.native_equations,
            ia.embedded_equation_objects,
        ) == (
            ib.equations,
            ib.native_equations,
            ib.embedded_equation_objects,
        ),
        tables_identical=ia.tables == ib.tables,
        bookmarks_not_lost=ib.citation.bookmarks >= ia.citation.bookmarks,
        hyperlinks_not_lost=ib.hyperlinks >= ia.hyperlinks,
        details={
            "before": ia.to_dict(),
            "after": ib.to_dict(),
            "text_length_before": len(text_a),
            "text_length_after": len(text_b),
            "numeric_token_count_before": len(a.numeric_tokens()),
            "numeric_token_count_after": len(b.numeric_tokens()),
        },
    )
