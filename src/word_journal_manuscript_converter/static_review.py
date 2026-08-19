from __future__ import annotations

import copy
import tempfile
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from .audit import inspect_docx
from .docx_package import DocxPackage, NS, W_NS
from .linking import link_plain_numbered_citations

ET.register_namespace("w", W_NS)

_MANAGER_MARKERS = (
    "EN.CITE",
    "EN.REFLIST",
    "ZOTERO_ITEM",
    "ZOTERO_BIBL",
    "CSL_CITATION",
    "CSL_BIBLIOGRAPHY",
    "MENDELEY CITATION",
    "MENDELEY_BIBLIOGRAPHY",
)


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def is_manager_instruction(instruction: str) -> bool:
    upper = " ".join(instruction.upper().split())
    return any(marker in upper for marker in _MANAGER_MARKERS)


def _field_char_type(node: ET.Element) -> str | None:
    fld = node.find(".//w:fldChar", NS)
    if fld is None:
        return None
    return fld.attrib.get(_q("fldCharType"))


def _instruction_text(node: ET.Element) -> str:
    return " ".join(
        " ".join((part.text or "").split())
        for part in node.findall(".//w:instrText", NS)
        if part.text
    ).strip()


def _flatten_simple_fields(root: ET.Element) -> int:
    flattened = 0
    for parent in list(root.iter()):
        for child in list(parent):
            if child.tag != _q("fldSimple"):
                continue
            instruction = child.attrib.get(_q("instr"), "")
            if not is_manager_instruction(instruction):
                continue
            index = list(parent).index(child)
            replacement = [copy.deepcopy(node) for node in list(child)]
            parent.remove(child)
            for offset, node in enumerate(replacement):
                parent.insert(index + offset, node)
            flattened += 1
    return flattened


def _flatten_complex_fields(root: ET.Element) -> int:
    """Flatten complex citation-manager fields while keeping their visible results.

    This state machine works across paragraph boundaries, which matters for
    bibliography fields whose begin/instruction/separate/end controls may span
    more than one Word paragraph.
    """
    flattened = 0
    stack: list[dict] = []

    for paragraph in root.findall(".//w:p", NS):
        for child in list(paragraph):
            field_type = _field_char_type(child)

            if field_type == "begin":
                stack.append(
                    {
                        "controls": [(paragraph, child)],
                        "instruction_parts": [],
                    }
                )
                continue

            if not stack:
                continue

            current = stack[-1]
            instruction = _instruction_text(child)
            if instruction:
                current["instruction_parts"].append(instruction)
                current["controls"].append((paragraph, child))

            if field_type == "separate":
                current["controls"].append((paragraph, child))
                continue

            if field_type == "end":
                current["controls"].append((paragraph, child))
                completed = stack.pop()
                full_instruction = " ".join(completed["instruction_parts"]).strip()

                if is_manager_instruction(full_instruction):
                    for parent, control in completed["controls"]:
                        if control in list(parent):
                            parent.remove(control)
                    flattened += 1

    return flattened


def flatten_manager_fields(input_path: str | Path, output_path: str | Path) -> dict:
    src = Path(input_path)
    dst = Path(output_path)
    if src.resolve() == dst.resolve():
        raise ValueError("The linked review copy must use a different output path from the original manuscript.")

    package = DocxPackage(src)
    replacements: dict[str, bytes] = {}
    simple_count = 0
    complex_count = 0

    for part in package.story_parts():
        root = package.xml(part)
        simple_part = _flatten_simple_fields(root)
        complex_part = _flatten_complex_fields(root)
        simple_count += simple_part
        complex_count += complex_part
        if simple_part or complex_part:
            replacements[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    total = simple_count + complex_count
    if total == 0:
        raise ValueError("No supported live citation-manager fields were found to flatten in the review copy.")

    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)

    before = DocxPackage(src)
    after = DocxPackage(dst)
    if before.visible_text() != after.visible_text() or before.numeric_tokens() != after.numeric_tokens():
        dst.unlink(missing_ok=True)
        raise ValueError(
            "Static review-copy conversion changed visible manuscript content, so the output was removed."
        )

    remaining = [x for x in after.field_instructions() if is_manager_instruction(x)]
    if remaining:
        dst.unlink(missing_ok=True)
        raise ValueError(
            "Some citation-manager fields could not be safely flattened, so the review copy was removed."
        )

    return {
        "simple_fields_flattened": simple_count,
        "complex_fields_flattened": complex_count,
        "manager_fields_flattened": total,
    }


def verify_static_review_copy(before_path: str | Path, after_path: str | Path) -> dict:
    before = DocxPackage(before_path)
    after = DocxPackage(after_path)
    before_inventory = inspect_docx(before_path)
    after_inventory = inspect_docx(after_path)

    before_fields = before.field_instructions()
    after_fields = after.field_instructions()
    before_manager = [x for x in before_fields if is_manager_instruction(x)]
    after_manager = [x for x in after_fields if is_manager_instruction(x)]
    before_other = Counter(x for x in before_fields if not is_manager_instruction(x))
    after_other = Counter(x for x in after_fields if not is_manager_instruction(x))

    checks = {
        "visible_text_identical": before.visible_text() == after.visible_text(),
        "numeric_tokens_identical": before.numeric_tokens() == after.numeric_tokens(),
        "manager_fields_removed": bool(before_manager) and not after_manager,
        "non_manager_fields_identical": before_other == after_other,
        "media_hashes_identical": Counter(before.media_hashes().values()) == Counter(after.media_hashes().values()),
        "custom_xml_hashes_identical": Counter(before.custom_xml_hashes().values()) == Counter(after.custom_xml_hashes().values()),
        "relationship_hashes_identical": before.relationship_hashes() == after.relationship_hashes(),
        "content_types_identical": before.content_types_hash() == after.content_types_hash(),
        "tracked_changes_identical": (
            before_inventory.tracked_insertions,
            before_inventory.tracked_deletions,
        ) == (
            after_inventory.tracked_insertions,
            after_inventory.tracked_deletions,
        ),
        "comments_identical": before_inventory.comments == after_inventory.comments,
        "notes_identical": (
            before_inventory.footnotes,
            before_inventory.endnotes,
        ) == (
            after_inventory.footnotes,
            after_inventory.endnotes,
        ),
        "equations_identical": before_inventory.equations == after_inventory.equations,
        "tables_identical": before_inventory.tables == after_inventory.tables,
        "bookmarks_not_lost": after_inventory.citation.bookmarks >= before_inventory.citation.bookmarks,
        "hyperlinks_not_lost": after_inventory.hyperlinks >= before_inventory.hyperlinks,
    }

    return {
        "passed": all(checks.values()),
        "checks": checks,
        "manager_fields_before": len(before_manager),
        "manager_fields_after": len(after_manager),
        "citation_fields_before": before_inventory.citation.total_candidate_fields,
        "citation_fields_after": after_inventory.citation.total_candidate_fields,
        "warning": (
            "This is a static review/navigation copy. Live citation-manager fields were intentionally "
            "removed from the copy only. Keep the original manuscript as the editable EndNote/Zotero/Mendeley master."
        ),
    }


def create_linked_review_copy(input_path: str | Path, output_path: str | Path) -> dict:
    src = Path(input_path)
    dst = Path(output_path)
    if src.resolve() == dst.resolve():
        raise ValueError("The linked review copy must use a different output path from the original manuscript.")

    with tempfile.TemporaryDirectory(prefix="wjmc-review-") as tmpdir:
        flattened = Path(tmpdir) / "flattened.docx"
        flatten_report = flatten_manager_fields(src, flattened)
        link_result = link_plain_numbered_citations(flattened, dst)

    if not dst.exists():
        return {
            "created": False,
            "mode": "linked-review-copy",
            "input": str(src),
            "output": None,
            "message": "The static review copy could not be created because the link-stage preservation audit failed.",
            "flattening": flatten_report,
            "linking": link_result.to_dict(),
        }

    verification = verify_static_review_copy(src, dst)
    if not verification["passed"]:
        dst.unlink(missing_ok=True)
        return {
            "created": False,
            "mode": "linked-review-copy",
            "input": str(src),
            "output": None,
            "message": "The linked review copy failed the static-copy preservation audit and was removed.",
            "flattening": flatten_report,
            "linking": link_result.to_dict(),
            "preservation": verification,
        }

    return {
        "created": True,
        "mode": "linked-review-copy",
        "input": str(src),
        "output": str(dst),
        "citation_manager_master_preserved": True,
        "flattening": flatten_report,
        "links_added": link_result.links_added,
        "references_bookmarked": link_result.references_bookmarked,
        "preservation": verification,
        "warnings": [
            "The original manuscript was not modified.",
            "This output is a static review/navigation copy. Do not use it as the master file for refreshing citation-manager fields.",
        ] + list(link_result.warnings),
        "message": (
            "Created a separate linked review copy. The original live citation-manager manuscript remains unchanged."
        ),
    }
