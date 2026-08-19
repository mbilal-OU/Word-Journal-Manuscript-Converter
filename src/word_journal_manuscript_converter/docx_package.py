from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "m": M_NS, "r": R_NS}


class DocxError(ValueError):
    pass


@dataclass
class PartData:
    name: str
    root: ET.Element


class DocxPackage:
    """Read-only OOXML package view.

    Word Journal Manuscript Converter's first rule is preservation. Inspection therefore starts by
    reading the ZIP package directly instead of round-tripping the document
    through a high-level writer that may discard unsupported Word features.
    """

    def __init__(self, path: str | Path):
        self.path = Path(path)
        if not self.path.exists():
            raise DocxError(f"File not found: {self.path}")
        if self.path.suffix.lower() != ".docx":
            raise DocxError("Word Journal Manuscript Converter currently accepts .docx files only")
        if not zipfile.is_zipfile(self.path):
            raise DocxError("File is not a valid ZIP-based DOCX package")
        with zipfile.ZipFile(self.path) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names or "[Content_Types].xml" not in names:
                raise DocxError("ZIP is missing required DOCX package parts")

    def names(self) -> list[str]:
        with zipfile.ZipFile(self.path) as zf:
            return zf.namelist()

    def read(self, name: str) -> bytes:
        with zipfile.ZipFile(self.path) as zf:
            return zf.read(name)

    def xml(self, name: str) -> ET.Element:
        try:
            return ET.fromstring(self.read(name))
        except KeyError as exc:
            raise DocxError(f"Missing package part: {name}") from exc
        except ET.ParseError as exc:
            raise DocxError(f"Malformed XML in {name}: {exc}") from exc

    def has(self, name: str) -> bool:
        return name in set(self.names())

    def word_xml_parts(self) -> list[str]:
        return [
            n for n in self.names()
            if n.startswith("word/") and n.endswith(".xml")
        ]

    def story_parts(self) -> list[str]:
        names = set(self.names())
        selected = []
        for n in sorted(names):
            if n == "word/document.xml":
                selected.append(n)
            elif re.fullmatch(r"word/header\d+\.xml", n):
                selected.append(n)
            elif re.fullmatch(r"word/footer\d+\.xml", n):
                selected.append(n)
            elif n in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}:
                selected.append(n)
        return selected

    def field_instructions(self) -> list[str]:
        instructions: list[str] = []
        for part in self.story_parts():
            root = self.xml(part)
            # Complex fields store code in one or more instrText nodes.
            for node in root.findall(".//w:instrText", NS):
                if node.text:
                    instructions.append(" ".join(node.text.split()))
            # fldSimple stores the instruction in an attribute.
            for node in root.findall(".//w:fldSimple", NS):
                instr = node.attrib.get(f"{{{W_NS}}}instr")
                if instr:
                    instructions.append(" ".join(instr.split()))
        return instructions

    def visible_text(self) -> str:
        """Extract text as a preservation fingerprint, not as a renderer.

        Deleted tracked-change text is intentionally excluded; inserted text
        is included because that is what users ordinarily see in the current
        document state. Field *results* are included because they are visible.
        """
        chunks: list[str] = []
        for part in self.story_parts():
            root = self.xml(part)
            deleted_text = set(id(n) for d in root.findall(".//w:del", NS) for n in d.iter())
            for node in root.iter():
                if id(node) in deleted_text:
                    continue
                if node.tag == f"{{{W_NS}}}t" and node.text:
                    chunks.append(node.text)
                elif node.tag == f"{{{W_NS}}}tab":
                    chunks.append("\t")
                elif node.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                    chunks.append("\n")
            chunks.append("\n")
        return "".join(chunks)

    def numeric_tokens(self) -> list[str]:
        return re.findall(r"(?<![A-Za-z])[-+]?\d+(?:[.,]\d+)*(?:[eE][-+]?\d+)?", self.visible_text())

    def media_hashes(self) -> dict[str, str]:
        hashes: dict[str, str] = {}
        with zipfile.ZipFile(self.path) as zf:
            for name in sorted(zf.namelist()):
                if name.startswith("word/media/") and not name.endswith("/"):
                    hashes[name] = hashlib.sha256(zf.read(name)).hexdigest()
        return hashes

    def custom_xml_hashes(self) -> dict[str, str]:
        hashes: dict[str, str] = {}
        with zipfile.ZipFile(self.path) as zf:
            for name in sorted(zf.namelist()):
                if name.startswith("customXml/") and not name.endswith("/"):
                    hashes[name] = hashlib.sha256(zf.read(name)).hexdigest()
        return hashes
    def relationship_hashes(self) -> dict[str, str]:
        hashes: dict[str, str] = {}
        with zipfile.ZipFile(self.path) as zf:
            for name in sorted(zf.namelist()):
                if name.endswith(".rels") and not name.endswith("/"):
                    hashes[name] = hashlib.sha256(zf.read(name)).hexdigest()
        return hashes

    def content_types_hash(self) -> str:
        return hashlib.sha256(self.read("[Content_Types].xml")).hexdigest()
