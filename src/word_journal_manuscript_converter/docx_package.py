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
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W_NS, "m": M_NS, "r": R_NS, "mc": MC_NS}

_XMLNS_RE = re.compile(rb"\bxmlns(?::([A-Za-z_][\w.-]*))?=[\"']([^\"']+)[\"']")
_MC_PREFIX_VALUE_ATTRS = {
    f"{{{MC_NS}}}Ignorable",
    f"{{{MC_NS}}}PreserveAttributes",
    f"{{{MC_NS}}}PreserveElements",
    f"{{{MC_NS}}}ProcessContent",
}


class DocxError(ValueError):
    pass


@dataclass
class PartData:
    name: str
    root: ET.Element


def _root_start_tag(data: bytes) -> bytes:
    """Return the serialized root start tag without relying on an XML parser."""
    offset = data.find(b"?>")
    start = data.find(b"<", offset + 2 if offset >= 0 else 0)
    while start >= 0 and data[start:start + 4] in {b"<!--", b"<?xm"}:
        start = data.find(b"<", start + 1)
    if start < 0:
        return b""
    quote: int | None = None
    for index in range(start + 1, len(data)):
        char = data[index]
        if quote is not None:
            if char == quote:
                quote = None
            continue
        if char in (34, 39):
            quote = char
        elif char == 62:
            return data[start:index + 1]
    return b""


def namespace_declarations(data: bytes) -> dict[str, str]:
    declarations: dict[str, str] = {}
    for match in _XMLNS_RE.finditer(_root_start_tag(data)):
        prefix = (match.group(1) or b"").decode("utf-8", "replace")
        uri = match.group(2).decode("utf-8", "replace")
        declarations[prefix] = uri
    return declarations


def _register_namespaces(declarations: dict[str, str]) -> None:
    for prefix, uri in declarations.items():
        if prefix == "xml":
            continue
        try:
            ET.register_namespace(prefix, uri)
        except (ValueError, TypeError):
            continue


def _used_namespace_uris(root: ET.Element) -> set[str]:
    uris: set[str] = set()
    for node in root.iter():
        if isinstance(node.tag, str) and node.tag.startswith("{"):
            uris.add(node.tag[1:].split("}", 1)[0])
        for name in node.attrib:
            if name.startswith("{"):
                uris.add(name[1:].split("}", 1)[0])
    return uris


def _compatibility_prefixes(root: ET.Element) -> set[str]:
    prefixes: set[str] = set()
    for attr_name in _MC_PREFIX_VALUE_ATTRS:
        raw = root.attrib.get(attr_name, "")
        for token in raw.split():
            prefix = token.split(":", 1)[0]
            if prefix:
                prefixes.add(prefix)
    return prefixes


def missing_compatibility_prefixes(data: bytes) -> list[str]:
    """Return mc:* prefixes referenced by the root but not declared there."""
    declarations = namespace_declarations(data)
    _register_namespaces(declarations)
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []
    return sorted(prefix for prefix in _compatibility_prefixes(root) if prefix not in declarations)


def parse_xml_preserving_namespaces(data: bytes) -> ET.Element:
    """Parse Word XML while protecting namespace prefixes used by mc:Ignorable.

    ElementTree normally discards xmlns declarations that are not otherwise
    used by an element or attribute QName. Microsoft Word commonly keeps newer
    Office namespace prefixes only in mc:Ignorable. If those declarations are
    lost during a rewrite, the XML remains parseable but Word can repair the
    document as unreadable content. This parser registers original prefixes and
    pins compatibility-only declarations on the root so ordinary ET.tostring()
    calls elsewhere in the package remain safe.
    """
    declarations = namespace_declarations(data)
    _register_namespaces(declarations)
    root = ET.fromstring(data)
    used_uris = _used_namespace_uris(root)
    for prefix in _compatibility_prefixes(root):
        uri = declarations.get(prefix)
        if not uri:
            continue
        if uri not in used_uris:
            root.set(f"xmlns:{prefix}", uri)
    return root


def serialize_xml_preserving_namespaces(root: ET.Element, original_data: bytes) -> bytes:
    """Serialize an edited XML root while retaining Word compatibility prefixes."""
    declarations = namespace_declarations(original_data)
    _register_namespaces(declarations)
    used_uris = _used_namespace_uris(root)
    for prefix in _compatibility_prefixes(root):
        uri = declarations.get(prefix)
        if not uri:
            continue
        if uri not in used_uris and f"xmlns:{prefix}" not in root.attrib:
            root.set(f"xmlns:{prefix}", uri)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


class DocxPackage:
    """Read-only OOXML package view with compatibility namespace protection."""

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
            for name in sorted(n for n in names if n.startswith("word/") and n.endswith(".xml")):
                data = zf.read(name)
                missing = missing_compatibility_prefixes(data)
                if missing:
                    raise DocxError(
                        f"{name} references undeclared markup-compatibility namespace prefix(es): "
                        + ", ".join(missing)
                    )
                try:
                    parse_xml_preserving_namespaces(data)
                except ET.ParseError as exc:
                    raise DocxError(f"Malformed XML in {name}: {exc}") from exc

    def names(self) -> list[str]:
        with zipfile.ZipFile(self.path) as zf:
            return zf.namelist()

    def read(self, name: str) -> bytes:
        with zipfile.ZipFile(self.path) as zf:
            return zf.read(name)

    def xml(self, name: str) -> ET.Element:
        try:
            return parse_xml_preserving_namespaces(self.read(name))
        except KeyError as exc:
            raise DocxError(f"Missing package part: {name}") from exc
        except ET.ParseError as exc:
            raise DocxError(f"Malformed XML in {name}: {exc}") from exc

    def has(self, name: str) -> bool:
        return name in set(self.names())

    def word_xml_parts(self) -> list[str]:
        return [n for n in self.names() if n.startswith("word/") and n.endswith(".xml")]

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
            for node in root.findall(".//w:instrText", NS):
                if node.text:
                    instructions.append(" ".join(node.text.split()))
            for node in root.findall(".//w:fldSimple", NS):
                instr = node.attrib.get(f"{{{W_NS}}}instr")
                if instr:
                    instructions.append(" ".join(instr.split()))
        return instructions

    def visible_text(self) -> str:
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
