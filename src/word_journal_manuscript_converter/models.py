from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class CitationInventory:
    total_candidate_fields: int = 0
    endnote_fields: int = 0
    zotero_fields: int = 0
    csl_fields: int = 0
    bibliography_fields: int = 0
    internal_hyperlinks: int = 0
    bookmarks: int = 0
    ref_fields: int = 0


@dataclass
class DocxInventory:
    path: str
    valid_docx: bool
    paragraphs: int = 0
    tables: int = 0
    sections: int = 0
    images: int = 0
    equations: int = 0
    native_equations: int = 0
    embedded_equation_objects: int = 0
    equation_story_parts: dict[str, int] = field(default_factory=dict)
    embedded_objects: int = 0
    comments: int = 0
    footnotes: int = 0
    endnotes: int = 0
    tracked_insertions: int = 0
    tracked_deletions: int = 0
    fields: int = 0
    hyperlinks: int = 0
    custom_xml_parts: int = 0
    citation: CitationInventory = field(default_factory=CitationInventory)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PreservationReport:
    before: str
    after: str
    text_identical: bool
    numeric_tokens_identical: bool
    field_instructions_identical: bool
    citation_field_count_identical: bool
    media_hashes_identical: bool
    custom_xml_hashes_identical: bool
    relationship_hashes_identical: bool
    content_types_identical: bool
    tracked_changes_identical: bool
    comments_identical: bool
    notes_identical: bool
    equations_identical: bool
    tables_identical: bool
    bookmarks_not_lost: bool
    hyperlinks_not_lost: bool
    details: dict[str, Any] = field(default_factory=dict)

    @property
    def passed(self) -> bool:
        return all(
            [
                self.text_identical,
                self.numeric_tokens_identical,
                self.field_instructions_identical,
                self.citation_field_count_identical,
                self.media_hashes_identical,
                self.custom_xml_hashes_identical,
                self.relationship_hashes_identical,
                self.content_types_identical,
                self.tracked_changes_identical,
                self.comments_identical,
                self.notes_identical,
                self.equations_identical,
                self.tables_identical,
                self.bookmarks_not_lost,
                self.hyperlinks_not_lost,
            ]
        )

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["passed"] = self.passed
        return data
