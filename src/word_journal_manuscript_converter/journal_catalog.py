from __future__ import annotations

from typing import Any

CHECKED_ON = "2026-08-19"


def _profile(
    journal: str,
    article_type: str,
    source_url: str,
    requirements: dict[str, Any],
    notes: str,
    *,
    source_urls: list[str] | None = None,
) -> dict[str, Any]:
    data: dict[str, Any] = {
        "journal": journal,
        "article_type": article_type,
        "source_url": source_url,
        "checked_on": CHECKED_ON,
        "notes": notes,
        "requirements": requirements,
    }
    if source_urls:
        data["source_urls"] = source_urls
    return data


_PLOS_BASE = {
    "abstract_required": True,
    "citations_must_resolve": True,
    "line_spacing": 2.0,
    "line_numbering": {"count_by": 1, "restart": "continuous"},
}

_PLOS_RESEARCH_SECTIONS = ["Introduction", "Materials and Methods", "Results", "Discussion"]

BUILTIN_PROFILE_CATALOG: dict[str, dict[str, Any]] = {
    "plos-one-research-article": _profile(
        "PLOS ONE",
        "research-article",
        "https://journals.plos.org/plosone/s/submission-guidelines",
        {**_PLOS_BASE, "abstract_max_words": 300, "required_sections": _PLOS_RESEARCH_SECTIONS},
        "Source-dated PLOS ONE Research Article profile. Initial submissions can be flexible; this profile checks only explicit rules supported by the current engine.",
    ),
    "plos-biology-research-article": _profile(
        "PLOS Biology",
        "research-article",
        "https://journals.plos.org/plosbiology/s/submission-guidelines",
        {**_PLOS_BASE, "required_sections": _PLOS_RESEARCH_SECTIONS},
        "PLOS Biology permits format-free initial submission. The profile captures the supported manuscript organization and review-format targets from the official guidance.",
    ),
    "plos-genetics-research-article": _profile(
        "PLOS Genetics",
        "research-article",
        "https://journals.plos.org/plosgenetics/s/submission-guidelines",
        {**_PLOS_BASE, "abstract_max_words": 300, "required_sections": ["Author Summary", *_PLOS_RESEARCH_SECTIONS]},
        "Research Article profile including the required Author Summary, 300-word abstract ceiling, double spacing, and continuous line numbering.",
    ),
    "plos-computational-biology-research-article": _profile(
        "PLOS Computational Biology",
        "research-article",
        "https://journals.plos.org/ploscompbiol/s/submission-guidelines",
        {**_PLOS_BASE, "abstract_max_words": 300, "required_sections": ["Author Summary", *_PLOS_RESEARCH_SECTIONS]},
        "Research Article profile based on the official PLOS Computational Biology submission guidance.",
    ),
    "plos-pathogens-research-article": _profile(
        "PLOS Pathogens",
        "research-article",
        "https://journals.plos.org/plospathogens/s/submission-guidelines",
        {**_PLOS_BASE, "abstract_max_words": 300, "required_sections": ["Author Summary", *_PLOS_RESEARCH_SECTIONS]},
        "Research Article profile including Author Summary and the supported PLOS review-format requirements.",
    ),
    "plos-neglected-tropical-diseases-research-article": _profile(
        "PLOS Neglected Tropical Diseases",
        "research-article",
        "https://journals.plos.org/plosntds/s/submission-guidelines",
        {**_PLOS_BASE, "required_sections": ["Author Summary", *_PLOS_RESEARCH_SECTIONS]},
        "Research Article profile. Only requirements that can be safely checked by the current engine are encoded.",
    ),
    "plos-medicine-research-article": _profile(
        "PLOS Medicine",
        "research-article",
        "https://journals.plos.org/plosmedicine/s/submission-guidelines",
        {**_PLOS_BASE, "required_sections": ["Introduction", "Methods", "Results", "Discussion"]},
        "Research Article profile using the supported structure and review-format requirements in the official PLOS Medicine guidance.",
    ),
    "scientific-reports-article": _profile(
        "Scientific Reports",
        "article",
        "https://www.nature.com/srep/author-instructions/submission-guidelines",
        {
            "abstract_required": True,
            "abstract_max_words": 200,
            "keywords_max": 6,
            "required_sections": ["Author contributions", "Data availability", "Competing interests"],
            "citations_must_resolve": True,
        },
        "Source-dated Scientific Reports profile. The engine evaluates the supported subset of the journal's current manuscript requirements.",
    ),
    "nature-communications-article": _profile(
        "Nature Communications",
        "article",
        "https://www.nature.com/ncomms/submit/article",
        {
            "abstract_required": True,
            "abstract_max_words": 200,
            "required_sections": ["Introduction", "Results", "Discussion", "Methods", "Author contributions", "Competing interests"],
            "citations_must_resolve": True,
            "line_spacing": 2.0,
        },
        "Nature Communications allows flexible first submission, but provides a Word manuscript format for revisions. This profile checks supported Article requirements without claiming production-ready typesetting.",
        source_urls=["https://www.nature.com/ncomms/submit/how-to-submit", "https://www.nature.com/ncomms/submit/article"],
    ),
    "nature-microbiology-article": _profile(
        "Nature Microbiology",
        "article",
        "https://www.nature.com/nmicrobiol/content",
        {
            "abstract_required": True,
            "abstract_max_words": 150,
            "required_sections": ["Results", "Discussion", "Methods"],
            "citations_must_resolve": True,
        },
        "Nature Microbiology Article profile. The Introduction is intentionally not required as a heading because the journal specifies an unheaded introduction.",
        source_urls=["https://www.nature.com/nmicrobiol/submission-guidelines", "https://www.nature.com/nmicrobiol/content"],
    ),
    "nucleic-acids-research-research-article": _profile(
        "Nucleic Acids Research",
        "research-article",
        "https://academic.oup.com/nar/pages/author-guidelines",
        {
            "abstract_required": True,
            "abstract_max_words": 200,
            "required_sections": ["Introduction", "Materials and Methods", "Results", "Discussion", "Data Availability"],
            "citations_must_resolve": True,
        },
        "NAR profile using its 200-word text abstract, core research sections, and mandatory Data Availability statement.",
        source_urls=["https://academic.oup.com/nar/pages/author-guidelines", "https://academic.oup.com/nar/pages/data_deposition_and_standardization"],
    ),
    "bioinformatics-original-paper": _profile(
        "Bioinformatics",
        "original-paper",
        "https://academic.oup.com/bioinformatics/pages/author-guidelines",
        {"abstract_required": True, "abstract_recommended_max_words": 150, "citations_must_resolve": True},
        "Bioinformatics supports format-free initial submission. Its 150-word structured-abstract recommendation is reported as a warning target rather than a hard rejection rule.",
    ),
}


def _microbiology_society_profile(journal: str, key: str, *, microbial_genomics: bool = False) -> None:
    sections = ["Introduction", "Methods", "Results", "Discussion", "Conflicts of interest", "Author contributions"]
    if microbial_genomics:
        sections = ["Data Summary", "Impact Statement", *sections]
    BUILTIN_PROFILE_CATALOG[key] = _profile(
        journal,
        "research-article",
        "https://www.microbiologyresearch.org/prepare-an-article",
        {
            "abstract_required": True,
            "keywords_min": 3,
            "keywords_max": 6,
            "required_sections": sections,
            "citations_must_resolve": True,
            "line_numbering": {"count_by": 1, "restart": "continuous"},
        },
        "Microbiology Society journals use format-free initial submission. This profile checks the supported cross-journal submission checklist and article-structure requirements."
        + (" It also includes Microbial Genomics-specific Data Summary and Impact Statement requirements." if microbial_genomics else ""),
    )


_microbiology_society_profile("Microbial Genomics", "microbial-genomics-research-article", microbial_genomics=True)
_microbiology_society_profile("Microbiology", "microbiology-research-article")
_microbiology_society_profile("Journal of General Virology", "journal-general-virology-research-article")
_microbiology_society_profile("Journal of Medical Microbiology", "journal-medical-microbiology-research-article")
_microbiology_society_profile("International Journal of Systematic and Evolutionary Microbiology", "ijsem-research-article")


BUILTIN_PROFILE_CATALOG.update(
    {
        "mbio-research-article": _profile(
            "mBio",
            "research-article",
            "https://journals.asm.org/journal/mbio/article-types",
            {
                "abstract_required": True,
                "abstract_max_words": 250,
                "required_sections": ["Importance", "Introduction", "Results", "Discussion", "Materials and Methods"],
                "citations_must_resolve": True,
            },
            "ASM journals are format neutral at initial submission. This profile checks the supported mBio Research Article content requirements, including the Importance section.",
            source_urls=["https://journals.asm.org/journal/mbio/submit", "https://journals.asm.org/journal/mbio/article-types"],
        ),
        "aem-research-article": _profile(
            "Applied and Environmental Microbiology",
            "research-article",
            "https://journals.asm.org/journal/aem/article-types",
            {"abstract_required": True, "citations_must_resolve": True},
            "ASM journals are format neutral at initial submission. This conservative profile avoids enforcing unsupported or production-stage formatting rules.",
        ),
        "msphere-research-article": _profile(
            "mSphere",
            "research-article",
            "https://journals.asm.org/journal/msphere/article-types",
            {"abstract_required": True, "citations_must_resolve": True},
            "ASM journals are format neutral at initial submission. This conservative profile focuses on checks the current engine can evaluate safely.",
        ),
    }
)


def _frontiers_profile(journal: str, slug: str) -> dict[str, Any]:
    return _profile(
        journal,
        "original-research",
        f"https://www.frontiersin.org/journals/{slug}/for-authors/author-guidelines",
        {
            "abstract_required": True,
            "abstract_max_words": 350,
            "keywords_min": 5,
            "keywords_max": 8,
            "required_sections": ["Introduction", "Materials and Methods", "Results", "Discussion"],
            "citations_must_resolve": True,
            "line_spacing": 1.0,
            "line_numbering": {"count_by": 1, "restart": "continuous"},
        },
        "Frontiers Original Research profile using the shared author-guideline requirements supported by this engine: 5-8 keywords, single spacing, page/line-number review formatting, and the recommended research-section structure.",
        source_urls=[
            f"https://www.frontiersin.org/journals/{slug}/for-authors/author-guidelines",
            "https://www.frontiersin.org/for-authors/where-to-publish/article-types",
        ],
    )


for _key, _journal, _slug in [
    ("frontiers-microbiology-original-research", "Frontiers in Microbiology", "microbiology"),
    ("frontiers-genetics-original-research", "Frontiers in Genetics", "genetics"),
    ("frontiers-bioinformatics-original-research", "Frontiers in Bioinformatics", "bioinformatics"),
    ("frontiers-cellular-infection-microbiology-original-research", "Frontiers in Cellular and Infection Microbiology", "cellular-and-infection-microbiology"),
    ("frontiers-ecology-evolution-original-research", "Frontiers in Ecology and Evolution", "ecology-and-evolution"),
    ("frontiers-plant-science-original-research", "Frontiers in Plant Science", "plant-science"),
    ("frontiers-molecular-biosciences-original-research", "Frontiers in Molecular Biosciences", "molecular-biosciences"),
    ("frontiers-immunology-original-research", "Frontiers in Immunology", "immunology"),
    ("frontiers-medicine-original-research", "Frontiers in Medicine", "medicine"),
    ("frontiers-veterinary-science-original-research", "Frontiers in Veterinary Science", "veterinary-science"),
]:
    BUILTIN_PROFILE_CATALOG[_key] = _frontiers_profile(_journal, _slug)
