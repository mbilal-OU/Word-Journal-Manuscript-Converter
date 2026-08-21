from __future__ import annotations

import html
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .audit import inspect_docx
from .citations import build_citation_graph
from .linking import link_plain_citations
from .static_review import create_linked_review_copy


def _manager_names(citation_inventory: dict[str, Any]) -> list[str]:
    names: list[str] = []
    if citation_inventory.get("endnote_fields", 0):
        names.append("EndNote")
    if citation_inventory.get("zotero_fields", 0):
        names.append("Zotero")
    if citation_inventory.get("csl_fields", 0):
        names.append("CSL/Mendeley")
    return names


def _style_label(style: str) -> str:
    return {
        "numeric-brackets": "Numbered [1]",
        "numeric-parentheses": "Numbered (1)",
        "numeric-superscript": "Superscript numbered",
        "author-year": "Author-year",
        "live-fields": "Live citation-manager fields",
        "undetermined": "Undetermined",
    }.get(style, style.replace("-", " ").title())


def analyze_citation_navigation(path: str | Path) -> dict[str, Any]:
    """Build a citation-navigation report without modifying the document."""
    p = Path(path)
    inventory = inspect_docx(p)
    graph = build_citation_graph(p)
    citation_inventory = asdict(inventory.citation)
    managers = _manager_names(citation_inventory)
    live = bool(citation_inventory.get("total_candidate_fields", 0))

    if live:
        strategy = "live-safe-or-static-review"
        capability = (
            "Live citation-manager fields were detected. Keep the master manuscript live. "
            "For a clickable Word document, create a separate static linked review copy; "
            "the original EndNote/Zotero/Mendeley master is not altered."
        )
    elif graph.mode in {"numbered", "author-year"}:
        # Keep the established strategy identifier for API/backward compatibility.
        strategy = "clickable-docx-export"
        capability = (
            "A separate clickable DOCX copy can be created automatically. In-text citations "
            "jump to matched references, and matched references link back to the first in-text occurrence."
        )
    else:
        strategy = "analysis-only"
        capability = (
            "The citation system can be audited, but a sufficiently reliable plain-text "
            "citation-to-reference pattern was not detected, so the document is not modified."
        )

    warnings = list(graph.warnings)
    if live:
        warnings.append(
            "The original live citation-manager manuscript is never flattened automatically. "
            "Static linking is available only as an explicitly requested separate review copy."
        )

    return {
        "tool": "Word Journal Manuscript Converter",
        "workflow": "Make Manuscript Clickable",
        "manuscript": p.name,
        "citation_manager": ", ".join(managers) if managers else "None detected",
        "live_fields": live,
        "live_field_count": citation_inventory.get("total_candidate_fields", 0),
        "bibliography_field_count": citation_inventory.get("bibliography_fields", 0),
        "detected_citation_style": _style_label(graph.citation_style),
        "citation_style_key": graph.citation_style,
        "detection_confidence": graph.detection_confidence,
        "navigation_strategy": strategy,
        "navigation_direction": "bidirectional" if graph.mode in {"numbered", "author-year"} else "analysis-only",
        "capability": capability,
        "citation_graph": graph.to_dict(),
        "warnings": warnings,
        "privacy": "Analysis is local. Manuscript text is not uploaded by this operation.",
    }


def make_navigable_copy(
    input_path: str | Path,
    output_path: str | Path,
    *,
    static_review_copy: bool = False,
) -> dict[str, Any]:
    """Create a navigable copy while preserving the original manuscript."""
    analysis = analyze_citation_navigation(input_path)
    if analysis["live_fields"]:
        if static_review_copy:
            data = create_linked_review_copy(input_path, output_path)
            data["detected_citation_style"] = analysis["detected_citation_style"]
            data["detection_confidence"] = analysis["detection_confidence"]
            return data
        return {
            "created": False,
            "mode": "live-safe-word-navigation",
            "input": str(Path(input_path)),
            "output": None,
            "citation_manager": analysis["citation_manager"],
            "message": (
                "Live citation-manager fields were detected. Keep this file as the editable master, "
                "or explicitly create a separate static linked review copy."
            ),
        }

    graph = analysis["citation_graph"]
    if graph.get("mode") not in {"numbered", "author-year"}:
        return {
            "created": False,
            "mode": "analysis-only",
            "input": str(Path(input_path)),
            "output": None,
            "citation_manager": analysis["citation_manager"],
            "message": "No reliably linkable plain-text citation pattern was detected.",
        }

    result = link_plain_citations(input_path, output_path)
    data = result.to_dict()
    data["created"] = bool(result.passed and Path(output_path).exists())
    # Preserve the established public mode identifier; expose direction separately.
    data["mode"] = "clickable-docx-export"
    data["navigation_direction"] = "bidirectional"
    data["message"] = (
        "Created a separate clickable manuscript copy with forward citation links and "
        "reference backlinks where unambiguous matches were available."
        if data["created"]
        else "The clickable copy was not kept because a safety or structural check failed."
    )
    return data


def render_navigation_html(report: dict[str, Any]) -> str:
    graph = report.get("citation_graph", {})
    links = graph.get("links", [])
    citation_rows: list[str] = []
    reference_rows: list[str] = []

    for item in links:
        key = html.escape(str(item.get("reference_key", "")))
        citation = html.escape(str(item.get("citation", key)))
        matched = bool(item.get("matched"))
        ambiguous = bool(item.get("ambiguous"))
        status = "Matched" if matched else ("Ambiguous" if ambiguous else "Unresolved")
        citation_cell = f'<a href="#ref-{key}">{citation}</a>' if matched else citation
        citation_rows.append(f"<tr><td>{citation_cell}</td><td>{status}</td><td>{key}</td></tr>")
        if matched:
            ref_text = html.escape(str(item.get("reference_text") or ""))
            reference_rows.append(
                f'<div class="ref" id="ref-{key}"><a class="back" href="#top">↑</a><strong>{citation}</strong> {ref_text}</div>'
            )

    warnings = "".join(f"<li>{html.escape(str(w))}</li>" for w in report.get("warnings", []))
    live_note = ""
    if report.get("live_fields"):
        live_note = (
            '<p class="notice"><strong>Live citation manager detected.</strong> '
            "Keep the original document as the editable master. A separate static linked review copy "
            "can be created from the desktop app when you explicitly choose that option.</p>"
        )

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Citation Navigator report</title>
<style>
body{{margin:0;background:#f6f8fa;color:#17212b;font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif}}
main{{max-width:1050px;margin:auto;padding:32px 20px 60px}}section{{background:white;border:1px solid #dce3e8;border-radius:14px;padding:22px;margin-bottom:16px}}
h1{{margin:0 0 6px}}.muted{{color:#66788a}}.notice{{background:#eef6ff;border:1px solid #b9d8ff;border-radius:10px;padding:12px}}
.grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}}.metric{{border:1px solid #e2e8f0;border-radius:10px;padding:12px}}.metric b{{display:block;font-size:20px}}
table{{width:100%;border-collapse:collapse}}th,td{{padding:9px;border-bottom:1px solid #edf0f2;text-align:left}}a{{color:#155eef}}.ref{{padding:10px 0;border-bottom:1px solid #edf0f2;scroll-margin-top:12px}}.back{{float:right;text-decoration:none}}@media(max-width:760px){{.grid{{grid-template-columns:1fr 1fr}}}}
</style></head><body><main id="top">
<section><h1>Make Manuscript Clickable</h1><p><strong>{html.escape(str(report.get('manuscript', '')))}</strong></p><p class="muted">{html.escape(str(report.get('privacy', '')))}</p>{live_note}</section>
<section><div class="grid">
<div class="metric"><span>Citation style</span><b>{html.escape(str(report.get('detected_citation_style', '')))}</b></div>
<div class="metric"><span>Confidence</span><b>{int(report.get('detection_confidence', 0))}%</b></div>
<div class="metric"><span>References</span><b>{int(graph.get('reference_count', 0))}</b></div>
<div class="metric"><span>Unresolved</span><b>{len(graph.get('unmatched_citations', []))}</b></div>
<div class="metric"><span>Ambiguous</span><b>{len(graph.get('ambiguous_citations', []))}</b></div>
</div><p><strong>Mode:</strong> {html.escape(str(report.get('navigation_strategy', '')))}</p><p>{html.escape(str(report.get('capability', '')))}</p></section>
<section><h2>Citation map</h2><table><thead><tr><th>Citation</th><th>Status</th><th>Reference key</th></tr></thead><tbody>{''.join(citation_rows) or '<tr><td colspan="3">No linkable citation rows detected.</td></tr>'}</tbody></table></section>
<section><h2>References</h2>{''.join(reference_rows) or '<p>No matched reference rows available.</p>'}</section>
<section><h2>Notes</h2><ul>{warnings or '<li>No additional warnings.</li>'}</ul></section>
</main></body></html>"""


def write_navigation_html(report: dict[str, Any], output_path: str | Path) -> Path:
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render_navigation_html(report), encoding="utf-8")
    return out
