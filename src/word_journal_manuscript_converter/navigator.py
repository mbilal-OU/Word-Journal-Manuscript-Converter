from __future__ import annotations

import html
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .audit import inspect_docx
from .citations import build_citation_graph
from .linking import link_plain_numbered_citations


def _manager_names(citation_inventory: dict[str, Any]) -> list[str]:
    names: list[str] = []
    if citation_inventory.get("endnote_fields", 0):
        names.append("EndNote")
    if citation_inventory.get("zotero_fields", 0):
        names.append("Zotero")
    if citation_inventory.get("csl_fields", 0):
        names.append("CSL/Mendeley")
    return names


def analyze_citation_navigation(path: str | Path) -> dict[str, Any]:
    """Build a citation-navigation report without modifying the document.

    Live citation-manager fields are treated as protected content. For those
    documents the recommended navigation surface is the Word add-in, which
    selects existing paragraphs/ranges without rewriting field payloads.
    Plain numbered citations can additionally be exported as a new DOCX with
    internal bookmarks and hyperlinks.
    """
    p = Path(path)
    inventory = inspect_docx(p)
    graph = build_citation_graph(p)
    citation_inventory = asdict(inventory.citation)
    managers = _manager_names(citation_inventory)
    live = bool(citation_inventory.get("total_candidate_fields", 0))

    if live:
        strategy = "live-safe-word-navigation"
        capability = (
            "Use the Word add-in Citation Navigator to jump between visible citations "
            "and bibliography entries without modifying EndNote/Zotero/Mendeley fields."
        )
    elif graph.mode == "numbered":
        strategy = "clickable-docx-export"
        capability = (
            "A separate clickable DOCX copy can be created with bibliography bookmarks "
            "and internal hyperlinks for simple plain numbered citations."
        )
    else:
        strategy = "analysis-only"
        capability = (
            "The citation system can be audited, but this document does not currently "
            "match a safely linkable plain numbered pattern."
        )

    warnings = list(graph.warnings)
    if live:
        warnings.append(
            "Live citation-manager fields are protected. The desktop app will not wrap or rewrite those fields to force hyperlinks."
        )

    return {
        "tool": "Word Journal Manuscript Converter",
        "workflow": "Citation Navigator",
        "manuscript": p.name,
        "citation_manager": ", ".join(managers) if managers else "None detected",
        "live_fields": live,
        "live_field_count": citation_inventory.get("total_candidate_fields", 0),
        "navigation_strategy": strategy,
        "capability": capability,
        "citation_graph": graph.to_dict(),
        "warnings": warnings,
        "privacy": "Analysis is local. Manuscript text is not uploaded by this operation.",
    }


def make_navigable_copy(input_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    """Create a clickable copy only when doing so is safe.

    Live citation-manager documents intentionally return a safe refusal instead
    of altering citation fields. The Word add-in provides non-mutating live
    navigation for those documents.
    """
    analysis = analyze_citation_navigation(input_path)
    if analysis["live_fields"]:
        return {
            "created": False,
            "mode": "live-safe-word-navigation",
            "input": str(Path(input_path)),
            "output": None,
            "citation_manager": analysis["citation_manager"],
            "message": (
                "No DOCX was modified because live citation-manager fields were detected. "
                "Use the Word add-in Citation Navigator for safe in-document navigation."
            ),
        }

    graph = analysis["citation_graph"]
    if graph.get("mode") != "numbered":
        return {
            "created": False,
            "mode": "analysis-only",
            "input": str(Path(input_path)),
            "output": None,
            "citation_manager": analysis["citation_manager"],
            "message": "No safely linkable plain numbered citation pattern was detected.",
        }

    result = link_plain_numbered_citations(input_path, output_path)
    data = result.to_dict()
    data["created"] = bool(result.passed and Path(output_path).exists())
    data["mode"] = "clickable-docx-export"
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
        status = "Matched" if matched else "Unresolved"
        citation_cell = f'<a href="#ref-{key}">{citation}</a>' if matched else citation
        citation_rows.append(f"<tr><td>{citation_cell}</td><td>{status}</td><td>{key}</td></tr>")
        if matched:
            ref_text = html.escape(str(item.get("reference_text") or ""))
            reference_rows.append(
                f'<div class="ref" id="ref-{key}"><a class="back" href="#top">↑</a><strong>{key}</strong> {ref_text}</div>'
            )

    warnings = "".join(f"<li>{html.escape(str(w))}</li>" for w in report.get("warnings", []))
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Citation Navigator report</title>
<style>
body{{margin:0;background:#f6f8fa;color:#17212b;font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif}}
main{{max-width:1050px;margin:auto;padding:32px 20px 60px}}section{{background:white;border:1px solid #dce3e8;border-radius:14px;padding:22px;margin-bottom:16px}}
h1{{margin:0 0 6px}}.muted{{color:#66788a}}.grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}}.metric{{border:1px solid #e2e8f0;border-radius:10px;padding:12px}}.metric b{{display:block;font-size:22px}}
table{{width:100%;border-collapse:collapse}}th,td{{padding:9px;border-bottom:1px solid #edf0f2;text-align:left}}a{{color:#155eef}}.ref{{padding:10px 0;border-bottom:1px solid #edf0f2;scroll-margin-top:12px}}.back{{float:right;text-decoration:none}}@media(max-width:760px){{.grid{{grid-template-columns:1fr 1fr}}}}
</style></head><body><main id="top">
<section><h1>Citation Navigator</h1><p><strong>{html.escape(str(report.get('manuscript', '')))}</strong></p><p class="muted">{html.escape(str(report.get('privacy', '')))}</p></section>
<section><div class="grid">
<div class="metric"><span>Manager</span><b>{html.escape(str(report.get('citation_manager', '')))}</b></div>
<div class="metric"><span>Live fields</span><b>{int(report.get('live_field_count', 0))}</b></div>
<div class="metric"><span>References</span><b>{int(graph.get('reference_count', 0))}</b></div>
<div class="metric"><span>Unresolved</span><b>{len(graph.get('unmatched_citations', []))}</b></div>
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
