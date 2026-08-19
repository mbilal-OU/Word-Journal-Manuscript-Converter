from __future__ import annotations

import html
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from . import __version__
from .audit import inspect_docx
from .citations import build_citation_graph
from .journal import readiness_check
from .structure import extract_structure


def analyze_manuscript(docx_path: str | Path, profile: str | Path | None = None) -> dict[str, Any]:
    path = Path(docx_path)
    inventory = inspect_docx(path).to_dict()
    structure_obj = extract_structure(path)
    citations = build_citation_graph(path).to_dict()
    report: dict[str, Any] = {
        "tool": "Word Journal Manuscript Converter",
        "version": __version__,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "manuscript": str(path.resolve()),
        "privacy": "Analysis performed locally by the core engine. No manuscript content is uploaded by this operation.",
        "inventory": inventory,
        "structure": {
            "word_count": structure_obj.word_count,
            "abstract_word_count": structure_obj.abstract_word_count,
            "keywords": structure_obj.keywords,
            "headings": structure_obj.headings,
            "reference_count": len(structure_obj.reference_paragraphs),
        },
        "citation_graph": citations,
    }
    if profile:
        report["readiness"] = readiness_check(path, profile)
    return report


def _status_counts(checks: list[dict[str, Any]]) -> dict[str, int]:
    counts = {"pass": 0, "warn": 0, "fail": 0, "info": 0}
    for check in checks:
        status = str(check.get("status", "info"))
        counts[status] = counts.get(status, 0) + 1
    return counts


def format_text_report(report: dict[str, Any]) -> str:
    inv = report.get("inventory", {})
    structure = report.get("structure", {})
    citations = report.get("citation_graph", {})
    lines = [
        "WORD JOURNAL MANUSCRIPT CONVERTER",
        f"Version: {report.get('version', '')}",
        f"File: {Path(str(report.get('manuscript', ''))).name}",
        "",
        "MANUSCRIPT",
        f"  Words: {structure.get('word_count', 0)}",
        f"  Abstract words: {structure.get('abstract_word_count', 0)}",
        f"  Keywords: {len(structure.get('keywords', []))}",
        f"  References detected: {structure.get('reference_count', 0)}",
        "",
        "DOCX INTEGRITY INVENTORY",
        f"  Embedded media: {inv.get('images', 0)}",
        f"  Tables: {inv.get('tables', 0)}",
        f"  Equations: {inv.get('equations', 0)}",
        f"  Comments: {inv.get('comments', 0)}",
        f"  Tracked insertions/deletions: {inv.get('tracked_insertions', 0)}/{inv.get('tracked_deletions', 0)}",
        f"  Citation-manager candidate fields: {inv.get('citation', {}).get('total_candidate_fields', 0)}",
        "",
        "CITATIONS",
        f"  Mode: {citations.get('mode', 'unknown')}",
        f"  In-text citations: {citations.get('in_text_citation_count', 0)}",
        f"  Matched citation links: {citations.get('matched_links', 0)}",
        f"  Unresolved citation keys: {len(citations.get('unmatched_citations', []))}",
        f"  Uncited references: {len(citations.get('uncited_references', []))}",
    ]
    readiness = report.get("readiness")
    if isinstance(readiness, dict):
        checks = readiness.get("checks", [])
        counts = _status_counts(checks)
        lines += [
            "",
            "JOURNAL READINESS",
            f"  Journal: {readiness.get('journal', '')}",
            f"  Article type: {readiness.get('article_type', '')}",
            f"  Score: {readiness.get('readiness_score', 0)}/100",
            f"  Checks: {counts.get('pass', 0)} pass, {counts.get('warn', 0)} warn, {counts.get('fail', 0)} fail",
        ]
        failing = [c for c in checks if c.get("status") in {"fail", "warn"}]
        if failing:
            lines.append("  Attention:")
            for c in failing:
                lines.append(f"    [{str(c.get('status')).upper()}] {c.get('detail', '')}")
        source = readiness.get("profile_source_url")
        if source:
            lines.append(f"  Official source: {source}")
        checked = readiness.get("profile_checked_on")
        if checked:
            lines.append(f"  Profile checked: {checked}")
    lines += ["", "Always verify current official journal instructions before submission."]
    return "\n".join(lines)


def render_html_report(report: dict[str, Any]) -> str:
    inv = report.get("inventory", {})
    structure = report.get("structure", {})
    citations = report.get("citation_graph", {})
    readiness = report.get("readiness") if isinstance(report.get("readiness"), dict) else None

    def esc(value: Any) -> str:
        return html.escape(str(value))

    readiness_html = ""
    if readiness:
        checks = readiness.get("checks", [])
        rows = []
        for check in checks:
            status = str(check.get("status", "info"))
            rows.append(
                f'<tr><td><span class="badge {esc(status)}">{esc(status.upper())}</span></td>'
                f'<td>{esc(check.get("check", ""))}</td><td>{esc(check.get("detail", ""))}</td></tr>'
            )
        source = readiness.get("profile_source_url")
        source_html = f'<a href="{esc(source)}">official source</a>' if source else "No source URL recorded"
        readiness_html = f"""
        <section><h2>Journal readiness</h2>
        <div class="score">{esc(readiness.get('readiness_score', 0))}<span>/100</span></div>
        <p><strong>{esc(readiness.get('journal', ''))}</strong> · {esc(readiness.get('article_type', ''))}</p>
        <p class="muted">Profile checked {esc(readiness.get('profile_checked_on', 'unknown'))} · {source_html}</p>
        <table><thead><tr><th>Status</th><th>Check</th><th>Detail</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
        </section>
        """

    raw_json = html.escape(json.dumps(report, indent=2, ensure_ascii=False))
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Word Journal Manuscript Converter report</title>
<style>
:root{{--ink:#17212b;--muted:#66788a;--line:#dce3e8;--paper:#fff;--wash:#f5f8fa;--ok:#16794a;--warn:#a26608;--fail:#b42318;--info:#1769e0}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--wash);color:var(--ink);font:15px/1.55 system-ui,-apple-system,Segoe UI,sans-serif}}main{{max-width:1050px;margin:auto;padding:38px 22px 64px}}header,section{{background:var(--paper);border:1px solid var(--line);border-radius:16px;padding:24px;margin-bottom:18px}}h1{{font-size:30px;margin:0 0 6px}}h2{{margin:0 0 14px;font-size:21px}}.muted{{color:var(--muted)}}.grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}.metric{{border:1px solid var(--line);border-radius:12px;padding:14px}}.metric b{{display:block;font-size:24px}}table{{width:100%;border-collapse:collapse}}th,td{{text-align:left;padding:10px;border-bottom:1px solid var(--line);vertical-align:top}}.badge{{font-size:11px;font-weight:800;border-radius:999px;padding:4px 8px}}.pass{{background:#e9f8ef;color:var(--ok)}}.warn{{background:#fff4dc;color:var(--warn)}}.fail{{background:#ffebe9;color:var(--fail)}}.info{{background:#eaf2ff;color:var(--info)}}.score{{font-size:42px;font-weight:800}}.score span{{font-size:18px;color:var(--muted)}}details{{margin-top:12px}}pre{{white-space:pre-wrap;overflow-wrap:anywhere;background:#0f1720;color:#d9e4ef;padding:16px;border-radius:12px}}@media(max-width:760px){{.grid{{grid-template-columns:1fr 1fr}}}}
</style></head><body><main>
<header><h1>Word Journal Manuscript Converter</h1><p class="muted">Integrity and readiness report · v{esc(report.get('version',''))} · {esc(report.get('generated_at_utc',''))}</p><p><strong>{esc(Path(str(report.get('manuscript',''))).name)}</strong></p><p class="muted">{esc(report.get('privacy',''))}</p></header>
<section><h2>Manuscript overview</h2><div class="grid">
<div class="metric"><span>Words</span><b>{esc(structure.get('word_count',0))}</b></div>
<div class="metric"><span>Abstract</span><b>{esc(structure.get('abstract_word_count',0))}</b></div>
<div class="metric"><span>References</span><b>{esc(structure.get('reference_count',0))}</b></div>
<div class="metric"><span>Keywords</span><b>{esc(len(structure.get('keywords',[])))}</b></div>
</div></section>
<section><h2>Integrity inventory</h2><div class="grid">
<div class="metric"><span>Media</span><b>{esc(inv.get('images',0))}</b></div>
<div class="metric"><span>Tables</span><b>{esc(inv.get('tables',0))}</b></div>
<div class="metric"><span>Equations</span><b>{esc(inv.get('equations',0))}</b></div>
<div class="metric"><span>Live citation fields</span><b>{esc(inv.get('citation',{}).get('total_candidate_fields',0))}</b></div>
</div></section>
<section><h2>Citation graph</h2><p>Mode: <strong>{esc(citations.get('mode','unknown'))}</strong></p><div class="grid">
<div class="metric"><span>In-text</span><b>{esc(citations.get('in_text_citation_count',0))}</b></div>
<div class="metric"><span>Matched</span><b>{esc(citations.get('matched_links',0))}</b></div>
<div class="metric"><span>Unresolved</span><b>{esc(len(citations.get('unmatched_citations',[])))}</b></div>
<div class="metric"><span>Uncited refs</span><b>{esc(len(citations.get('uncited_references',[])))}</b></div>
</div></section>
{readiness_html}
<section><h2>Audit record</h2><p class="muted">Machine-readable data used to generate this report.</p><details><summary>Show JSON</summary><pre>{raw_json}</pre></details></section>
</main></body></html>"""


def write_html_report(report: dict[str, Any], path: str | Path) -> Path:
    target = Path(path)
    target.write_text(render_html_report(report), encoding="utf-8")
    return target
