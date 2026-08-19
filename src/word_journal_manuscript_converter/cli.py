from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .audit import inspect_docx, verify_preservation
from .citations import build_citation_graph
from .docx_package import DocxError
from .journal import readiness_check
from .linking import link_plain_numbered_citations
from .navigator import analyze_citation_navigation, make_navigable_copy, write_navigation_html
from .profiles import list_bundled_profiles, load_profile_data, validate_profile_data
from .reporting import analyze_manuscript, write_html_report
from .retarget import retarget_docx
from .structure import extract_structure
from .template_mode import inspect_template, retarget_from_template


def _dump(data: object, output: str | None = None) -> None:
    text = json.dumps(data, indent=2, ensure_ascii=False)
    if output:
        Path(output).write_text(text + "\n", encoding="utf-8")
    else:
        print(text)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="word-journal-converter",
        description="Local-first Word manuscript auditing, citation navigation, journal retargeting, and template adaptation.",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    inspect_p = sub.add_parser("inspect", help="Inventory preservation-sensitive DOCX features")
    inspect_p.add_argument("docx"); inspect_p.add_argument("--json-out")
    struct_p = sub.add_parser("structure", help="Extract manuscript structure without rewriting DOCX")
    struct_p.add_argument("docx"); struct_p.add_argument("--json-out")
    cite_p = sub.add_parser("citations", help="Map in-text citations to bibliography entries")
    cite_p.add_argument("docx"); cite_p.add_argument("--json-out")

    nav_p = sub.add_parser("navigate", help="Analyze citation/reference traceability with no journal required")
    nav_p.add_argument("docx"); nav_p.add_argument("--json-out"); nav_p.add_argument("--html-out")
    nav_copy_p = sub.add_parser("make-navigable", help="Create a navigable DOCX copy")
    nav_copy_p.add_argument("docx"); nav_copy_p.add_argument("--output", required=True); nav_copy_p.add_argument("--report")
    nav_copy_p.add_argument("--static-review-copy", action="store_true", help="Explicitly create a separate static review copy from a live citation-manager manuscript.")

    analyze_p = sub.add_parser("analyze", help="Run combined manuscript analysis")
    analyze_p.add_argument("docx"); analyze_p.add_argument("--profile"); analyze_p.add_argument("--json-out"); analyze_p.add_argument("--html-out")
    verify_p = sub.add_parser("verify", help="Compare source and transformed DOCX for preservation")
    verify_p.add_argument("before"); verify_p.add_argument("after"); verify_p.add_argument("--json-out")
    ready_p = sub.add_parser("readiness", help="Check a DOCX against a bundled or custom journal profile")
    ready_p.add_argument("docx"); ready_p.add_argument("--profile", required=True); ready_p.add_argument("--json-out")
    retarget_p = sub.add_parser("retarget", help="Apply safe profile formatting")
    retarget_p.add_argument("docx"); retarget_p.add_argument("--profile", required=True); retarget_p.add_argument("--output", required=True); retarget_p.add_argument("--report")

    template_inspect_p = sub.add_parser("template-inspect", help="Inspect transferable formatting in a Word .docx/.dotx journal template")
    template_inspect_p.add_argument("template"); template_inspect_p.add_argument("--json-out")
    template_retarget_p = sub.add_parser("template-retarget", help="Apply safe formatting from a Word journal template to a new manuscript copy")
    template_retarget_p.add_argument("docx"); template_retarget_p.add_argument("--template", required=True); template_retarget_p.add_argument("--output", required=True); template_retarget_p.add_argument("--report")

    link_p = sub.add_parser("link-citations", help="Add internal links for simple plain-text numbered citations")
    link_p.add_argument("docx"); link_p.add_argument("--output", required=True); link_p.add_argument("--report")
    profiles_p = sub.add_parser("profiles", help="List bundled journal profiles"); profiles_p.add_argument("--json-out")
    validate_p = sub.add_parser("validate-profile", help="Validate a bundled or custom profile")
    validate_p.add_argument("profile"); validate_p.add_argument("--json-out")
    sub.add_parser("gui", help="Launch the local desktop GUI")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "inspect": _dump(inspect_docx(args.docx).to_dict(), args.json_out); return 0
        if args.command == "structure": _dump(extract_structure(args.docx).to_dict(), args.json_out); return 0
        if args.command == "citations": _dump(build_citation_graph(args.docx).to_dict(), args.json_out); return 0
        if args.command == "navigate":
            report = analyze_citation_navigation(args.docx)
            if args.html_out: write_navigation_html(report, args.html_out)
            _dump(report, args.json_out); return 0
        if args.command == "make-navigable":
            report = make_navigable_copy(args.docx, args.output, static_review_copy=args.static_review_copy)
            _dump(report, args.report); return 0 if report.get("created") else 2
        if args.command == "analyze":
            report = analyze_manuscript(args.docx, args.profile)
            if args.html_out: write_html_report(report, args.html_out)
            _dump(report, args.json_out); return 0
        if args.command == "verify":
            report = verify_preservation(args.before, args.after); _dump(report.to_dict(), args.json_out); return 0 if report.passed else 2
        if args.command == "readiness": _dump(readiness_check(args.docx, args.profile), args.json_out); return 0
        if args.command == "retarget":
            report = retarget_docx(args.docx, args.output, args.profile); _dump(report.to_dict(), args.report); return 0 if report.passed else 2
        if args.command == "template-inspect": _dump(inspect_template(args.template), args.json_out); return 0
        if args.command == "template-retarget":
            report = retarget_from_template(args.docx, args.output, args.template); _dump(report.to_dict(), args.report); return 0 if report.passed else 2
        if args.command == "link-citations":
            report = link_plain_numbered_citations(args.docx, args.output); _dump(report.to_dict(), args.report); return 0 if report.passed else 2
        if args.command == "profiles": _dump([p.to_dict() for p in list_bundled_profiles()], args.json_out); return 0
        if args.command == "validate-profile":
            data, resolved = load_profile_data(args.profile); issues = validate_profile_data(data)
            _dump({"profile": resolved, "valid": not issues, "issues": issues, "journal": data.get("journal")}, args.json_out); return 0 if not issues else 2
        if args.command == "gui":
            from .gui import main as gui_main
            gui_main(); return 0
    except (DocxError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr); return 1
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
