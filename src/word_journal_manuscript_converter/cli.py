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
from .retarget import retarget_docx
from .structure import extract_structure


def _dump(data: object, output: str | None = None) -> None:
    text = json.dumps(data, indent=2, ensure_ascii=False)
    if output:
        Path(output).write_text(text + "\n", encoding="utf-8")
    else:
        print(text)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="word-journal-converter",
        description="Integrity-first Word manuscript retargeting with citation/reference preservation.",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    inspect_p = sub.add_parser("inspect", help="Inventory preservation-sensitive DOCX features")
    inspect_p.add_argument("docx")
    inspect_p.add_argument("--json-out")

    struct_p = sub.add_parser("structure", help="Extract manuscript structure without rewriting DOCX")
    struct_p.add_argument("docx")
    struct_p.add_argument("--json-out")

    cite_p = sub.add_parser("citations", help="Map in-text citations to bibliography entries")
    cite_p.add_argument("docx")
    cite_p.add_argument("--json-out")

    verify_p = sub.add_parser("verify", help="Compare source and transformed DOCX for preservation")
    verify_p.add_argument("before")
    verify_p.add_argument("after")
    verify_p.add_argument("--json-out")

    ready_p = sub.add_parser("readiness", help="Check a DOCX against a local journal profile")
    ready_p.add_argument("docx")
    ready_p.add_argument("--profile", required=True)
    ready_p.add_argument("--json-out")

    retarget_p = sub.add_parser("retarget", help="Apply safe profile formatting and fail closed if preservation changes")
    retarget_p.add_argument("docx")
    retarget_p.add_argument("--profile", required=True)
    retarget_p.add_argument("--output", required=True)
    retarget_p.add_argument("--report")

    link_p = sub.add_parser("link-citations", help="Add internal links for simple plain-text numbered citations like [12]")
    link_p.add_argument("docx")
    link_p.add_argument("--output", required=True)
    link_p.add_argument("--report")

    gui_p = sub.add_parser("gui", help="Launch the local desktop GUI")

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "inspect":
            _dump(inspect_docx(args.docx).to_dict(), args.json_out)
            return 0
        if args.command == "structure":
            _dump(extract_structure(args.docx).to_dict(), args.json_out)
            return 0
        if args.command == "citations":
            _dump(build_citation_graph(args.docx).to_dict(), args.json_out)
            return 0
        if args.command == "verify":
            report = verify_preservation(args.before, args.after)
            _dump(report.to_dict(), args.json_out)
            return 0 if report.passed else 2
        if args.command == "readiness":
            _dump(readiness_check(args.docx, args.profile), args.json_out)
            return 0
        if args.command == "retarget":
            report = retarget_docx(args.docx, args.output, args.profile)
            _dump(report.to_dict(), args.report)
            return 0 if report.passed else 2
        if args.command == "link-citations":
            report = link_plain_numbered_citations(args.docx, args.output)
            _dump(report.to_dict(), args.report)
            return 0 if report.passed else 2
        if args.command == "gui":
            from .gui import main as gui_main
            gui_main()
            return 0
    except (DocxError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
