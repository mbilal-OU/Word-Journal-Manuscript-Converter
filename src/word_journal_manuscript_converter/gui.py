from __future__ import annotations

import json
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .audit import inspect_docx
from .citations import build_citation_graph
from .journal import readiness_check
from .navigator import analyze_citation_navigation, make_navigable_copy, write_navigation_html
from .profiles import list_bundled_profiles
from .reporting import analyze_manuscript, format_text_report, write_html_report
from .retarget import retarget_docx

ADDIN_GUIDE_URL = "https://github.com/mbilal-OU/Word-Journal-Manuscript-Converter/tree/main/integrations/word-addin"


class WordJournalManuscriptConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Word Journal Manuscript Converter {__version__}")
        self.geometry("1120x820")
        self.minsize(940, 700)

        self.docx = tk.StringVar()
        self.profile = tk.StringVar()
        self.status = tk.StringVar(value="Local-only mode. Manuscript content is not uploaded.")
        self.last_report: dict | None = None
        self.last_navigation: dict | None = None
        self._profile_display_to_ref: dict[str, str] = {}

        self._build_styles()
        self._build_ui()
        self._load_profiles()

    def _build_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("TkDefaultFont", 22, "bold"))
        style.configure("Sub.TLabel", foreground="#556575")
        style.configure("Privacy.TLabel", foreground="#16794a", font=("TkDefaultFont", 10, "bold"))
        style.configure("ModeTitle.TLabel", font=("TkDefaultFont", 14, "bold"))
        style.configure("Primary.TButton", padding=(12, 8))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Word Journal Manuscript Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Journal conversion, citation/reference navigation, and manuscript integrity auditing in one local app",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 8))
        ttk.Label(
            outer,
            text="LOCAL PROCESSING  •  original file is never overwritten  •  live citation-manager fields are protected",
            style="Privacy.TLabel",
        ).pack(anchor="w", pady=(0, 12))

        source = ttk.LabelFrame(outer, text="Manuscript", padding=12)
        source.pack(fill="x", pady=(0, 12))
        row = ttk.Frame(source)
        row.pack(fill="x")
        ttk.Label(row, text="Word document", width=14).pack(side="left")
        ttk.Entry(row, textvariable=self.docx).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_docx).pack(side="left")

        workflow = ttk.Notebook(outer)
        workflow.pack(fill="x", pady=(0, 12))
        journal_tab = ttk.Frame(workflow, padding=14)
        navigator_tab = ttk.Frame(workflow, padding=14)
        audit_tab = ttk.Frame(workflow, padding=14)
        workflow.add(journal_tab, text="Journal Conversion")
        workflow.add(navigator_tab, text="Citation Navigator")
        workflow.add(audit_tab, text="Manuscript Audit")

        self._build_journal_tab(journal_tab)
        self._build_navigator_tab(navigator_tab)
        self._build_audit_tab(audit_tab)

        output_tools = ttk.Frame(outer)
        output_tools.pack(fill="x", pady=(0, 8))
        self.save_report_btn = ttk.Button(
            output_tools,
            text="Save manuscript HTML report…",
            command=self.save_html_report,
            state="disabled",
        )
        self.save_report_btn.pack(side="left")
        self.save_nav_btn = ttk.Button(
            output_tools,
            text="Save citation navigation HTML…",
            command=self.save_navigation_html,
            state="disabled",
        )
        self.save_nav_btn.pack(side="left", padx=6)
        ttk.Button(
            output_tools,
            text="Project website",
            command=lambda: webbrowser.open("https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/"),
        ).pack(side="right")

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)
        summary_frame = ttk.Frame(notebook, padding=4)
        json_frame = ttk.Frame(notebook, padding=4)
        notebook.add(summary_frame, text="Summary")
        notebook.add(json_frame, text="JSON")

        self.summary = tk.Text(summary_frame, wrap="word", font=("TkFixedFont", 10), relief="flat")
        self.summary.pack(fill="both", expand=True)
        self.output = tk.Text(json_frame, wrap="none", font=("TkFixedFont", 10), relief="flat")
        self.output.pack(fill="both", expand=True)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(8, 0))
        ttk.Label(footer, textvariable=self.status, style="Sub.TLabel").pack(side="left")
        ttk.Label(footer, text=f"v{__version__}", style="Sub.TLabel").pack(side="right")

    def _build_journal_tab(self, parent) -> None:
        ttk.Label(parent, text="Prepare for a target journal", style="ModeTitle.TLabel").pack(anchor="w")
        ttk.Label(
            parent,
            text="Check a source-dated journal profile, then retarget only supported formatting on a new copy.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 8))
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text="Journal profile", width=14).pack(side="left")
        self.profile_combo = ttk.Combobox(row, textvariable=self.profile, state="readonly")
        self.profile_combo.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Custom JSON…", command=self.pick_profile).pack(side="left")
        actions = ttk.Frame(parent)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(
            actions,
            text="Journal analysis",
            style="Primary.TButton",
            command=self.journal_analysis,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Readiness", command=self.readiness).pack(side="left", padx=6)
        ttk.Button(actions, text="Safe retarget…", command=self.retarget).pack(side="left", padx=6)

    def _build_navigator_tab(self, parent) -> None:
        ttk.Label(parent, text="Make citations traceable", style="ModeTitle.TLabel").pack(anchor="w")
        ttk.Label(
            parent,
            text=(
                "No journal is required. Keep EndNote/Zotero/Mendeley citations live for editing, "
                "or create a separate static linked review copy when you only need traceable references."
            ),
            style="Sub.TLabel",
            wraplength=930,
        ).pack(anchor="w", pady=(2, 8))
        actions = ttk.Frame(parent)
        actions.pack(fill="x")
        ttk.Button(
            actions,
            text="Analyze navigation",
            style="Primary.TButton",
            command=self.citation_navigator,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Citation map", command=self.citations).pack(side="left", padx=6)
        ttk.Button(actions, text="Create navigable copy…", command=self.create_clickable_copy).pack(side="left", padx=6)
        ttk.Button(actions, text="Word add-in guide", command=lambda: webbrowser.open(ADDIN_GUIDE_URL)).pack(side="left", padx=6)

    def _build_audit_tab(self, parent) -> None:
        ttk.Label(parent, text="Audit without converting", style="ModeTitle.TLabel").pack(anchor="w")
        ttk.Label(
            parent,
            text=(
                "Inspect structure, citations, figures, tables, fields, comments, and other "
                "preservation-sensitive DOCX features. No journal is required."
            ),
            style="Sub.TLabel",
            wraplength=930,
        ).pack(anchor="w", pady=(2, 8))
        actions = ttk.Frame(parent)
        actions.pack(fill="x")
        ttk.Button(
            actions,
            text="Full manuscript audit",
            style="Primary.TButton",
            command=self.audit_analysis,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Integrity inventory", command=self.inspect).pack(side="left", padx=6)

    def _load_profiles(self) -> None:
        values: list[str] = []
        for desc in list_bundled_profiles():
            label = f"{desc.journal} — {desc.article_type}"
            values.append(label)
            self._profile_display_to_ref[label] = desc.key
        self.profile_combo["values"] = values
        if values:
            preferred = next((v for v in values if v.startswith("Generic review-copy")), values[0])
            self.profile.set(preferred)

    def pick_docx(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word document", "*.docx")])
        if path:
            self.docx.set(path)

    def pick_profile(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Journal profile", "*.json")])
        if path:
            label = f"Custom — {Path(path).name}"
            self._profile_display_to_ref[label] = path
            values = list(self.profile_combo["values"])
            if label not in values:
                values.append(label)
                self.profile_combo["values"] = values
            self.profile.set(label)

    def _profile_ref(self) -> str | None:
        return self._profile_display_to_ref.get(self.profile.get())

    def _require_docx(self) -> Path | None:
        p = Path(self.docx.get())
        if not p.exists() or p.suffix.lower() != ".docx":
            messagebox.showerror(
                "Word Journal Manuscript Converter",
                "Choose an existing .docx manuscript first.",
            )
            return None
        return p

    def _show(self, data: dict, summary: str | None = None) -> None:
        self.output.delete("1.0", "end")
        self.output.insert("1.0", json.dumps(data, indent=2, ensure_ascii=False))
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", summary or self._compact_summary(data))
        self.status.set("Completed locally. No manuscript content was uploaded.")

    @staticmethod
    def _compact_summary(data: dict) -> str:
        if data.get("workflow") == "Citation Navigator":
            graph = data.get("citation_graph", {})
            manager = data.get("citation_manager", "None detected")
            live = int(data.get("live_field_count", 0))
            references = int(graph.get("reference_count", 0))
            matched = int(graph.get("matched_links", 0))
            unresolved = len(graph.get("unmatched_citations", []))
            if data.get("live_fields"):
                next_step = (
                    "Safe options:\n"
                    "  1. Keep the master live and navigate with the Word add-in.\n"
                    "  2. Create a separate static linked review copy."
                )
            else:
                next_step = "This document can be exported as a clickable navigable copy when its citation pattern is linkable."
            return (
                "CITATION NAVIGATOR\n\n"
                f"Citation manager: {manager}\n"
                f"Live citation fields: {live}\n"
                f"References detected: {references}\n"
                f"Matched citation keys: {matched}\n"
                f"Unresolved keys: {unresolved}\n\n"
                f"{next_step}\n\n"
                f"{data.get('capability', '')}"
            )
        if data.get("mode") == "linked-review-copy":
            return (
                "LINKED REVIEW COPY\n\n"
                f"Created: {'Yes' if data.get('created') else 'No'}\n"
                f"Links added: {data.get('links_added', 0)}\n"
                f"References bookmarked: {data.get('references_bookmarked', 0)}\n\n"
                f"{data.get('message', '')}\n\n"
                "Keep the original manuscript as the editable citation-manager master."
            )
        if "readiness_score" in data:
            lines = [f"{data.get('journal', 'Journal')} readiness: {data.get('readiness_score', 0)}/100", ""]
            for check in data.get("checks", []):
                lines.append(f"[{str(check.get('status', '')).upper():4}] {check.get('detail', '')}")
            return "\n".join(lines)
        if "matched_links" in data:
            return (
                f"Citation mode: {data.get('mode', 'unknown')}\n"
                f"In-text citations: {data.get('in_text_citation_count', 0)}\n"
                f"Matched links: {data.get('matched_links', 0)}\n"
                f"Unresolved: {len(data.get('unmatched_citations', []))}\n"
                f"Uncited references: {len(data.get('uncited_references', []))}"
            )
        if "paragraphs" in data and "citation" in data:
            return (
                f"Paragraphs: {data.get('paragraphs', 0)}\nTables: {data.get('tables', 0)}\n"
                f"Equations: {data.get('equations', 0)}\nEmbedded media: {data.get('images', 0)}\n"
                f"Citation-manager candidate fields: {data.get('citation', {}).get('total_candidate_fields', 0)}"
            )
        return json.dumps(data, indent=2, ensure_ascii=False)

    def _run(self, fn, *, summary_fn=None) -> dict | None:
        try:
            self.status.set("Working locally…")
            self.update_idletasks()
            data = fn()
            if data is not None:
                summary = summary_fn(data) if summary_fn else None
                self._show(data, summary=summary)
            return data
        except Exception as exc:
            messagebox.showerror("Word Journal Manuscript Converter", str(exc))
            self.status.set("Operation stopped safely.")
            return None

    def _choose_live_navigation_action(self, analysis: dict) -> str | None:
        dialog = tk.Toplevel(self)
        dialog.title("Live citation manager detected")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.grab_set()

        result: dict[str, str | None] = {"value": None}
        frame = ttk.Frame(dialog, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Live citation fields detected", style="ModeTitle.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text=(
                f"{analysis.get('citation_manager', 'Citation manager')} with "
                f"{analysis.get('live_field_count', 0)} live citation field(s) was detected.\n\n"
                "Choose how you want to work. Your original manuscript will not be changed."
            ),
            style="Sub.TLabel",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(6, 14))

        ttk.Label(
            frame,
            text=(
                "Live Navigation keeps the citation-manager fields editable and opens the Word add-in guide.\n\n"
                "Linked Review Copy creates a separate static DOCX for reading/review. "
                "The copy intentionally removes live citation-manager fields, then adds internal citation/reference navigation where supported."
            ),
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")

        def choose(value: str | None) -> None:
            result["value"] = value
            dialog.destroy()

        ttk.Button(buttons, text="Use Live Navigation", command=lambda: choose("live")).pack(side="left")
        ttk.Button(
            buttons,
            text="Create Linked Review Copy",
            style="Primary.TButton",
            command=lambda: choose("review"),
        ).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cancel", command=lambda: choose(None)).pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        dialog.update_idletasks()
        x = self.winfo_rootx() + max(20, (self.winfo_width() - dialog.winfo_reqwidth()) // 2)
        y = self.winfo_rooty() + max(20, (self.winfo_height() - dialog.winfo_reqheight()) // 2)
        dialog.geometry(f"+{x}+{y}")
        self.wait_window(dialog)
        return result["value"]

    def journal_analysis(self) -> None:
        p = self._require_docx()
        profile = self._profile_ref()
        if not p or not profile:
            if p:
                messagebox.showerror("Word Journal Manuscript Converter", "Choose a journal profile.")
            return
        data = self._run(lambda: analyze_manuscript(p, profile), summary_fn=format_text_report)
        if data:
            self.last_report = data
            self.save_report_btn.configure(state="normal")

    def audit_analysis(self) -> None:
        p = self._require_docx()
        if not p:
            return
        data = self._run(lambda: analyze_manuscript(p, None), summary_fn=format_text_report)
        if data:
            self.last_report = data
            self.save_report_btn.configure(state="normal")

    def citation_navigator(self) -> None:
        p = self._require_docx()
        if not p:
            return
        data = self._run(lambda: analyze_citation_navigation(p))
        if data:
            self.last_navigation = data
            self.save_nav_btn.configure(state="normal")

    def create_clickable_copy(self) -> None:
        p = self._require_docx()
        if not p:
            return

        analysis = self._run(lambda: analyze_citation_navigation(p))
        if not analysis:
            return
        self.last_navigation = analysis
        self.save_nav_btn.configure(state="normal")

        static_review = False
        initialfile = f"{p.stem}_navigable.docx"

        if analysis.get("live_fields"):
            choice = self._choose_live_navigation_action(analysis)
            if choice is None:
                self.status.set("No document was changed.")
                return
            if choice == "live":
                self.status.set("Opening the Word add-in guide. The manuscript remains unchanged.")
                webbrowser.open(ADDIN_GUIDE_URL)
                return
            static_review = True
            initialfile = f"{p.stem}_linked_review_copy.docx"

        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=initialfile,
            filetypes=[("Word document", "*.docx")],
        )
        if not out:
            return

        data = self._run(
            lambda: make_navigable_copy(
                p,
                out,
                static_review_copy=static_review,
            )
        )
        if data and data.get("created"):
            messagebox.showinfo(
                "Citation Navigator",
                (
                    "Navigable copy created successfully.\n\n"
                    "Your original manuscript was not modified."
                    + (
                        "\n\nThis is a static review copy. Keep the original file as the EndNote/Zotero/Mendeley master."
                        if static_review
                        else ""
                    )
                ),
            )

    def inspect(self) -> None:
        p = self._require_docx()
        if p:
            self._run(lambda: inspect_docx(p).to_dict())

    def citations(self) -> None:
        p = self._require_docx()
        if p:
            self._run(lambda: build_citation_graph(p).to_dict())

    def readiness(self) -> None:
        p = self._require_docx()
        profile = self._profile_ref()
        if p and profile:
            self._run(lambda: readiness_check(p, profile))
        elif p:
            messagebox.showerror("Word Journal Manuscript Converter", "Choose a journal profile.")

    def retarget(self) -> None:
        p = self._require_docx()
        profile = self._profile_ref()
        if not p or not profile:
            if p:
                messagebox.showerror("Word Journal Manuscript Converter", "Choose a journal profile.")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=f"{p.stem}_retargeted.docx",
            filetypes=[("Word document", "*.docx")],
        )
        if out:
            self._run(lambda: retarget_docx(p, out, profile).to_dict())

    def save_html_report(self) -> None:
        if not self.last_report:
            return
        initial = (
            f"{Path(self.docx.get()).stem}_manuscript_report.html"
            if self.docx.get()
            else "manuscript_report.html"
        )
        out = filedialog.asksaveasfilename(
            defaultextension=".html",
            initialfile=initial,
            filetypes=[("HTML report", "*.html")],
        )
        if out:
            write_html_report(self.last_report, out)
            self.status.set(f"Saved local report: {out}")
            if messagebox.askyesno("Word Journal Manuscript Converter", "Open the report in your browser?"):
                webbrowser.open(Path(out).resolve().as_uri())

    def save_navigation_html(self) -> None:
        if not self.last_navigation:
            return
        initial = (
            f"{Path(self.docx.get()).stem}_citation_navigation.html"
            if self.docx.get()
            else "citation_navigation.html"
        )
        out = filedialog.asksaveasfilename(
            defaultextension=".html",
            initialfile=initial,
            filetypes=[("HTML report", "*.html")],
        )
        if out:
            write_navigation_html(self.last_navigation, out)
            self.status.set(f"Saved local citation navigation report: {out}")
            if messagebox.askyesno("Citation Navigator", "Open the navigation report in your browser?"):
                webbrowser.open(Path(out).resolve().as_uri())


def main() -> None:
    try:
        app = WordJournalManuscriptConverterApp()
    except tk.TclError as exc:
        raise SystemExit(f"Could not start the GUI: {exc}")
    app.mainloop()


if __name__ == "__main__":
    main()
