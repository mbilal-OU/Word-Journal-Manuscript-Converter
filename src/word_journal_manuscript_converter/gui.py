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
from .linking import link_plain_numbered_citations
from .profiles import list_bundled_profiles
from .reporting import analyze_manuscript, format_text_report, write_html_report
from .retarget import retarget_docx


class WordJournalManuscriptConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Word Journal Manuscript Converter {__version__}")
        self.geometry("1060x760")
        self.minsize(900, 650)

        self.docx = tk.StringVar()
        self.profile = tk.StringVar()
        self.status = tk.StringVar(value="Local-only mode. Manuscript content is not uploaded.")
        self.last_report: dict | None = None
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
        style.configure("Primary.TButton", padding=(12, 8))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text="Word Journal Manuscript Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Integrity-first DOCX analysis, citation/reference checking, and safe journal retargeting",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 10))
        ttk.Label(
            outer,
            text="LOCAL PROCESSING  •  original file is never overwritten  •  unsafe transforms fail closed",
            style="Privacy.TLabel",
        ).pack(anchor="w", pady=(0, 14))

        source = ttk.LabelFrame(outer, text="Manuscript and journal", padding=12)
        source.pack(fill="x", pady=(0, 10))

        file_row = ttk.Frame(source)
        file_row.pack(fill="x", pady=4)
        ttk.Label(file_row, text="Manuscript", width=14).pack(side="left")
        ttk.Entry(file_row, textvariable=self.docx).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(file_row, text="Browse…", command=self.pick_docx).pack(side="left")

        profile_row = ttk.Frame(source)
        profile_row.pack(fill="x", pady=4)
        ttk.Label(profile_row, text="Journal profile", width=14).pack(side="left")
        self.profile_combo = ttk.Combobox(profile_row, textvariable=self.profile, state="readonly")
        self.profile_combo.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(profile_row, text="Custom JSON…", command=self.pick_profile).pack(side="left")

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(6, 10))
        ttk.Button(actions, text="Full analysis", style="Primary.TButton", command=self.full_analysis).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Inspect", command=self.inspect).pack(side="left", padx=6)
        ttk.Button(actions, text="Citation map", command=self.citations).pack(side="left", padx=6)
        ttk.Button(actions, text="Readiness", command=self.readiness).pack(side="left", padx=6)
        ttk.Button(actions, text="Safe retarget…", command=self.retarget).pack(side="left", padx=6)
        ttk.Button(actions, text="Link [N] citations…", command=self.link_citations).pack(side="left", padx=6)

        tools = ttk.Frame(outer)
        tools.pack(fill="x", pady=(0, 8))
        self.save_report_btn = ttk.Button(tools, text="Save HTML report…", command=self.save_html_report, state="disabled")
        self.save_report_btn.pack(side="left")
        ttk.Button(tools, text="Project website", command=lambda: webbrowser.open("https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/")).pack(side="right")

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
            messagebox.showerror("Word Journal Manuscript Converter", "Choose an existing .docx manuscript first.")
            return None
        return p

    def _show(self, data: dict, summary: str | None = None, report: bool = False) -> None:
        self.output.delete("1.0", "end")
        self.output.insert("1.0", json.dumps(data, indent=2, ensure_ascii=False))
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", summary or self._compact_summary(data))
        if report:
            self.last_report = data
            self.save_report_btn.configure(state="normal")
        self.status.set("Completed locally. No manuscript content was uploaded.")

    @staticmethod
    def _compact_summary(data: dict) -> str:
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

    def _run(self, fn, *, report: bool = False, summary_fn=None) -> None:
        try:
            self.status.set("Working locally…")
            self.update_idletasks()
            data = fn()
            if data is not None:
                summary = summary_fn(data) if summary_fn else None
                self._show(data, summary=summary, report=report)
        except Exception as exc:
            messagebox.showerror("Word Journal Manuscript Converter", str(exc))
            self.status.set("Operation stopped safely.")

    def full_analysis(self) -> None:
        p = self._require_docx()
        if not p:
            return
        profile = self._profile_ref()
        self._run(lambda: analyze_manuscript(p, profile), report=True, summary_fn=format_text_report)

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

    def link_citations(self) -> None:
        p = self._require_docx()
        if not p:
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=f"{p.stem}_linked.docx",
            filetypes=[("Word document", "*.docx")],
        )
        if out:
            self._run(lambda: link_plain_numbered_citations(p, out).to_dict())

    def save_html_report(self) -> None:
        if not self.last_report:
            return
        initial = "manuscript_report.html"
        if self.docx.get():
            initial = f"{Path(self.docx.get()).stem}_journal_report.html"
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


def main() -> None:
    try:
        app = WordJournalManuscriptConverterApp()
    except tk.TclError as exc:
        raise SystemExit(f"Could not start the GUI: {exc}")
    app.mainloop()


if __name__ == "__main__":
    main()
