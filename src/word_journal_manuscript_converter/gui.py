from __future__ import annotations

import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .audit import inspect_docx
from .citations import build_citation_graph
from .journal import readiness_check
from .linking import link_plain_numbered_citations
from .retarget import retarget_docx


class WordJournalManuscriptConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word Journal Manuscript Converter")
        self.geometry("980x720")
        self.minsize(820, 620)

        self.docx = tk.StringVar()
        self.profile = tk.StringVar()
        self.status = tk.StringVar(value="Local-only mode. Manuscript content is not uploaded.")

        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text="Word Journal Manuscript Converter", font=("TkDefaultFont", 20, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Word-native manuscript retargeting with citation integrity and preservation auditing",
        ).pack(anchor="w", pady=(2, 12))

        privacy = ttk.Label(
            outer,
            text="LOCAL PROCESSING  •  original file is never overwritten  •  post-transform integrity audit",
        )
        privacy.pack(anchor="w", pady=(0, 12))

        file_row = ttk.Frame(outer)
        file_row.pack(fill="x", pady=4)
        ttk.Label(file_row, text="Manuscript", width=13).pack(side="left")
        ttk.Entry(file_row, textvariable=self.docx).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(file_row, text="Browse…", command=self.pick_docx).pack(side="left")

        profile_row = ttk.Frame(outer)
        profile_row.pack(fill="x", pady=4)
        ttk.Label(profile_row, text="Journal profile", width=13).pack(side="left")
        ttk.Entry(profile_row, textvariable=self.profile).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(profile_row, text="Browse…", command=self.pick_profile).pack(side="left")

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(14, 10))
        ttk.Button(actions, text="Inspect", command=self.inspect).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Citation map", command=self.citations).pack(side="left", padx=6)
        ttk.Button(actions, text="Readiness", command=self.readiness).pack(side="left", padx=6)
        ttk.Button(actions, text="Safe retarget…", command=self.retarget).pack(side="left", padx=6)
        ttk.Button(actions, text="Link [N] citations…", command=self.link_citations).pack(side="left", padx=6)

        self.output = tk.Text(outer, wrap="word", font=("TkFixedFont", 10))
        self.output.pack(fill="both", expand=True, pady=(6, 8))

        footer = ttk.Label(outer, textvariable=self.status)
        footer.pack(anchor="w")

    def pick_docx(self):
        path = filedialog.askopenfilename(filetypes=[("Word document", "*.docx")])
        if path:
            self.docx.set(path)

    def pick_profile(self):
        path = filedialog.askopenfilename(filetypes=[("Journal profile", "*.json")])
        if path:
            self.profile.set(path)

    def _require_docx(self) -> Path | None:
        p = Path(self.docx.get())
        if not p.exists():
            messagebox.showerror("Word Journal Manuscript Converter", "Choose an existing .docx manuscript first.")
            return None
        return p

    def _show(self, data):
        self.output.delete("1.0", "end")
        self.output.insert("1.0", json.dumps(data, indent=2, ensure_ascii=False))
        self.status.set("Completed locally. No manuscript content was uploaded.")

    def _run(self, fn):
        try:
            data = fn()
            if data is not None:
                self._show(data)
        except Exception as exc:
            messagebox.showerror("Word Journal Manuscript Converter", str(exc))
            self.status.set("Operation stopped safely.")

    def inspect(self):
        p = self._require_docx()
        if p:
            self._run(lambda: inspect_docx(p).to_dict())

    def citations(self):
        p = self._require_docx()
        if p:
            self._run(lambda: build_citation_graph(p).to_dict())

    def readiness(self):
        p = self._require_docx()
        profile = Path(self.profile.get())
        if p and profile.exists():
            self._run(lambda: readiness_check(p, profile))
        elif p:
            messagebox.showerror("Word Journal Manuscript Converter", "Choose a journal profile JSON file.")

    def retarget(self):
        p = self._require_docx()
        profile = Path(self.profile.get())
        if not p or not profile.exists():
            if p:
                messagebox.showerror("Word Journal Manuscript Converter", "Choose a journal profile JSON file.")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".docx",
            initialfile=f"{p.stem}_retargeted.docx",
            filetypes=[("Word document", "*.docx")],
        )
        if out:
            self._run(lambda: retarget_docx(p, out, profile).to_dict())

    def link_citations(self):
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


def main() -> None:
    try:
        app = WordJournalManuscriptConverterApp()
    except tk.TclError as exc:
        raise SystemExit(f"Could not start the GUI: {exc}")
    app.mainloop()


if __name__ == "__main__":
    main()
