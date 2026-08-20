from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
import webbrowser
from importlib import resources
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .audit import inspect_docx
from .branding import DEVELOPER_NAME, DISPLAY_VERSION, PRIVACY_URL, PRODUCT_NAME, PROJECT_URL, WORD_ADDIN_GUIDE_URL
from .citations import build_citation_graph
from .journal import readiness_check
from .navigator import analyze_citation_navigation, make_navigable_copy, write_navigation_html
from .profiles import list_bundled_profiles
from .reporting import analyze_manuscript, format_text_report, write_html_report
from .retarget import retarget_docx
from .telemetry import SettingsStore, TelemetryClient
from .template_mode import inspect_template, retarget_from_template
from .updates import check_for_update, download_asset, launch_downloaded_update

NAVY = "#0b2f68"
BLUE = "#155eef"
GREEN = "#067647"
RED = "#b42318"
MUTED = "#667085"
WASH = "#f5f7fb"
PAPER = "#ffffff"
TAB_IDLE = "#e4e7ec"
ACTIVE_WASH = "#eaf2ff"
DONE_WASH = "#ecfdf3"


class WordJournalManuscriptConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = SettingsStore()
        self.telemetry = TelemetryClient(self.settings)
        self.docx = tk.StringVar()
        self.profile = tk.StringVar()
        self.template = tk.StringVar()
        self.status = tk.StringVar(value="Ready. Manuscript processing stays on this computer.")
        self.active_task = tk.StringVar(value="Ready")
        self.activity_prefix = tk.StringVar(value="STATUS")
        self.last_report: dict | None = None
        self.last_navigation: dict | None = None
        self._profile_display_to_ref: dict[str, str] = {}
        self._icon_image: tk.PhotoImage | None = None
        self._action_buttons: dict[str, ttk.Button] = {}
        self._action_base_styles: dict[str, str] = {}
        self._active_action_key: str | None = None
        self.title(f"{PRODUCT_NAME} - {DISPLAY_VERSION}")
        self.geometry("1200x900")
        self.minsize(1000, 760)
        self.configure(bg=WASH)
        self._set_icon()
        self._styles()
        self._build_ui()
        self._load_profiles()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(700, self._first_run_consent)
        self.after(1800, self._auto_update_check)
        self.after(60000, self._heartbeat)
        if self.telemetry.enabled:
            self.telemetry.track("app_start", properties={"channel": "desktop"})

    def _set_icon(self) -> None:
        try:
            root = resources.files("word_journal_manuscript_converter").joinpath("assets")
            try:
                self.iconbitmap(default=str(root.joinpath("app-icon.ico")))
            except (tk.TclError, OSError):
                pass
            try:
                self._icon_image = tk.PhotoImage(file=str(root.joinpath("app-icon.png")))
                self.iconphoto(True, self._icon_image)
            except (tk.TclError, OSError):
                pass
        except Exception:
            pass

    def _styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background=WASH)
        style.configure("TLabel", background=WASH, foreground="#17212b")
        style.configure("Muted.TLabel", background=WASH, foreground=MUTED)
        style.configure("Mode.TLabel", background=WASH, foreground=NAVY, font=("Segoe UI", 15, "bold"))
        style.configure("Section.TLabel", background=WASH, foreground="#344054", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8), background=BLUE, foreground="white")
        style.map("Primary.TButton", background=[("active", "#004eeb"), ("!disabled", BLUE)], foreground=[("!disabled", "white")])
        style.configure("Active.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8), background=NAVY, foreground="white")
        style.map("Active.TButton", background=[("!disabled", NAVY)], foreground=[("!disabled", "white")])
        style.configure("TButton", padding=(10, 7))
        style.configure("Main.TNotebook", background=WASH, borderwidth=0)
        style.configure("Main.TNotebook.Tab", padding=(18, 10), font=("Segoe UI", 10, "bold"), background=TAB_IDLE, foreground="#344054")
        style.map(
            "Main.TNotebook.Tab",
            background=[("selected", BLUE), ("active", "#dbe8ff"), ("!selected", TAB_IDLE)],
            foreground=[("selected", "white"), ("active", NAVY), ("!selected", "#344054")],
        )
        style.configure("Mode.TNotebook", background=WASH, borderwidth=0)
        style.configure("Mode.TNotebook.Tab", padding=(15, 8), font=("Segoe UI", 10, "bold"), background="#eef2f6", foreground="#475467")
        style.map(
            "Mode.TNotebook.Tab",
            background=[("selected", NAVY), ("active", "#dbe8ff"), ("!selected", "#eef2f6")],
            foreground=[("selected", "white"), ("active", NAVY), ("!selected", "#475467")],
        )
        style.configure("Result.TNotebook.Tab", padding=(12, 7), font=("Segoe UI", 9, "bold"), background=TAB_IDLE, foreground="#475467")
        style.map(
            "Result.TNotebook.Tab",
            background=[("selected", "#344054"), ("!selected", TAB_IDLE)],
            foreground=[("selected", "white"), ("!selected", "#475467")],
        )

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=NAVY, height=105)
        header.pack(fill="x")
        header.pack_propagate(False)
        left = tk.Frame(header, bg=NAVY)
        left.pack(side="left", fill="both", expand=True, padx=24, pady=17)
        tk.Label(left, text=PRODUCT_NAME, bg=NAVY, fg="white", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        tk.Label(left, text="Convert journals | Navigate citations | Audit manuscripts | Adapt templates", bg=NAVY, fg="#cfe0ff", font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 0))
        right = tk.Frame(header, bg=NAVY)
        right.pack(side="right", padx=24, pady=13)
        tk.Label(right, text=DISPLAY_VERSION, bg="#123f80", fg="#dbe8ff", padx=10, pady=4, font=("Segoe UI", 9, "bold")).pack(anchor="e")
        tk.Label(right, text=f"Developed by {DEVELOPER_NAME}", bg=NAVY, fg="white", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(9, 0))

        privacy = tk.Frame(self, bg="#edf9f3", height=34)
        privacy.pack(fill="x")
        privacy.pack_propagate(False)
        tk.Label(privacy, text="LOCAL PROCESSING | original is never overwritten | live citation fields are protected", bg="#edf9f3", fg=GREEN, font=("Segoe UI", 9, "bold")).pack(side="left", padx=22, pady=7)

        body = ttk.Frame(self, padding=(20, 15, 20, 12))
        body.pack(fill="both", expand=True)
        src = ttk.LabelFrame(body, text="Select manuscript", padding=11)
        src.pack(fill="x")
        row = ttk.Frame(src)
        row.pack(fill="x")
        ttk.Label(row, text="Word document", width=14).pack(side="left")
        ttk.Entry(row, textvariable=self.docx).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Browse...", command=self.pick_docx).pack(side="left")

        self.activity_frame = tk.Frame(body, bg=ACTIVE_WASH, highlightbackground="#bfd3ff", highlightthickness=1)
        self.activity_frame.pack(fill="x", pady=(10, 0))
        tk.Label(self.activity_frame, textvariable=self.activity_prefix, bg=ACTIVE_WASH, fg=NAVY, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 8), pady=7)
        self.activity_label = tk.Label(self.activity_frame, textvariable=self.active_task, bg=ACTIVE_WASH, fg="#17212b", font=("Segoe UI", 10, "bold"))
        self.activity_label.pack(side="left", pady=7)
        self.progress = ttk.Progressbar(self.activity_frame, mode="indeterminate", length=150)
        self.progress.pack(side="right", padx=12, pady=7)
        self.progress.pack_forget()

        self.main_tabs = ttk.Notebook(body, style="Main.TNotebook")
        self.main_tabs.pack(fill="x", pady=(12, 10))
        jt, nt, at = (ttk.Frame(self.main_tabs, padding=15) for _ in range(3))
        self.main_tabs.add(jt, text="Journal Conversion")
        self.main_tabs.add(nt, text="Citation Navigator")
        self.main_tabs.add(at, text="Manuscript Audit")
        self._journal_tab(jt)
        self._navigator_tab(nt)
        self._audit_tab(at)
        self.main_tabs.bind("<<NotebookTabChanged>>", self._main_tab_changed)

        tools = ttk.Frame(body)
        tools.pack(fill="x", pady=(0, 7))
        self.save_report_btn = ttk.Button(tools, text="Save manuscript report...", command=self.save_html_report, state="disabled")
        self.save_report_btn.pack(side="left")
        self.save_nav_btn = ttk.Button(tools, text="Save citation map...", command=self.save_navigation_html, state="disabled")
        self.save_nav_btn.pack(side="left", padx=6)
        ttk.Button(tools, text="Feedback", command=self.feedback_dialog).pack(side="right")
        ttk.Button(tools, text="Privacy & analytics", command=self.privacy_dialog).pack(side="right", padx=6)
        ttk.Button(tools, text="Check updates", command=lambda: self.check_updates(manual=True)).pack(side="right")
        ttk.Button(tools, text="Website", command=lambda: webbrowser.open(PROJECT_URL)).pack(side="right", padx=6)

        outtabs = ttk.Notebook(body, style="Result.TNotebook")
        outtabs.pack(fill="both", expand=True)
        sf, jf = ttk.Frame(outtabs, padding=4), ttk.Frame(outtabs, padding=4)
        outtabs.add(sf, text="Summary")
        outtabs.add(jf, text="Details / JSON")
        self.summary = tk.Text(sf, wrap="word", font=("Consolas", 10), relief="flat", bg=PAPER)
        self.summary.pack(fill="both", expand=True)
        self.output = tk.Text(jf, wrap="none", font=("Consolas", 9), relief="flat", bg=PAPER)
        self.output.pack(fill="both", expand=True)

        footer = ttk.Frame(body)
        footer.pack(fill="x", pady=(8, 0))
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel").pack(side="left")
        ttk.Label(footer, text=f"{DISPLAY_VERSION} | Developed by {DEVELOPER_NAME} | engine {__version__}", style="Muted.TLabel").pack(side="right")

    def _main_tab_changed(self, _event=None) -> None:
        try:
            text = self.main_tabs.tab(self.main_tabs.select(), "text")
        except tk.TclError:
            return
        self._set_activity("WORKFLOW", text, state="ready")

    def _set_activity(self, prefix: str, text: str, *, state: str) -> None:
        self.activity_prefix.set(prefix)
        self.active_task.set(text)
        if state == "running":
            bg, border, fg = ACTIVE_WASH, "#84adff", NAVY
            if not self.progress.winfo_ismapped():
                self.progress.pack(side="right", padx=12, pady=7)
            self.progress.start(12)
        elif state == "done":
            bg, border, fg = DONE_WASH, "#abefc6", GREEN
            self.progress.stop(); self.progress.pack_forget()
        elif state == "error":
            bg, border, fg = "#fef3f2", "#fecdca", RED
            self.progress.stop(); self.progress.pack_forget()
        else:
            bg, border, fg = ACTIVE_WASH, "#bfd3ff", NAVY
            self.progress.stop(); self.progress.pack_forget()
        self.activity_frame.configure(bg=bg, highlightbackground=border)
        for child in self.activity_frame.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=bg)
        self.activity_label.configure(fg=fg)
        self.update_idletasks()

    def _action_button(self, parent, *, text: str, command, key: str, primary: bool = False) -> ttk.Button:
        base_style = "Primary.TButton" if primary else "TButton"
        button = ttk.Button(parent, text=text, style=base_style, command=lambda: self._select_action(key, text, command))
        self._action_buttons[key] = button
        self._action_base_styles[key] = base_style
        return button

    def _select_action(self, key: str, label: str, command) -> None:
        for existing_key, button in self._action_buttons.items():
            button.configure(style=self._action_base_styles.get(existing_key, "TButton"))
        button = self._action_buttons.get(key)
        if button is not None:
            button.configure(style="Active.TButton")
        self._active_action_key = key
        self._set_activity("RUNNING", label, state="running")
        self.status.set(f"Running: {label}")
        try:
            command()
            if self.activity_prefix.get() == "RUNNING" and self.active_task.get() == label:
                self._set_activity("DONE", label, state="done")
        except Exception as exc:
            self._set_activity("STOPPED", label, state="error")
            messagebox.showerror(PRODUCT_NAME, str(exc))
        finally:
            if button is not None:
                button.configure(style=self._action_base_styles.get(key, "TButton"))

    def _journal_tab(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Prepare for a target journal", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(parent, text="Choose one conversion path. Only the selected path is shown.", style="Muted.TLabel").pack(anchor="w", pady=(2, 9))

        self.journal_modes = ttk.Notebook(parent, style="Mode.TNotebook")
        self.journal_modes.pack(fill="x")
        profile_page = ttk.Frame(self.journal_modes, padding=12)
        template_page = ttk.Frame(self.journal_modes, padding=12)
        self.journal_modes.add(profile_page, text="Journal Profile")
        self.journal_modes.add(template_page, text="Word Template")

        ttk.Label(profile_page, text="Use a journal-specific rules profile", style="Section.TLabel").pack(anchor="w")
        ttk.Label(profile_page, text="Analyze requirements first, then create a new converted copy with independent assurance checks.", style="Muted.TLabel").pack(anchor="w", pady=(2, 7))
        row = ttk.Frame(profile_page); row.pack(fill="x")
        ttk.Label(row, text="Journal", width=14).pack(side="left")
        self.profile_combo = ttk.Combobox(row, textvariable=self.profile, state="readonly", height=24)
        self.profile_combo.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Custom JSON...", command=self.pick_profile).pack(side="left")
        actions = ttk.Frame(profile_page); actions.pack(fill="x", pady=(8, 0))
        self._action_button(actions, text="Journal analysis", command=self.journal_analysis, key="journal_analysis", primary=True).pack(side="left")
        self._action_button(actions, text="Readiness", command=self.readiness, key="journal_readiness").pack(side="left", padx=7)
        self._action_button(actions, text="Convert to journal format...", command=self.retarget, key="journal_convert").pack(side="left")

        ttk.Label(template_page, text="Adapt to a supplied Word template", style="Section.TLabel").pack(anchor="w")
        ttk.Label(template_page, text="The template is a formatting source. Supported fidelity and total template coverage are reported separately.", style="Muted.TLabel").pack(anchor="w", pady=(2, 7))
        tr = ttk.Frame(template_page); tr.pack(fill="x")
        ttk.Label(tr, text="Template", width=14).pack(side="left")
        ttk.Entry(tr, textvariable=self.template).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(tr, text="Browse template...", command=self.pick_template).pack(side="left")
        ta = ttk.Frame(template_page); ta.pack(fill="x", pady=(8, 0))
        self._action_button(ta, text="Inspect template", command=self.inspect_journal_template, key="template_inspect").pack(side="left")
        self._action_button(ta, text="Adapt to template...", command=self.apply_journal_template, key="template_apply", primary=True).pack(side="left", padx=7)

    def _navigator_tab(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Make citations traceable", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(parent, text="No journal required. Live citation fields stay protected unless you explicitly create a separate static review copy.", style="Muted.TLabel").pack(anchor="w", pady=(2, 9))
        actions = ttk.Frame(parent); actions.pack(fill="x")
        self._action_button(actions, text="Analyze navigation", command=self.citation_navigator, key="nav_analyze", primary=True).pack(side="left")
        self._action_button(actions, text="Citation map", command=self.citations, key="nav_map").pack(side="left", padx=7)
        self._action_button(actions, text="Create navigable copy...", command=self.create_clickable_copy, key="nav_copy").pack(side="left")
        self._action_button(actions, text="Word add-in setup", command=self.open_word_addin_guide, key="addin_setup").pack(side="left", padx=7)

    def _audit_tab(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Audit without converting", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(parent, text="Inspect structure, citations, figures, tables, equations, fields, comments and tracked changes.", style="Muted.TLabel").pack(anchor="w", pady=(2, 9))
        actions = ttk.Frame(parent); actions.pack(fill="x")
        self._action_button(actions, text="Full manuscript audit", command=self.audit_analysis, key="audit_full", primary=True).pack(side="left")
        self._action_button(actions, text="Integrity inventory", command=self.inspect, key="audit_inventory").pack(side="left", padx=7)

    def _load_profiles(self) -> None:
        vals: list[str] = []
        for descriptor in sorted(list_bundled_profiles(), key=lambda x: (x.journal.lower(), x.article_type.lower())):
            label = f"{descriptor.journal} | {descriptor.article_type}"
            vals.append(label)
            self._profile_display_to_ref[label] = descriptor.key
        self.profile_combo["values"] = vals
        if vals:
            self.profile.set(next((x for x in vals if x.startswith("Generic review-copy")), vals[0]))

    def pick_docx(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word document", "*.docx")])
        if path:
            self.docx.set(path)

    def pick_profile(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Journal profile", "*.json")])
        if not path:
            return
        label = f"Custom | {Path(path).name}"
        self._profile_display_to_ref[label] = path
        values = list(self.profile_combo["values"])
        if label not in values:
            values.append(label)
            self.profile_combo["values"] = values
        self.profile.set(label)

    def pick_template(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word template/document", "*.docx *.dotx"), ("All files", "*.*")])
        if path:
            self.template.set(path)

    def _profile_ref(self) -> str | None:
        return self._profile_display_to_ref.get(self.profile.get())

    def _require_docx(self) -> Path | None:
        path = Path(self.docx.get())
        if not path.exists() or path.suffix.lower() != ".docx":
            messagebox.showerror(PRODUCT_NAME, "Choose an existing .docx manuscript first.")
            return None
        return path

    def _require_template(self) -> Path | None:
        path = Path(self.template.get())
        if not path.exists() or path.suffix.lower() not in {".docx", ".dotx"}:
            messagebox.showerror("Template Mode", "Choose an existing .docx or .dotx journal template first.")
            return None
        return path

    @staticmethod
    def _same_path(a: str | Path, b: str | Path) -> bool:
        try:
            return Path(a).resolve() == Path(b).resolve()
        except OSError:
            return False

    def _show(self, data: dict, summary: str | None = None) -> None:
        self.output.delete("1.0", "end")
        self.output.insert("1.0", json.dumps(data, indent=2, ensure_ascii=False))
        self.summary.delete("1.0", "end")
        self.summary.insert("1.0", summary or self._compact_summary(data))
        self.status.set("Completed locally. No manuscript content was uploaded.")

    @staticmethod
    def _compact_summary(data: dict) -> str:
        if "supported_fidelity_score" in data and "template_coverage_score" in data:
            level = str(data.get("adaptation_level", "unknown")).upper()
            return (
                "TEMPLATE ADAPTATION\n\n"
                f"Adaptation level: {level}\n"
                f"Supported fidelity: {data.get('supported_fidelity_score', 0)}%\n"
                f"Template coverage: {data.get('template_coverage_score', 0)}%\n"
                f"Verdict: {data.get('verdict', 'Not available')}\n"
                f"Output kept: {'Yes' if data.get('passed') else 'No'}\n\n"
                "Important: 100% supported fidelity does not mean 100% template conversion when coverage is lower."
            )
        if "formatting_compliance_score" in data and "assurance" in data:
            assurance = data.get("assurance", {})
            return (
                "JOURNAL CONVERSION ASSURANCE\n\n"
                f"Formatting compliance: {data.get('formatting_compliance_score', 0)}%\n"
                f"Manuscript requirements: {data.get('manuscript_requirement_score', 0)}/100\n"
                f"Document integrity: {assurance.get('document_integrity', 'Unknown')}\n"
                f"Structural sanity: {assurance.get('structural_sanity', 'Unknown')}\n"
                f"Blocking failures: {assurance.get('blocking_failures', 0)}\n"
                f"Verdict: {data.get('verdict', 'Not available')}"
            )
        if data.get("workflow") == "Citation Navigator":
            graph = data.get("citation_graph", {})
            return (
                "CITATION NAVIGATOR\n\n"
                f"Manager: {data.get('citation_manager', 'None detected')}\n"
                f"Live fields: {data.get('live_field_count', 0)}\n"
                f"References: {graph.get('reference_count', 0)}\n"
                f"Matched keys: {graph.get('matched_links', 0)}\n"
                f"Unresolved: {len(graph.get('unmatched_citations', []))}\n\n"
                f"{data.get('capability', '')}"
            )
        if data.get("mode") == "linked-review-copy":
            return (
                "LINKED REVIEW COPY\n\n"
                f"Created: {'Yes' if data.get('created') else 'No'}\n"
                f"Links added: {data.get('links_added', 0)}\n"
                f"References bookmarked: {data.get('references_bookmarked', 0)}\n\n"
                "Keep the original manuscript as the editable citation-manager master."
            )
        if data.get("workflow") == "Template Mode":
            return (
                "TEMPLATE INSPECTION\n\n"
                f"Type: {data.get('template_type', '')}\n"
                f"Transferable paragraph styles: {data.get('transferable_style_count', 0)}\n"
                f"Inferred style roles: {len(data.get('inferred_style_roles', {}))}\n"
                f"Page settings detected: {'Yes' if data.get('page_format') else 'No'}\n"
                "Exact visual match guaranteed: No"
            )
        if "readiness_score" in data:
            lines = [f"{data.get('journal', 'Journal')} readiness: {data.get('readiness_score', 0)}/100", ""]
            lines += [f"[{str(check.get('status', '')).upper():4}] {check.get('detail', '')}" for check in data.get("checks", [])]
            return "\n".join(lines)
        if "paragraphs" in data and "citation" in data:
            return (
                f"Paragraphs: {data.get('paragraphs', 0)}\n"
                f"Tables: {data.get('tables', 0)}\n"
                f"Equations total: {data.get('equations', 0)}\n"
                f"  Native Word OMML: {data.get('native_equations', 0)}\n"
                f"  Embedded equation objects: {data.get('embedded_equation_objects', 0)}\n"
                f"Embedded media: {data.get('images', 0)}\n"
                f"Citation-manager fields: {data.get('citation', {}).get('total_candidate_fields', 0)}"
            )
        return json.dumps(data, indent=2, ensure_ascii=False)

    def _run(self, fn, *, summary_fn=None, feature: str | None = None) -> dict | None:
        try:
            self.configure(cursor="watch")
            self.update_idletasks()
            data = fn()
            if data is not None:
                self._show(data, summary_fn(data) if summary_fn else None)
            if feature:
                self.telemetry.track_feature(feature, result="success")
            return data
        except Exception as exc:
            if feature:
                self.telemetry.track_feature(feature, result="stopped")
            self._set_activity("STOPPED", self.active_task.get(), state="error")
            messagebox.showerror(PRODUCT_NAME, str(exc))
            self.status.set("Operation stopped safely.")
            return None
        finally:
            try:
                self.configure(cursor="")
            except tk.TclError:
                pass

    def _first_run_consent(self) -> None:
        if self.settings.analytics_consent is not None:
            return
        answer = messagebox.askyesno(
            "Optional anonymous usage statistics",
            "Help improve Word Journal Manuscript Converter by sharing anonymous product usage statistics?\n\n"
            "Collected: feature names, app version, operating system, anonymous random install ID, session duration.\n\n"
            "Never collected: manuscript text, filenames, paths, citations, references, figures, document hashes, or scientific data.\n\n"
            "You can change this anytime under Privacy & analytics.",
        )
        self.settings.set_analytics_consent(answer)
        if answer:
            self.telemetry.track("app_start", properties={"channel": "desktop"})

    def privacy_dialog(self) -> None:
        dlg = tk.Toplevel(self); dlg.title("Privacy & analytics"); dlg.transient(self); dlg.grab_set()
        frame = ttk.Frame(dlg, padding=18); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Privacy & analytics", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Manuscript processing remains local. Analytics is optional and contains no manuscript content.", style="Muted.TLabel", wraplength=520).pack(anchor="w", pady=(5, 12))
        analytics = tk.BooleanVar(value=self.telemetry.enabled)
        auto_updates = tk.BooleanVar(value=bool(self.settings.get("auto_update_check", True)))
        ttk.Checkbutton(frame, text="Share anonymous usage statistics", variable=analytics).pack(anchor="w", pady=3)
        ttk.Checkbutton(frame, text="Automatically check GitHub for updates once per day", variable=auto_updates).pack(anchor="w", pady=3)
        buttons = ttk.Frame(frame); buttons.pack(fill="x", pady=(14, 0))
        def save() -> None:
            self.settings.set_analytics_consent(analytics.get())
            self.settings.set("auto_update_check", auto_updates.get())
            dlg.destroy()
        ttk.Button(buttons, text="Save", style="Primary.TButton", command=save).pack(side="left")
        ttk.Button(buttons, text="Privacy policy", command=lambda: webbrowser.open(PRIVACY_URL)).pack(side="left", padx=7)
        ttk.Button(buttons, text="Cancel", command=dlg.destroy).pack(side="right")

    def feedback_dialog(self) -> None:
        self.telemetry.track("feedback_open")
        dlg = tk.Toplevel(self); dlg.title("Send feedback"); dlg.transient(self); dlg.grab_set(); dlg.geometry("560x460")
        frame = ttk.Frame(dlg, padding=18); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Send feedback", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Feedback is sent only when you press Send. Do not paste unpublished manuscript content.", style="Muted.TLabel", wraplength=510).pack(anchor="w", pady=(4, 10))
        rating = tk.StringVar(value="5")
        category = tk.StringVar(value="usability")
        email = tk.StringVar()
        contact = tk.BooleanVar(value=False)
        ttk.Label(frame, text="Rating (1-5)").pack(anchor="w")
        ttk.Combobox(frame, textvariable=rating, values=["5", "4", "3", "2", "1"], state="readonly").pack(fill="x", pady=(2, 7))
        ttk.Label(frame, text="Category").pack(anchor="w")
        ttk.Combobox(frame, textvariable=category, values=["bug", "feature", "usability", "journal", "citation", "template", "other"], state="readonly").pack(fill="x", pady=(2, 7))
        ttk.Label(frame, text="Message").pack(anchor="w")
        text = tk.Text(frame, height=7, wrap="word"); text.pack(fill="both", expand=True, pady=(2, 7))
        ttk.Label(frame, text="Email (optional)").pack(anchor="w")
        ttk.Entry(frame, textvariable=email).pack(fill="x", pady=(2, 4))
        ttk.Checkbutton(frame, text="You may contact me about this feedback", variable=contact).pack(anchor="w")
        buttons = ttk.Frame(frame); buttons.pack(fill="x", pady=(10, 0))
        def submit() -> None:
            msg = text.get("1.0", "end").strip()
            if not msg:
                messagebox.showerror("Feedback", "Enter a feedback message.")
                return
            buttons.winfo_children()[0].configure(state="disabled")
            self.status.set("Sending feedback...")
            def worker() -> None:
                result = self.telemetry.submit_feedback(rating=int(rating.get()), category=category.get(), message=msg, contact_email=email.get(), consent_to_contact=contact.get())
                self.after(0, lambda: self._feedback_done(dlg, result.sent, result.error))
            threading.Thread(target=worker, daemon=True).start()
        ttk.Button(buttons, text="Send feedback", style="Primary.TButton", command=submit).pack(side="left")
        ttk.Button(buttons, text="Cancel", command=dlg.destroy).pack(side="right")

    def _feedback_done(self, dlg: tk.Toplevel, sent: bool, error: str | None) -> None:
        if sent:
            dlg.destroy(); self.status.set("Thank you. Feedback was sent."); messagebox.showinfo("Feedback", "Thank you. Your feedback was sent.")
        else:
            self.status.set("Feedback could not be sent."); messagebox.showerror("Feedback", f"Could not send feedback.\n\n{error or 'Network unavailable.'}")

    def _auto_update_check(self) -> None:
        if self.settings.get("auto_update_check", True):
            self.check_updates(manual=False)

    def check_updates(self, *, manual: bool) -> None:
        self.telemetry.track("update_check", properties={"action": "manual" if manual else "automatic"})
        if manual:
            self.status.set("Checking for updates...")
        def worker() -> None:
            info = check_for_update()
            self.after(0, lambda: self._update_result(info, manual))
        threading.Thread(target=worker, daemon=True).start()

    def _update_result(self, info, manual: bool) -> None:
        if not info:
            if manual:
                messagebox.showinfo("Updates", "No newer product build was found.")
            self.status.set("App is up to date.")
            return
        asset = info.preferred_asset()
        if not messagebox.askyesno("Update available", f"{info.name} ({info.tag}) is available.\n\nDownload it now?"):
            return
        if not asset:
            webbrowser.open(info.html_url)
            return
        self.status.set(f"Downloading {asset.name}...")
        def worker() -> None:
            try:
                path = download_asset(asset)
                self.after(0, lambda: self._update_downloaded(path))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Update", str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def _update_downloaded(self, path: Path) -> None:
        self.status.set(f"Update downloaded: {path.name}")
        if messagebox.askyesno("Update downloaded", "The update is ready. Open the installer/package now?"):
            if launch_downloaded_update(path):
                self.telemetry.track_feature("update_launch", result="success")

    def _choose_live_action(self, analysis: dict) -> str | None:
        result: dict[str, str | None] = {"value": None}
        dlg = tk.Toplevel(self)
        dlg.title("Live citation manager detected")
        dlg.transient(self); dlg.grab_set(); dlg.resizable(False, False)
        frame = ttk.Frame(dlg, padding=20); frame.pack(fill="both", expand=True)
        manager = analysis.get("citation_manager", "Citation manager")
        count = analysis.get("live_field_count", 0)
        ttk.Label(frame, text="Live citation fields detected", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"{manager}: {count} live citation field(s)", style="Section.TLabel").pack(anchor="w", pady=(5, 10))
        ttk.Label(
            frame,
            text=(
                "Choose an action. Your original manuscript will not be changed.\n\n"
                "Open Word add-in: keep citation fields live and navigate inside Word.\n"
                "Create static review copy: create a separate non-live copy with citation links for reading and review."
            ),
            style="Muted.TLabel", wraplength=560, justify="left",
        ).pack(anchor="w")
        buttons = ttk.Frame(frame); buttons.pack(fill="x", pady=(18, 0))
        def choose(value: str | None) -> None:
            result["value"] = value; dlg.destroy()
        ttk.Button(buttons, text="Open Word add-in", style="Primary.TButton", command=lambda: choose("live")).pack(side="left")
        ttk.Button(buttons, text="Create static review copy", command=lambda: choose("review")).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cancel", command=lambda: choose(None)).pack(side="right")
        dlg.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        self.wait_window(dlg)
        return result["value"]

    def journal_analysis(self) -> None:
        manuscript, profile = self._require_docx(), self._profile_ref()
        if not manuscript or not profile:
            return
        data = self._run(lambda: analyze_manuscript(manuscript, profile), summary_fn=format_text_report, feature="journal_analysis")
        if data:
            self.last_report = data; self.save_report_btn.configure(state="normal")

    def audit_analysis(self) -> None:
        manuscript = self._require_docx()
        if not manuscript:
            return
        data = self._run(lambda: analyze_manuscript(manuscript, None), summary_fn=format_text_report, feature="manuscript_audit")
        if data:
            self.last_report = data; self.save_report_btn.configure(state="normal")

    def citation_navigator(self) -> None:
        manuscript = self._require_docx()
        if not manuscript:
            return
        data = self._run(lambda: analyze_citation_navigation(manuscript), feature="citation_navigation")
        if data:
            self.last_navigation = data; self.save_nav_btn.configure(state="normal")

    def create_clickable_copy(self) -> None:
        manuscript = self._require_docx()
        if not manuscript:
            return
        analysis = self._run(lambda: analyze_citation_navigation(manuscript), feature="citation_navigation")
        if not analysis:
            return
        self.last_navigation = analysis; self.save_nav_btn.configure(state="normal")
        static_review = False
        initial = f"{manuscript.stem}_navigable.docx"
        if analysis.get("live_fields"):
            choice = self._choose_live_action(analysis)
            if choice is None:
                return
            if choice == "live":
                self.open_word_addin_guide(); return
            static_review = True
            initial = f"{manuscript.stem}_linked_review_copy.docx"
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=initial, filetypes=[("Word document", "*.docx")])
        if not out or self._same_path(manuscript, out):
            return
        data = self._run(lambda: make_navigable_copy(manuscript, out, static_review_copy=static_review), feature="create_navigable_copy")
        if data and data.get("created"):
            messagebox.showinfo("Citation Navigator", "Navigable copy created.\n\nYour original manuscript was not modified. Open the copy in Word and confirm that no repair dialog appears.")
        elif data:
            messagebox.showwarning("Citation Navigator", str(data.get("message", "The navigable copy was not created.")))

    def inspect_journal_template(self) -> None:
        template = self._require_template()
        if template:
            self._run(lambda: inspect_template(template), feature="template_inspect")

    def apply_journal_template(self) -> None:
        manuscript, template = self._require_docx(), self._require_template()
        if not manuscript or not template:
            return
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"{manuscript.stem}_template_adapted.docx", filetypes=[("Word document", "*.docx")])
        if not out or self._same_path(manuscript, out):
            return
        data = self._run(lambda: retarget_from_template(manuscript, out, template).to_dict(), feature="template_retarget")
        if not data:
            return
        if not data.get("passed"):
            messagebox.showwarning(
                "Template Mode",
                "The output was withheld because a preservation, Word-package, structural, or fidelity gate failed. See Summary and Details for the exact reason.",
            )
            return
        fidelity = int(data.get("supported_fidelity_score", 0))
        coverage = int(data.get("template_coverage_score", 0))
        level = str(data.get("adaptation_level", "limited")).lower()
        if level == "high":
            messagebox.showinfo(
                "Template Mode",
                f"High-coverage template adaptation created.\n\nSupported fidelity: {fidelity}%\nTemplate coverage: {coverage}%\n\nReview the new copy in Word before submission.",
            )
        else:
            messagebox.showwarning(
                "Partial template adaptation",
                f"A safe output copy was created, but this is NOT a full template conversion.\n\nSupported fidelity: {fidelity}%\nTemplate coverage: {coverage}%\nVerdict: {data.get('verdict', '')}\n\nUnsupported template features remain. Use the Details tab to see exactly what was not transferred.",
            )

    def inspect(self) -> None:
        manuscript = self._require_docx()
        if manuscript:
            self._run(lambda: inspect_docx(manuscript).to_dict(), feature="integrity_inventory")

    def citations(self) -> None:
        manuscript = self._require_docx()
        if manuscript:
            self._run(lambda: build_citation_graph(manuscript).to_dict(), feature="citation_map")

    def readiness(self) -> None:
        manuscript, profile = self._require_docx(), self._profile_ref()
        if manuscript and profile:
            self._run(lambda: readiness_check(manuscript, profile), feature="journal_readiness")

    def retarget(self) -> None:
        manuscript, profile = self._require_docx(), self._profile_ref()
        if not manuscript or not profile:
            return
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"{manuscript.stem}_journal_converted.docx", filetypes=[("Word document", "*.docx")])
        if not out or self._same_path(manuscript, out):
            return
        data = self._run(lambda: retarget_docx(manuscript, out, profile).to_dict(), feature="journal_retarget")
        if not data:
            return
        if data.get("passed"):
            messagebox.showinfo(
                "Journal Conversion",
                f"Converted copy created.\n\nFormatting compliance: {data.get('formatting_compliance_score', 0)}%\nManuscript readiness: {data.get('manuscript_requirement_score', 0)}/100\nVerdict: {data.get('verdict', '')}\n\nOpen the copy in Microsoft Word. Final author review is still required.",
            )
        else:
            messagebox.showwarning(
                "Journal Conversion",
                "The output was withheld because a preservation, Word-package, structural, or conversion-assurance check failed. See Summary and Details for the exact reason.",
            )

    def save_html_report(self) -> None:
        if not self.last_report:
            return
        out = filedialog.asksaveasfilename(defaultextension=".html", initialfile=f"{Path(self.docx.get()).stem}_manuscript_report.html", filetypes=[("HTML report", "*.html")])
        if out:
            write_html_report(self.last_report, out)
            self.telemetry.track_feature("save_manuscript_report", result="success")
            if messagebox.askyesno(PRODUCT_NAME, "Open report in browser?"):
                webbrowser.open(Path(out).resolve().as_uri())

    def save_navigation_html(self) -> None:
        if not self.last_navigation:
            return
        out = filedialog.asksaveasfilename(defaultextension=".html", initialfile=f"{Path(self.docx.get()).stem}_citation_navigation.html", filetypes=[("HTML report", "*.html")])
        if out:
            write_navigation_html(self.last_navigation, out)
            self.telemetry.track_feature("save_navigation_report", result="success")
            if messagebox.askyesno("Citation Navigator", "Open report in browser?"):
                webbrowser.open(Path(out).resolve().as_uri())

    def _manifest_candidates(self) -> list[Path]:
        candidates: list[Path] = []
        if getattr(sys, "frozen", False):
            candidates.append(Path(sys.executable).resolve().parent / "word-addin" / "manifest.xml")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            candidates.append(Path(local) / "Programs" / "Word Journal Manuscript Converter" / "word-addin" / "manifest.xml")
        try:
            candidates.append(Path(__file__).resolve().parents[2] / "integrations" / "word-addin" / "manifest.xml")
        except IndexError:
            pass
        return candidates

    def _find_manifest(self) -> Path | None:
        return next((path for path in self._manifest_candidates() if path.exists()), None)

    def open_word_addin_guide(self) -> None:
        self.telemetry.track_feature("word_addin_guide", result="opened")
        manifest = self._find_manifest()
        dlg = tk.Toplevel(self)
        dlg.title("Word add-in setup")
        dlg.transient(self); dlg.grab_set(); dlg.geometry("700x500")
        frame = ttk.Frame(dlg, padding=20); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Word add-in setup", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text="The current setup is a Microsoft Early Access sideloading process. It is not specific to your computer. Stable distribution is intended to use Microsoft Marketplace so normal users can install the add-in directly from Word.",
            style="Muted.TLabel", wraplength=650, justify="left",
        ).pack(anchor="w", pady=(5, 12))

        ttk.Label(frame, text="Fastest Early Access route", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text=(
                "1. Open Word on the web.\n"
                "2. Open Add-ins, then the custom or upload add-in option.\n"
                "3. Upload the manifest.xml shown below.\n"
                "4. Open your manuscript and launch Citation Navigator.\n\n"
                "Desktop Word can also use Microsoft's trusted catalog method, but it requires more setup."
            ),
            style="Muted.TLabel", wraplength=650, justify="left",
        ).pack(anchor="w", pady=(3, 12))

        ttk.Label(frame, text="Local manifest", style="Section.TLabel").pack(anchor="w")
        manifest_text = str(manifest) if manifest else "Manifest not found in the current portable/install location. Use the hosted setup guide."
        path_var = tk.StringVar(value=manifest_text)
        ttk.Entry(frame, textvariable=path_var, state="readonly").pack(fill="x", pady=(4, 10))
        ttk.Label(
            frame,
            text="The app will not silently weaken Word Trust Center settings. University or company Microsoft 365 policies may limit sideloading.",
            style="Muted.TLabel", wraplength=650, justify="left",
        ).pack(anchor="w", pady=(0, 14))
        buttons = ttk.Frame(frame); buttons.pack(fill="x")
        def copy_manifest() -> None:
            if not manifest:
                messagebox.showinfo("Word add-in setup", "No local manifest was found. Open the setup guide instead.")
                return
            self.clipboard_clear(); self.clipboard_append(str(manifest)); self.update()
            self.status.set("Word add-in manifest path copied to clipboard.")
        def open_folder() -> None:
            if not manifest:
                messagebox.showinfo("Word add-in setup", "No local manifest was found. Open the setup guide instead.")
                return
            try:
                if os.name == "nt":
                    os.startfile(manifest.parent)  # type: ignore[attr-defined]
                else:
                    webbrowser.open(manifest.parent.as_uri())
            except OSError as exc:
                messagebox.showerror("Word add-in setup", str(exc))
        ttk.Button(buttons, text="Copy manifest path", command=copy_manifest).pack(side="left")
        ttk.Button(buttons, text="Open add-in folder", command=open_folder).pack(side="left", padx=7)
        ttk.Button(buttons, text="Open setup guide", style="Primary.TButton", command=lambda: webbrowser.open(WORD_ADDIN_GUIDE_URL)).pack(side="left")
        ttk.Button(buttons, text="Close", command=dlg.destroy).pack(side="right")

    def _heartbeat(self) -> None:
        self.telemetry.heartbeat_if_due(); self.after(60000, self._heartbeat)

    def _on_close(self) -> None:
        self.telemetry.close(); self.destroy()


def main() -> None:
    try:
        WordJournalManuscriptConverterApp().mainloop()
    except tk.TclError as exc:
        raise SystemExit(f"Could not start the GUI: {exc}")


if __name__ == "__main__":
    main()
