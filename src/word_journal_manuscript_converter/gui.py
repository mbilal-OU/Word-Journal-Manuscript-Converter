from __future__ import annotations

import json
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
MUTED = "#667085"
WASH = "#f5f7fb"
PAPER = "#ffffff"


class WordJournalManuscriptConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.settings = SettingsStore()
        self.telemetry = TelemetryClient(self.settings)
        self.docx = tk.StringVar()
        self.profile = tk.StringVar()
        self.template = tk.StringVar()
        self.status = tk.StringVar(value="Ready. Manuscript processing stays on this computer.")
        self.last_report: dict | None = None
        self.last_navigation: dict | None = None
        self._profile_display_to_ref: dict[str, str] = {}
        self._icon_image: tk.PhotoImage | None = None
        self.title(f"{PRODUCT_NAME} - {DISPLAY_VERSION}")
        self.geometry("1180x860")
        self.minsize(980, 720)
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
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure(".", font=("Segoe UI", 10))
        s.configure("TFrame", background=WASH)
        s.configure("TLabel", background=WASH, foreground="#17212b")
        s.configure("Muted.TLabel", background=WASH, foreground=MUTED)
        s.configure("Mode.TLabel", background=WASH, foreground=NAVY, font=("Segoe UI", 15, "bold"))
        s.configure("Section.TLabel", background=WASH, foreground="#344054", font=("Segoe UI", 10, "bold"))
        s.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        s.map("Primary.TButton", background=[("!disabled", BLUE)], foreground=[("!disabled", "white")])
        s.configure("TButton", padding=(10, 7))
        s.configure("TNotebook.Tab", padding=(13, 8), font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        header = tk.Frame(self, bg=NAVY, height=105)
        header.pack(fill="x")
        header.pack_propagate(False)
        left = tk.Frame(header, bg=NAVY)
        left.pack(side="left", fill="both", expand=True, padx=24, pady=17)
        tk.Label(left, text=PRODUCT_NAME, bg=NAVY, fg="white", font=("Segoe UI", 23, "bold")).pack(anchor="w")
        tk.Label(left, text="Convert journals  •  Navigate citations  •  Audit manuscripts  •  Apply templates", bg=NAVY, fg="#cfe0ff", font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 0))
        right = tk.Frame(header, bg=NAVY)
        right.pack(side="right", padx=24, pady=13)
        tk.Label(right, text=DISPLAY_VERSION, bg="#123f80", fg="#dbe8ff", padx=10, pady=4, font=("Segoe UI", 9, "bold")).pack(anchor="e")
        tk.Label(right, text=f"Developed by {DEVELOPER_NAME}", bg=NAVY, fg="white", font=("Segoe UI", 10, "bold")).pack(anchor="e", pady=(9, 0))
        privacy = tk.Frame(self, bg="#edf9f3", height=34)
        privacy.pack(fill="x")
        privacy.pack_propagate(False)
        tk.Label(privacy, text="LOCAL PROCESSING   •   original is never overwritten   •   live citation fields are protected", bg="#edf9f3", fg=GREEN, font=("Segoe UI", 9, "bold")).pack(side="left", padx=22, pady=7)

        body = ttk.Frame(self, padding=(20, 15, 20, 12))
        body.pack(fill="both", expand=True)
        src = ttk.LabelFrame(body, text="Select manuscript", padding=11)
        src.pack(fill="x")
        r = ttk.Frame(src)
        r.pack(fill="x")
        ttk.Label(r, text="Word document", width=14).pack(side="left")
        ttk.Entry(r, textvariable=self.docx).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(r, text="Browse...", command=self.pick_docx).pack(side="left")

        tabs = ttk.Notebook(body)
        tabs.pack(fill="x", pady=(12, 10))
        jt, nt, at = (ttk.Frame(tabs, padding=15) for _ in range(3))
        tabs.add(jt, text="Journal Conversion")
        tabs.add(nt, text="Citation Navigator")
        tabs.add(at, text="Manuscript Audit")
        self._journal_tab(jt)
        self._navigator_tab(nt)
        self._audit_tab(at)

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

        outtabs = ttk.Notebook(body)
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
        ttk.Label(footer, text=f"{DISPLAY_VERSION}  •  Developed by {DEVELOPER_NAME}  •  engine {__version__}", style="Muted.TLabel").pack(side="right")

    def _journal_tab(self, p: ttk.Frame) -> None:
        ttk.Label(p, text="Prepare for a target journal", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(p, text="Choose a verified profile, custom JSON profile, or journal-supplied Word template.", style="Muted.TLabel").pack(anchor="w", pady=(2, 8))
        row = ttk.Frame(p); row.pack(fill="x")
        ttk.Label(row, text="Journal", width=14).pack(side="left")
        self.profile_combo = ttk.Combobox(row, textvariable=self.profile, state="readonly", height=24)
        self.profile_combo.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Custom JSON...", command=self.pick_profile).pack(side="left")
        a = ttk.Frame(p); a.pack(fill="x", pady=(7, 10))
        ttk.Button(a, text="Journal analysis", style="Primary.TButton", command=self.journal_analysis).pack(side="left")
        ttk.Button(a, text="Readiness", command=self.readiness).pack(side="left", padx=7)
        ttk.Button(a, text="Safe retarget...", command=self.retarget).pack(side="left")
        ttk.Separator(p, orient="horizontal").pack(fill="x", pady=(0, 9))
        ttk.Label(p, text="Journal Word template", style="Section.TLabel").pack(anchor="w")
        tr = ttk.Frame(p); tr.pack(fill="x", pady=(4, 4))
        ttk.Label(tr, text="Template", width=14).pack(side="left")
        ttk.Entry(tr, textvariable=self.template).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(tr, text="Browse template...", command=self.pick_template).pack(side="left")
        ta = ttk.Frame(p); ta.pack(fill="x")
        ttk.Button(ta, text="Inspect template", command=self.inspect_journal_template).pack(side="left")
        ttk.Button(ta, text="Apply template safely...", style="Primary.TButton", command=self.apply_journal_template).pack(side="left", padx=7)

    def _navigator_tab(self, p: ttk.Frame) -> None:
        ttk.Label(p, text="Make citations traceable", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(p, text="No journal required. Keep EndNote/Zotero/Mendeley citations live or create a separate linked review copy.", style="Muted.TLabel").pack(anchor="w", pady=(2, 9))
        a = ttk.Frame(p); a.pack(fill="x")
        ttk.Button(a, text="Analyze navigation", style="Primary.TButton", command=self.citation_navigator).pack(side="left")
        ttk.Button(a, text="Citation map", command=self.citations).pack(side="left", padx=7)
        ttk.Button(a, text="Create navigable copy...", command=self.create_clickable_copy).pack(side="left")
        ttk.Button(a, text="Word add-in", command=self.open_word_addin_guide).pack(side="left", padx=7)

    def _audit_tab(self, p: ttk.Frame) -> None:
        ttk.Label(p, text="Audit without converting", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(p, text="Inspect structure, citations, figures, tables, equations, fields, comments and tracked changes.", style="Muted.TLabel").pack(anchor="w", pady=(2, 9))
        a = ttk.Frame(p); a.pack(fill="x")
        ttk.Button(a, text="Full manuscript audit", style="Primary.TButton", command=self.audit_analysis).pack(side="left")
        ttk.Button(a, text="Integrity inventory", command=self.inspect).pack(side="left", padx=7)

    def _load_profiles(self) -> None:
        vals: list[str] = []
        for d in sorted(list_bundled_profiles(), key=lambda x: (x.journal.lower(), x.article_type.lower())):
            label = f"{d.journal} — {d.article_type}"
            vals.append(label)
            self._profile_display_to_ref[label] = d.key
        self.profile_combo["values"] = vals
        if vals:
            self.profile.set(next((x for x in vals if x.startswith("Generic review-copy")), vals[0]))

    def pick_docx(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word document", "*.docx")])
        if path: self.docx.set(path)

    def pick_profile(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Journal profile", "*.json")])
        if not path: return
        label = f"Custom — {Path(path).name}"
        self._profile_display_to_ref[label] = path
        vals = list(self.profile_combo["values"])
        if label not in vals:
            vals.append(label); self.profile_combo["values"] = vals
        self.profile.set(label)

    def pick_template(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word template/document", "*.docx *.dotx"), ("All files", "*.*")])
        if path: self.template.set(path)

    def _profile_ref(self) -> str | None:
        return self._profile_display_to_ref.get(self.profile.get())

    def _require_docx(self) -> Path | None:
        p = Path(self.docx.get())
        if not p.exists() or p.suffix.lower() != ".docx":
            messagebox.showerror(PRODUCT_NAME, "Choose an existing .docx manuscript first.")
            return None
        return p

    def _require_template(self) -> Path | None:
        p = Path(self.template.get())
        if not p.exists() or p.suffix.lower() not in {".docx", ".dotx"}:
            messagebox.showerror("Template Mode", "Choose an existing .docx or .dotx journal template first.")
            return None
        return p

    @staticmethod
    def _same_path(a: str | Path, b: str | Path) -> bool:
        try: return Path(a).resolve() == Path(b).resolve()
        except OSError: return False

    def _show(self, data: dict, summary: str | None = None) -> None:
        self.output.delete("1.0", "end"); self.output.insert("1.0", json.dumps(data, indent=2, ensure_ascii=False))
        self.summary.delete("1.0", "end"); self.summary.insert("1.0", summary or self._compact_summary(data))
        self.status.set("Completed locally. No manuscript content was uploaded.")

    @staticmethod
    def _compact_summary(d: dict) -> str:
        if d.get("workflow") == "Citation Navigator":
            g = d.get("citation_graph", {})
            return f"CITATION NAVIGATOR\n\nManager: {d.get('citation_manager','None detected')}\nLive fields: {d.get('live_field_count',0)}\nReferences: {g.get('reference_count',0)}\nMatched keys: {g.get('matched_links',0)}\nUnresolved: {len(g.get('unmatched_citations',[]))}\n\n{d.get('capability','')}"
        if d.get("mode") == "linked-review-copy":
            return f"LINKED REVIEW COPY\n\nCreated: {'Yes' if d.get('created') else 'No'}\nLinks added: {d.get('links_added',0)}\nReferences bookmarked: {d.get('references_bookmarked',0)}\n\nKeep the original manuscript as the editable citation-manager master."
        if d.get("workflow") == "Template Mode":
            return f"TEMPLATE MODE\n\nType: {d.get('template_type','')}\nTransferable styles: {d.get('transferable_style_count',0)}\nPage settings: {'Yes' if d.get('page_format') else 'No'}"
        if "readiness_score" in d:
            lines = [f"{d.get('journal','Journal')} readiness: {d.get('readiness_score',0)}/100", ""]
            lines += [f"[{str(c.get('status','')).upper():4}] {c.get('detail','')}" for c in d.get("checks", [])]
            return "\n".join(lines)
        if "paragraphs" in d and "citation" in d:
            return f"Paragraphs: {d.get('paragraphs',0)}\nTables: {d.get('tables',0)}\nEquations: {d.get('equations',0)}\nEmbedded media: {d.get('images',0)}\nCitation-manager fields: {d.get('citation',{}).get('total_candidate_fields',0)}"
        return json.dumps(d, indent=2, ensure_ascii=False)

    def _run(self, fn, *, summary_fn=None, feature: str | None = None) -> dict | None:
        try:
            self.status.set("Working locally..."); self.update_idletasks(); data = fn()
            if data is not None: self._show(data, summary_fn(data) if summary_fn else None)
            if feature: self.telemetry.track_feature(feature, result="success")
            return data
        except Exception as exc:
            if feature: self.telemetry.track_feature(feature, result="stopped")
            messagebox.showerror(PRODUCT_NAME, str(exc)); self.status.set("Operation stopped safely.")
            return None

    def _first_run_consent(self) -> None:
        if self.settings.analytics_consent is not None: return
        answer = messagebox.askyesno("Optional anonymous usage statistics", "Help improve Word Journal Manuscript Converter by sharing anonymous product usage statistics?\n\nCollected: feature names, app version, operating system, anonymous random install ID, session duration.\n\nNever collected: manuscript text, filenames, paths, citations, references, figures, document hashes, or scientific data.\n\nYou can change this anytime under Privacy & analytics.")
        self.settings.set_analytics_consent(answer)
        if answer: self.telemetry.track("app_start", properties={"channel": "desktop"})

    def privacy_dialog(self) -> None:
        dlg = tk.Toplevel(self); dlg.title("Privacy & analytics"); dlg.transient(self); dlg.grab_set()
        f = ttk.Frame(dlg, padding=18); f.pack(fill="both", expand=True)
        ttk.Label(f, text="Privacy & analytics", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(f, text="Manuscript processing remains local. Analytics is optional and contains no manuscript content.", style="Muted.TLabel", wraplength=520).pack(anchor="w", pady=(5, 12))
        analytics = tk.BooleanVar(value=self.telemetry.enabled)
        auto_updates = tk.BooleanVar(value=bool(self.settings.get("auto_update_check", True)))
        ttk.Checkbutton(f, text="Share anonymous usage statistics", variable=analytics).pack(anchor="w", pady=3)
        ttk.Checkbutton(f, text="Automatically check GitHub for updates once per day", variable=auto_updates).pack(anchor="w", pady=3)
        b = ttk.Frame(f); b.pack(fill="x", pady=(14, 0))
        def save() -> None:
            self.settings.set_analytics_consent(analytics.get()); self.settings.set("auto_update_check", auto_updates.get()); dlg.destroy()
        ttk.Button(b, text="Save", style="Primary.TButton", command=save).pack(side="left")
        ttk.Button(b, text="Privacy policy", command=lambda: webbrowser.open(PRIVACY_URL)).pack(side="left", padx=7)
        ttk.Button(b, text="Cancel", command=dlg.destroy).pack(side="right")

    def feedback_dialog(self) -> None:
        self.telemetry.track("feedback_open")
        dlg = tk.Toplevel(self); dlg.title("Send feedback"); dlg.transient(self); dlg.grab_set(); dlg.geometry("560x460")
        f = ttk.Frame(dlg, padding=18); f.pack(fill="both", expand=True)
        ttk.Label(f, text="Send feedback", style="Mode.TLabel").pack(anchor="w")
        ttk.Label(f, text="Feedback is sent only when you press Send. Do not paste unpublished manuscript content.", style="Muted.TLabel", wraplength=510).pack(anchor="w", pady=(4, 10))
        rating, category, email, contact = tk.StringVar(value="5"), tk.StringVar(value="usability"), tk.StringVar(), tk.BooleanVar(value=False)
        ttk.Label(f, text="Rating (1-5)").pack(anchor="w"); ttk.Combobox(f, textvariable=rating, values=["5","4","3","2","1"], state="readonly").pack(fill="x", pady=(2, 7))
        ttk.Label(f, text="Category").pack(anchor="w"); ttk.Combobox(f, textvariable=category, values=["bug","feature","usability","journal","citation","template","other"], state="readonly").pack(fill="x", pady=(2, 7))
        ttk.Label(f, text="Message").pack(anchor="w"); text = tk.Text(f, height=7, wrap="word"); text.pack(fill="both", expand=True, pady=(2, 7))
        ttk.Label(f, text="Email (optional)").pack(anchor="w"); ttk.Entry(f, textvariable=email).pack(fill="x", pady=(2, 4))
        ttk.Checkbutton(f, text="You may contact me about this feedback", variable=contact).pack(anchor="w")
        buttons = ttk.Frame(f); buttons.pack(fill="x", pady=(10,0))
        def submit() -> None:
            msg = text.get("1.0", "end").strip()
            if not msg: messagebox.showerror("Feedback", "Enter a feedback message."); return
            buttons.winfo_children()[0].configure(state="disabled"); self.status.set("Sending feedback...")
            def worker() -> None:
                result = self.telemetry.submit_feedback(rating=int(rating.get()), category=category.get(), message=msg, contact_email=email.get(), consent_to_contact=contact.get())
                self.after(0, lambda: self._feedback_done(dlg, result.sent, result.error))
            threading.Thread(target=worker, daemon=True).start()
        ttk.Button(buttons, text="Send feedback", style="Primary.TButton", command=submit).pack(side="left")
        ttk.Button(buttons, text="Cancel", command=dlg.destroy).pack(side="right")

    def _feedback_done(self, dlg: tk.Toplevel, sent: bool, error: str | None) -> None:
        if sent: dlg.destroy(); self.status.set("Thank you. Feedback was sent."); messagebox.showinfo("Feedback", "Thank you. Your feedback was sent.")
        else: self.status.set("Feedback could not be sent."); messagebox.showerror("Feedback", f"Could not send feedback.\n\n{error or 'Network unavailable.'}")

    def _auto_update_check(self) -> None:
        if self.settings.get("auto_update_check", True): self.check_updates(manual=False)

    def check_updates(self, *, manual: bool) -> None:
        self.telemetry.track("update_check", properties={"action": "manual" if manual else "automatic"})
        if manual: self.status.set("Checking for updates...")
        def worker() -> None:
            info = check_for_update()
            self.after(0, lambda: self._update_result(info, manual))
        threading.Thread(target=worker, daemon=True).start()

    def _update_result(self, info, manual: bool) -> None:
        if not info:
            if manual: messagebox.showinfo("Updates", "No newer product build was found.")
            self.status.set("App is up to date."); return
        asset = info.preferred_asset()
        if not messagebox.askyesno("Update available", f"{info.name} ({info.tag}) is available.\n\nDownload it now?"):
            return
        if not asset: webbrowser.open(info.html_url); return
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
            if launch_downloaded_update(path): self.telemetry.track_feature("update_launch", result="success")

    def _choose_live_action(self, analysis: dict) -> str | None:
        choice = messagebox.askyesnocancel("Live citation manager detected", f"{analysis.get('citation_manager','Citation manager')} with {analysis.get('live_field_count',0)} live citation fields was detected.\n\nYES: keep fields live and open Word add-in navigation.\nNO: create a separate static linked review copy.\nCANCEL: do nothing.\n\nThe original manuscript is never changed.")
        if choice is None: return None
        return "live" if choice else "review"

    def journal_analysis(self) -> None:
        p, profile = self._require_docx(), self._profile_ref()
        if not p or not profile: return
        data = self._run(lambda: analyze_manuscript(p, profile), summary_fn=format_text_report, feature="journal_analysis")
        if data: self.last_report = data; self.save_report_btn.configure(state="normal")

    def audit_analysis(self) -> None:
        p = self._require_docx()
        if not p: return
        data = self._run(lambda: analyze_manuscript(p, None), summary_fn=format_text_report, feature="manuscript_audit")
        if data: self.last_report = data; self.save_report_btn.configure(state="normal")

    def citation_navigator(self) -> None:
        p = self._require_docx()
        if not p: return
        data = self._run(lambda: analyze_citation_navigation(p), feature="citation_navigation")
        if data: self.last_navigation = data; self.save_nav_btn.configure(state="normal")

    def create_clickable_copy(self) -> None:
        p = self._require_docx()
        if not p: return
        analysis = self._run(lambda: analyze_citation_navigation(p), feature="citation_navigation")
        if not analysis: return
        self.last_navigation = analysis; self.save_nav_btn.configure(state="normal")
        static_review = False; initial = f"{p.stem}_navigable.docx"
        if analysis.get("live_fields"):
            choice = self._choose_live_action(analysis)
            if choice is None: return
            if choice == "live": self.open_word_addin_guide(); return
            static_review = True; initial = f"{p.stem}_linked_review_copy.docx"
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=initial, filetypes=[("Word document","*.docx")])
        if not out or self._same_path(p, out): return
        data = self._run(lambda: make_navigable_copy(p, out, static_review_copy=static_review), feature="create_navigable_copy")
        if data and data.get("created"): messagebox.showinfo("Citation Navigator", "Navigable copy created successfully.\n\nYour original manuscript was not modified.")

    def inspect_journal_template(self) -> None:
        t = self._require_template()
        if t: self._run(lambda: inspect_template(t), feature="template_inspect")

    def apply_journal_template(self) -> None:
        p, t = self._require_docx(), self._require_template()
        if not p or not t: return
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"{p.stem}_template_retargeted.docx", filetypes=[("Word document","*.docx")])
        if not out or self._same_path(p, out): return
        data = self._run(lambda: retarget_from_template(p, out, t).to_dict(), feature="template_retarget")
        if data and data.get("passed"): messagebox.showinfo("Template Mode", "Template formatting applied to a new manuscript copy.\n\nReview the new copy in Word before submission.")

    def inspect(self) -> None:
        p = self._require_docx()
        if p: self._run(lambda: inspect_docx(p).to_dict(), feature="integrity_inventory")

    def citations(self) -> None:
        p = self._require_docx()
        if p: self._run(lambda: build_citation_graph(p).to_dict(), feature="citation_map")

    def readiness(self) -> None:
        p, profile = self._require_docx(), self._profile_ref()
        if p and profile: self._run(lambda: readiness_check(p, profile), feature="journal_readiness")

    def retarget(self) -> None:
        p, profile = self._require_docx(), self._profile_ref()
        if not p or not profile: return
        out = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"{p.stem}_retargeted.docx", filetypes=[("Word document","*.docx")])
        if out and not self._same_path(p, out): self._run(lambda: retarget_docx(p, out, profile).to_dict(), feature="journal_retarget")

    def save_html_report(self) -> None:
        if not self.last_report: return
        out = filedialog.asksaveasfilename(defaultextension=".html", initialfile=f"{Path(self.docx.get()).stem}_manuscript_report.html", filetypes=[("HTML report","*.html")])
        if out:
            write_html_report(self.last_report, out); self.telemetry.track_feature("save_manuscript_report", result="success")
            if messagebox.askyesno(PRODUCT_NAME, "Open report in browser?"): webbrowser.open(Path(out).resolve().as_uri())

    def save_navigation_html(self) -> None:
        if not self.last_navigation: return
        out = filedialog.asksaveasfilename(defaultextension=".html", initialfile=f"{Path(self.docx.get()).stem}_citation_navigation.html", filetypes=[("HTML report","*.html")])
        if out:
            write_navigation_html(self.last_navigation, out); self.telemetry.track_feature("save_navigation_report", result="success")
            if messagebox.askyesno("Citation Navigator", "Open report in browser?"): webbrowser.open(Path(out).resolve().as_uri())

    def open_word_addin_guide(self) -> None:
        self.telemetry.track_feature("word_addin_guide", result="opened"); webbrowser.open(WORD_ADDIN_GUIDE_URL)

    def _heartbeat(self) -> None:
        self.telemetry.heartbeat_if_due(); self.after(60000, self._heartbeat)

    def _on_close(self) -> None:
        self.telemetry.close(); self.destroy()


def main() -> None:
    try: WordJournalManuscriptConverterApp().mainloop()
    except tk.TclError as exc: raise SystemExit(f"Could not start the GUI: {exc}")


if __name__ == "__main__": main()
