r"""
Portfolio Report Maker - desktop window for build_report.py
============================================================
Drag the client's files into the five boxes, click "Generate Report".

This file only draws the window. All the real work (reading the files,
matching, colouring, the 4 lists) is done by the functions in
build_report.py, which must sit in the same folder.

Start it with run_app.bat (no black console window).
"""

import datetime as dt
import gc
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
import zipfile
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "settings.json")
ERROR_LOG = os.path.join(APP_DIR, "gui_error_log.txt")
sys.path.insert(0, APP_DIR)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:                     # window still works, click-to-browse only
    HAS_DND = False

try:
    import build_report as br
    IMPORT_ERROR = None
except Exception as e:                  # shown in a message box once the window is up
    br = None
    IMPORT_ERROR = f"{type(e).__name__}: {e}"


# ----------------------------------------------------------------------
# THE FIVE BOXES
# ----------------------------------------------------------------------
# key, title, hint, allowed extensions, required?, words used to auto-sort
# several files dropped at once
ZONES = [
    ("master", "1. Master file", "Portfolio Valuation (.xlsx)",
     (".xlsx", ".xlsm"), True, ("portfoliovaluation", "portfolio valuation", "valuation")),
    ("income", "2. Income file", "Income / Brokerage Rate (.xlsx)",
     (".xlsx", ".xlsm"), True, ("income", "brokerage")),
    ("scheme_bm", "3. Scheme Benchmark", "Scheme benchmark PDF",
     (".pdf",), True, ("scheme",)),
    ("broad_bm", "4. Broad Benchmark", "Broad benchmark PDF",
     (".pdf",), True, ("broad",)),
    ("sip", "5. SIP files (optional)",
     "XSIP / My SIP files - one per family member, drop as many as needed. Empty = NO SIP",
     (".xlsx", ".xlsm", ".xls", ".csv"), False, ("sip",)),
]
ZONE_BY_KEY = {z[0]: z for z in ZONES}

# ----------------------------------------------------------------------
# LOOK & FEEL
# ----------------------------------------------------------------------
C = {
    "bg": "#EEF2F6",
    "surface": "#FFFFFF",
    "border": "#DCE3EA",
    "text": "#16233B",
    "muted": "#6B778C",
    "faint": "#A3AEBD",
    "header": "#0B2545",
    "header_btn": "#17365F",
    "header_btn_hover": "#22497C",
    "header_sub": "#9FB3CF",
    "accent": "#1D6FE0",
    "accent_hover": "#155AB8",
    "accent_soft": "#EAF2FE",
    "disabled": "#AEBBCC",
    "ok": "#157347",
    "ok_bg": "#F4FBF7",
    "ok_border": "#9CD4B6",
    "ok_soft": "#E6F5ED",
    "bad": "#C0362C",
    "bad_soft": "#FDECEA",
    "warn": "#9A6700",
    "warn_soft": "#FFF5DA",
    "chip": "#EEF2F7",
}
FONT = "Segoe UI"
FONT_SB = "Segoe UI Semibold"

# title, hint, badge text, badge colour, badge background
ZONE_UI = {
    "master": ("Master file", "Portfolio Valuation export (.xlsx)", "XLSX", "#1E7145", "#E4F3EA"),
    "income": ("Income file", "Income / Brokerage Rate report (.xlsx)", "XLSX", "#1E7145", "#E4F3EA"),
    "scheme_bm": ("Scheme Benchmark", "Scheme benchmark PDF", "PDF", "#B42318", "#FCE9E7"),
    "broad_bm": ("Broad Benchmark", "Broad benchmark PDF", "PDF", "#B42318", "#FCE9E7"),
    "sip": ("SIP files", "XSIP or My SIP files - one per family member. Leave empty for NO SIP.",
            "SIP", "#5B3CC4", "#EFEAFC"),
}

GLYPHS_ICON_FONT = {"settings": "", "folder": "", "check": "", "close": "",
                    "upload": "", "error": "", "warn": "", "info": "",
                    "play": ""}
GLYPHS_PLAIN = {"settings": "⚙", "folder": "▤", "check": "✔", "close": "✕", "upload": "⇧",
                "error": "✖", "warn": "!", "info": "i", "play": "▶"}



class UserError(Exception):
    """A problem the user can fix (wrong file, missing file ...). Shown as-is."""


# ----------------------------------------------------------------------
# SETTINGS (settings.json next to this file)
# ----------------------------------------------------------------------
def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
        json.dump(settings, fh, indent=2)


# ----------------------------------------------------------------------
# THE REPORT RUN (uses build_report.py's own functions, same order as its main())
# ----------------------------------------------------------------------
def detect_client_name(master_path, records):
    """'PortfolioValuation-KIRTAN VERMA-26-09-2026-01-28-10.xlsx' -> 'KIRTAN VERMA'.
    If the file name doesn't have it, use the first client heading inside the file."""
    base = os.path.splitext(os.path.basename(master_path))[0]
    m = re.match(r"portfolio\s*valuation[\s_-]+(.+?)[\s_-]+\d{1,2}-\d{1,2}-\d{4}", base, re.I)
    if m and m.group(1).strip():
        return m.group(1).strip().upper()
    for rec in records:
        if rec["type"] == "client_header":
            m = re.match(r"^(.*?)\s*\([A-Z0-9]+\)\s*$", rec["label"])
            return (m.group(1) if m else rec["label"]).strip().upper()
    raise UserError("Could not find the client name in the Master file.\n"
                    "Is this really a Portfolio Valuation export?")


def safe_file_part(name):
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    return re.sub(r"\s+", "_", name.strip()) or "CLIENT"


def check_files(files):
    """Catch the easy mistakes before any reading starts."""
    for key, title, _hint, exts, required, _words in ZONES:
        paths = files.get(key) or []
        if required and not paths:
            raise UserError(f"{title} is empty.\nDrop the file into that box and try again.")
        for p in paths:
            if not os.path.isfile(p):
                raise UserError(f"{title}: this file no longer exists (moved or renamed?)\n\n{p}")
            if not p.lower().endswith(exts):
                raise UserError(f"{title}: '{os.path.basename(p)}' is the wrong type of file.\n"
                                f"Expected: {', '.join(exts)}")

    everything = [os.path.normcase(os.path.abspath(p)) for ps in files.values() for p in ps]
    if len(everything) != len(set(everything)):
        raise UserError("The same file has been put into two boxes.\nEach box needs its own file.")

    # Swapped boxes (going by the file names)
    def name(key):
        return os.path.basename(files[key][0]).lower()
    s, b = name("scheme_bm"), name("broad_bm")
    if ("broad" in s and "scheme" not in s) or ("scheme" in b and "broad" not in b):
        raise UserError("The Scheme and Broad benchmark PDFs look swapped (going by their names).\n"
                        "Scheme benchmark PDF -> box 3, Broad benchmark PDF -> box 4.")
    m, i = name("master"), name("income")
    if "income" in m or "valuation" in i:
        raise UserError("The Master and Income files look swapped (going by their names).\n"
                        "Portfolio Valuation -> box 1, Income Report -> box 2.")


def generate_report(files, reports_folder):
    """Runs build_report.py's logic. Returns a result dict or raises."""
    check_files(files)
    master = files["master"][0]
    income = files["income"][0]
    scheme_bm = files["scheme_bm"][0]
    broad_bm = files["broad_bm"][0]
    sip_paths = files.get("sip") or []

    brokerage_rates = br.load_brokerage_rates(income)

    notes = []
    sip_totals = br.SipLookup()
    if sip_paths:
        if len(sip_paths) > 1 and not hasattr(br, "find_all_sip_files"):
            raise UserError("Several SIP files were dropped, but this version of build_report.py "
                            "can only read ONE SIP file.\nKeep one SIP file in box 5.")
        try:
            sip_totals = br.load_sip_totals(sip_paths if len(sip_paths) > 1 else sip_paths[0])
        except Exception as e:
            raise UserError(f"Could not read the SIP file.\n\n{e}")
        sip_errors = getattr(sip_totals, "errors", [])
        if len(sip_totals) == 0 and sip_errors:
            # every SIP file failed - stop, rather than quietly making a "NO SIP" report
            raise UserError("None of the SIP files could be used:\n\n" + "\n".join(sip_errors))
        if len(sip_totals) == 0:
            notes.append("SIP file had no usable SIPs -> NO SIP")
    else:
        notes.append("NO SIP")
    has_sip = bool(sip_paths) and len(sip_totals) > 0

    scheme_sections = br.load_benchmark_diffs(scheme_bm)
    broad_sections = br.load_benchmark_diffs(broad_bm)
    records = br.parse_master(master)

    master_clients = {r["client_name"] for r in records
                      if r["type"] == "scheme" and r.get("client_name")}
    if not master_clients:
        raise UserError("No schemes were found in the Master file.")
    client = detect_client_name(master, records)

    # Benchmark PDFs must belong to the same client/family as the Master file
    for label, sections in (("Scheme Benchmark PDF", scheme_sections),
                            ("Broad Benchmark PDF", broad_sections)):
        if not sections:
            raise UserError(f"The {label} has no client sections in it.\n"
                            f"Is it the right PDF (ComprehensivePortfolio)?")
        if not master_clients & set(sections):
            raise UserError(f"The {label} does not belong to this client.\n\n"
                            f"Master file   : {', '.join(sorted(master_clients))}\n"
                            f"{label}: {', '.join(sorted(sections))}")

    out_path = os.path.join(reports_folder, f"Portfolio_Report_{safe_file_part(client)}.xlsx")
    if os.path.exists(out_path):
        try:
            with open(out_path, "a"):
                pass
        except PermissionError:
            raise UserError(f"'{os.path.basename(out_path)}' is open in Excel.\n"
                            f"Close it and click Generate again.")

    flags = br.build_output(
        records, brokerage_rates, sip_totals, scheme_sections, broad_sections,
        report_title="Portfolio Valuation Report", client_line="", rm_line="",
        printed_line="", has_sip=has_sip, out_path=out_path,
    )

    unused = sip_totals.unused() if has_sip else []
    mismatch = getattr(sip_totals, "name_mismatch", []) if has_sip else []
    members = sorted(master_clients)
    file_clients = getattr(sip_totals, "file_clients", {})
    # SIP files whose members are nowhere in this Master file (another family's file?)
    foreign = [f"{fname} -> {', '.join(sorted(names))}" for fname, names in file_clients.items()
               if names and not any(br.client_names_match(n, m) for n in names for m in members)]
    summary = sip_totals.summary() if sip_paths and hasattr(sip_totals, "summary") else None
    if len(members) > 1:
        notes.insert(0, f"{len(members)} members")
    if summary:
        notes.append(f"SIP: {summary['updated']} updated · {summary['unmatched']} unmatched · "
                     f"{summary['errors']} error{'s' if summary['errors'] != 1 else ''}")
    if foreign:
        notes.append(f"{len(foreign)} SIP file(s) of another client")
    if flags:
        notes.append(f"{len(flags)} blank cell(s)")

    detail_lines = [f"Saved to: {out_path}", "", f"Members in Master file: {', '.join(members)}", ""]
    if summary:
        detail_lines.append(
            f"SIP UPLOAD SUMMARY:  {summary['updated']} updated (Rs {summary['updated_amount']:,.0f})"
            f"  |  {summary['unmatched']} unmatched  |  {summary['errors']} error(s)")
        detail_lines.append(f"   ({summary['not_counted']} rows not counted - inactive / 0 this month / "
                            f"already in XSIP / a newer file;  {summary['duplicates']} duplicate(s) not added twice)")
        detail_lines.append("")
        for fname, info in getattr(sip_totals, "file_info", {}).items():
            names = ", ".join(sorted(file_clients.get(fname, ()))) or "(no member name)"
            extra = f", {info['replaced']} already in XSIP / a newer file" if info.get("replaced") else ""
            detail_lines.append(f"{info['kind']}: {fname}")
            detail_lines.append(f"   rule: {info['rule']}  |  member(s): {names}")
            detail_lines.append(f"   counted: {info['counted']}  |  not counted: {info['skipped']}{extra}")
        for fname, same in getattr(sip_totals, "duplicate_files", []):
            detail_lines.append(f"DUPLICATE: {fname} is the same file as {same} - read once only")
        if sip_totals.errors:
            detail_lines.append("")
            detail_lines.append("ERRORS (nothing used from these):")
            detail_lines += [f" - {e}" for e in sip_totals.errors]
        fallbacks = [p for p in sip_totals.placed if p[4]]
        if fallbacks:
            detail_lines.append("")
            detail_lines.append("Last month's amount used (current month was 0):")
            detail_lines += [f" - {m}: '{s}' / {f}  Rs {a:,.0f}  ({n})" for m, s, f, a, n in fallbacks]
        detail_lines.append("")
    if foreign:
        detail_lines.append("These SIP files belong to nobody in this Master file (nothing used from them):")
        detail_lines += [f" - {x}" for x in foreign]
        detail_lines.append("")
    if mismatch:
        detail_lines.append("Same scheme + folio, but registered under a DIFFERENT member (left blank):")
        detail_lines += [f" - '{s}' / folio {f}: SIP file says {c}, Master row is {m} (Rs {a:,.0f})"
                         for c, m, s, f, a in mismatch]
        detail_lines.append("")
    if flags:
        detail_lines.append(f"{len(flags)} cell(s) left BLANK (no confident match):")
        detail_lines += [f" - {f}" for f in flags]
        detail_lines.append("")
    if unused:
        detail_lines.append(f"{len(unused)} SIP(s) in the SIP file did not match any Master row:")
        detail_lines += [f" - '{s}' / folio {f}  (Rs {a:,.0f})" for s, f, a in unused]

    return {"client": client, "out_path": out_path, "notes": notes,
            "details": "\n".join(detail_lines).strip(),
            "sip_errors": list(getattr(sip_totals, "errors", []))}


def release_file_handles(paths):
    """openpyxl's read-only mode (used for the SIP file) leaves the spreadsheet open after
    reading, which blocks moving it. The report is already saved, so close those leftovers."""
    wanted = {os.path.normcase(os.path.abspath(p)) for p in paths}
    gc.collect()
    for obj in gc.get_objects():
        if isinstance(obj, zipfile.ZipFile) and isinstance(obj.filename, str):
            if os.path.normcase(os.path.abspath(obj.filename)) in wanted:
                try:
                    obj.close()
                except Exception:
                    pass


def move_to_processed(paths):
    """Move each input file into a 'processed' folder next to it. Returns problems (if any)."""
    release_file_handles(paths)
    problems = []
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    for p in paths:
        dest_dir = os.path.join(os.path.dirname(p), "processed")
        dest = os.path.join(dest_dir, os.path.basename(p))
        if os.path.exists(dest):
            root, ext = os.path.splitext(os.path.basename(p))
            dest = os.path.join(dest_dir, f"{root} ({stamp}){ext}")
        for attempt in range(5):
            try:
                os.makedirs(dest_dir, exist_ok=True)
                os.rename(p, dest)          # same drive -> a true move, never a half-copy
                break
            except PermissionError as e:
                if attempt == 4:
                    problems.append(f"{os.path.basename(p)}: still open in another program ({e.strerror})")
                else:
                    gc.collect()
                    time.sleep(0.4)
            except OSError as e:
                problems.append(f"{os.path.basename(p)}: {e}")
                break
    return problems


def friendly_error(exc):
    if isinstance(exc, UserError):
        return str(exc)
    if isinstance(exc, PermissionError):
        return (f"A file is locked (probably open in Excel or a PDF viewer).\n"
                f"Close it and try again.\n\n{exc}")
    if isinstance(exc, FileNotFoundError):
        return f"A file could not be found (moved or renamed?).\n\n{exc}"
    return (f"Something went wrong while building the report:\n\n{type(exc).__name__}: {exc}\n\n"
            f"Check that each file is in the right box. Full details were saved to "
            f"gui_error_log.txt")


# ----------------------------------------------------------------------
# SMALL WIDGETS
# ----------------------------------------------------------------------
class Icons:
    """Windows icon font (Segoe Fluent / MDL2) if present, else plain symbols."""
    family = None

    @classmethod
    def setup(cls, root):
        fams = set(tkfont.families(root))
        for f in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
            if f in fams:
                cls.family = f
                break

    @classmethod
    def get(cls, name):
        return (GLYPHS_ICON_FONT if cls.family else GLYPHS_PLAIN)[name]

    @classmethod
    def font(cls, size):
        return (cls.family, size) if cls.family else (FONT, size)


def card(parent, **kw):
    return tk.Frame(parent, bg=C["surface"], highlightthickness=1,
                    highlightbackground=C["border"], highlightcolor=C["border"], **kw)


class IconButton(tk.Frame):
    """Flat button: optional icon + text, hover colour, can be greyed out."""

    def __init__(self, master, text, command, icon=None, bg=C["accent"], hover=C["accent_hover"],
                 fg="white", font=(FONT_SB, 10), padx=14, pady=7, icon_size=11):
        super().__init__(master, bg=bg, cursor="hand2")
        self.command, self._bg, self._hover, self._fg = command, bg, hover, fg
        self.enabled = True
        inner = tk.Frame(self, bg=bg)
        inner.pack(padx=padx, pady=pady)
        self.parts = [self, inner]
        self.icon = None
        if icon:
            self.icon = tk.Label(inner, text=Icons.get(icon), font=Icons.font(icon_size), bg=bg, fg=fg)
            self.icon.pack(side="left", padx=(0, 8))
            self.parts.append(self.icon)
        self.label = tk.Label(inner, text=text, font=font, bg=bg, fg=fg)
        self.label.pack(side="left")
        self.parts.append(self.label)
        for w in self.parts:
            w.bind("<Enter>", lambda e: self._set_bg(self._hover if self.enabled else self._bg))
            w.bind("<Leave>", lambda e: self._set_bg(self._bg))
            w.bind("<Button-1>", lambda e: self.command())

    def _set_bg(self, colour):
        for w in self.parts:
            w.configure(bg=colour)

    def style(self, text=None, bg=None, hover=None, fg=None, cursor=None):
        if text is not None:
            self.label.configure(text=text)
        if bg:
            self._bg, self._hover = bg, hover or bg
            self._set_bg(bg)
        if fg:
            for w in (self.label, self.icon):
                if w:
                    w.configure(fg=fg)
        if cursor:
            for w in self.parts:
                w.configure(cursor=cursor)


class Link(tk.Label):
    def __init__(self, master, text, command, fg=C["accent"], size=9, bg=C["surface"]):
        super().__init__(master, text=text, fg=fg, bg=bg, font=(FONT_SB, size), cursor="hand2")
        self._font_plain = tkfont.Font(family=FONT_SB, size=size)
        self._font_line = tkfont.Font(family=FONT_SB, size=size, underline=True)
        self.configure(font=self._font_plain)
        self.bind("<Enter>", lambda e: self.configure(font=self._font_line))
        self.bind("<Leave>", lambda e: self.configure(font=self._font_plain))
        self.bind("<Button-1>", lambda e: (command(), "break")[1])


def human_size(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return ""
    return f"{n / 1024:,.0f} KB" if n < 1024 * 1024 else f"{n / 1024 / 1024:,.1f} MB"


# ----------------------------------------------------------------------
# ONE DROP BOX
# ----------------------------------------------------------------------
class DropZone(tk.Frame):
    def __init__(self, master, app, key, number):
        _, _title, _hint, exts, required, _ = ZONE_BY_KEY[key]
        title, hint, badge, badge_fg, badge_bg = ZONE_UI[key]
        super().__init__(master, bg=C["surface"], highlightthickness=1,
                         highlightbackground=C["border"], cursor="hand2")
        self.app, self.key, self.exts, self.hint = app, key, exts, hint
        self.paths = []
        self._drag = self._mouse = False
        self.columnconfigure(1, weight=1)

        tile = tk.Frame(self, bg=badge_bg, width=42, height=42)
        tile.grid(row=0, column=0, rowspan=3, sticky="nw", padx=(12, 12), pady=11)
        tile.grid_propagate(False)
        tile.pack_propagate(False)
        tag = tk.Label(tile, text=badge, bg=badge_bg, fg=badge_fg, font=(FONT_SB, 8))
        tag.pack(expand=True)

        head = tk.Frame(self, bg=C["surface"])
        head.grid(row=0, column=1, sticky="ew", pady=(10, 0))
        self.title = tk.Label(head, text=f"{number}. {title}", font=(FONT_SB, 11),
                              fg=C["text"], bg=C["surface"])
        self.title.pack(side="left")
        chip = tk.Label(head, text="OPTIONAL", font=(FONT_SB, 7), fg=C["muted"], bg=C["chip"], padx=6)
        if not required:
            chip.pack(side="left", padx=8)

        self._body_font = tkfont.Font(family=FONT, size=9)
        self._width = 300
        self.body = tk.Label(self, text="", font=self._body_font, fg=C["muted"], bg=C["surface"],
                             anchor="w", justify="left")
        self.body.grid(row=1, column=1, sticky="ew", pady=(2, 0))
        self.meta = tk.Label(self, text="", font=(FONT, 8), fg=C["faint"], bg=C["surface"],
                             anchor="w", justify="left")
        self.meta.grid(row=2, column=1, sticky="ew", pady=(0, 10))

        side = tk.Frame(self, bg=C["surface"])
        side.grid(row=0, column=2, rowspan=3, sticky="ne", padx=(8, 12), pady=10)
        self.state_icon = tk.Label(side, text="", font=Icons.font(15), bg=C["surface"])
        self.state_icon.pack(anchor="e")
        self.remove = Link(side, "Remove", self.clear, fg=C["bad"], size=8)

        self._bg_widgets = [head, self.title, self.body, self.meta, side, self.state_icon, self.remove]
        for w in (self, tile, tag, head, self.title, chip, self.body, self.meta, side, self.state_icon):
            w.bind("<Button-1>", self.browse)
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            if HAS_DND:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop)
                w.dnd_bind("<<DropEnter>>", self._drag_enter)
                w.dnd_bind("<<DropLeave>>", self._drag_leave)
        self.bind("<Configure>", lambda e: self._wrap(e.width))
        self._paint()

    # ---------- look ----------
    def _wrap(self, width):
        width = max(140, width - 150)
        if abs(width - self._width) > 4:
            self._width = width
            self._paint()

    def _fit(self, text):
        """One line only: shorten in the MIDDLE so the start and the extension stay visible."""
        f, w = self._body_font, self._width
        if f.measure(text) <= w:
            return text
        keep_end = 14
        start, end = text[:-keep_end], text[-keep_end:]
        while start and f.measure(start + "…" + end) > w:
            start = start[:-1]
        return start + "…" + end

    def _paint(self):
        full = bool(self.paths)
        if self._drag or (self._mouse and not self.app.busy):
            bg, border = C["accent_soft"], C["accent"]
        elif full:
            bg, border = C["ok_bg"], C["ok_border"]
        else:
            bg, border = C["surface"], C["border"]
        self.configure(bg=bg, highlightbackground=border, highlightcolor=border)
        for w in self._bg_widgets:
            w.configure(bg=bg)
        if full:
            shown = self.paths if len(self.paths) <= 3 else self.paths[:2]
            names = "\n".join(self._fit(os.path.basename(p)) for p in shown)
            if len(self.paths) > len(shown):
                names += f"\n+ {len(self.paths) - len(shown)} more"
            self.body.configure(text=names, fg=C["text"])
            if self.key == "sip":
                n = len(self.paths)
                self.meta.configure(text=f"{n} file{'s' if n > 1 else ''} added · drop or click to add more")
            else:
                self.meta.configure(text=f"{human_size(self.paths[0])} · ready")
            self.state_icon.configure(text=Icons.get("check"), fg=C["ok"])
            self.remove.pack(anchor="e", pady=(4, 0))
        else:
            self.body.configure(text=self._fit(self.hint), fg=C["muted"])
            self.meta.configure(text="Drop file here or click to browse" if HAS_DND
                                else "Click to browse")
            self.state_icon.configure(text=Icons.get("upload"), fg=C["faint"])
            self.remove.pack_forget()

    def _inside(self):
        try:
            w = self.winfo_containing(*self.winfo_pointerxy())
        except (KeyError, tk.TclError):
            return False
        while w is not None:
            if w is self:
                return True
            w = w.master
        return False

    def _on_enter(self, _e):
        if not self._mouse:
            self._mouse = True
            self._paint()

    def _on_leave(self, _e):
        if self._mouse and not self._inside():
            self._mouse = False
            self._paint()

    def _drag_enter(self, e):
        self._drag = True
        self._paint()
        return e.action

    def _drag_leave(self, e):
        self._drag = False
        self._paint()
        return e.action

    # ---------- files ----------
    def set_paths(self, paths):
        if self.key == "sip":
            # SIP box collects files (one per family member); Remove clears them all
            paths = self.paths + [os.path.abspath(p) for p in paths]
            seen = set()
            paths = [p for p in paths if not (os.path.normcase(p) in seen or seen.add(os.path.normcase(p)))]
        else:
            paths = paths[:1]
        self.paths = [os.path.abspath(p) for p in paths]
        self._paint()
        self.app.on_zone_change()

    def clear(self):
        if self.app.busy:
            return
        self.paths = []
        self._paint()
        self.app.on_zone_change()

    def browse(self, _event=None):
        if self.app.busy:
            return
        types = [(self.hint, " ".join("*" + e for e in self.exts)), ("All files", "*.*")]
        start = os.path.dirname(self.paths[0]) if self.paths else self.app.last_browse_dir
        if self.key == "sip":
            picked = list(filedialog.askopenfilenames(title=self.hint, filetypes=types,
                                                      initialdir=start))
        else:
            one = filedialog.askopenfilename(title=self.hint, filetypes=types, initialdir=start)
            picked = [one] if one else []
        self._mouse = False
        if picked:
            self.app.last_browse_dir = os.path.dirname(picked[0])
            self.set_paths(picked)
        else:
            self._paint()

    def on_drop(self, event):
        self._drag = False
        self._paint()
        if self.app.busy:
            return event.action
        paths = [p for p in self.tk.splitlist(event.data) if os.path.isfile(p)]
        if not paths:
            return event.action
        self.app.last_browse_dir = os.path.dirname(paths[0])
        if len(paths) == 1 or self.key == "sip" and all(
                p.lower().endswith(self.exts) and "sip" in os.path.basename(p).lower() for p in paths):
            self.set_paths(paths)
        else:
            self.app.auto_sort(paths)       # several files -> each goes to its own box
        return event.action


# ----------------------------------------------------------------------
# MESSAGE STRIP (success / error shown inside the window)
# ----------------------------------------------------------------------
class Banner(tk.Frame):
    KINDS = {"ok": ("ok", "ok_soft", "check"), "bad": ("bad", "bad_soft", "error"),
             "warn": ("warn", "warn_soft", "warn"), "info": ("accent", "accent_soft", "info")}

    def __init__(self, master, on_show=None, on_hide=None):
        super().__init__(master, highlightthickness=1)
        self._on_show, self._on_hide = on_show, on_hide
        self.bar = tk.Frame(self, width=4)
        self.bar.pack(side="left", fill="y")
        self.icon = tk.Label(self, font=Icons.font(15))
        self.icon.pack(side="left", anchor="n", padx=(14, 10), pady=10)
        self.textbox = tk.Frame(self)
        self.textbox.pack(side="left", fill="both", expand=True, pady=8)
        self.title = tk.Label(self.textbox, font=(FONT_SB, 11), anchor="w", justify="left")
        self.title.pack(fill="x")
        self.msg = tk.Label(self.textbox, font=(FONT, 9), anchor="w", justify="left")
        self.msg.pack(fill="x", pady=(2, 0))
        self.actions = tk.Frame(self.textbox)
        self.actions.pack(fill="x", pady=(6, 0))
        self.close = tk.Label(self, text=Icons.get("close"), font=Icons.font(9), cursor="hand2")
        self.close.pack(side="right", anchor="n", padx=12, pady=12)
        self.close.bind("<Button-1>", lambda e: self.hide())
        self.bind("<Configure>", lambda e: self.msg.configure(wraplength=max(200, e.width - 110)))

    def show(self, kind, title, message, actions=()):
        fg_key, bg_key, icon = self.KINDS[kind]
        fg, bg = C[fg_key], C[bg_key]
        self.configure(bg=bg, highlightbackground=bg, highlightcolor=bg)
        self.bar.configure(bg=fg)
        for w in (self.textbox, self.actions, self.close):
            w.configure(bg=bg)
        self.icon.configure(text=Icons.get(icon), fg=fg, bg=bg)
        self.title.configure(text=title, fg=C["text"], bg=bg)
        self.msg.configure(text=message, fg=C["text"], bg=bg)
        self.close.configure(fg=C["muted"])
        for w in self.actions.winfo_children():
            w.destroy()
        for text, cmd in actions:
            Link(self.actions, text, cmd, fg=fg, bg=bg).pack(side="left", padx=(0, 16))
        if actions:
            self.actions.pack(fill="x", pady=(6, 0))
        else:
            self.actions.pack_forget()
        if self._on_show:
            self._on_show()
        self.grid()

    def hide(self):
        self.grid_remove()
        if self._on_hide:
            self._on_hide()


# ----------------------------------------------------------------------
# MAIN WINDOW
# ----------------------------------------------------------------------
class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title("Portfolio Report Maker")
        self.root.configure(bg=C["bg"])
        self.root.minsize(1020, 700)
        Icons.setup(self.root)
        self.settings = load_settings()
        self.busy = False
        self.results = queue.Queue()
        self.done = self.failed = 0
        self.row_info = {}
        self.last_browse_dir = os.path.expanduser("~")

        self._style()
        self._build()
        self._refresh_folder_label()
        self.on_zone_change()
        self._center(1240, 800)
        self.root.bind("<Control-Return>", lambda e: self.on_generate())
        self.root.after(250, self._startup_checks)

    # ---------- layout ----------
    def _style(self):
        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Session.Treeview", background=C["surface"], fieldbackground=C["surface"],
                     foreground=C["text"], rowheight=32, borderwidth=0, font=(FONT, 9))
        st.configure("Session.Treeview.Heading", background="#F5F7FA", foreground=C["muted"],
                     font=(FONT_SB, 9), relief="flat", borderwidth=0, padding=(6, 7))
        st.map("Session.Treeview.Heading", background=[("active", "#EBEFF5")])
        st.map("Session.Treeview", background=[("selected", C["accent_soft"])],
               foreground=[("selected", C["text"])])
        st.layout("Session.Treeview", [("Session.Treeview.treearea", {"sticky": "nswe"})])
        st.configure("Slim.Vertical.TScrollbar", troughcolor=C["surface"], background="#CFD8E3",
                     bordercolor=C["surface"], arrowcolor=C["muted"], lightcolor="#CFD8E3",
                     darkcolor="#CFD8E3", gripcount=0, arrowsize=11)
        st.configure("Accent.Horizontal.TProgressbar", troughcolor="#DCE6F5", background=C["accent"],
                     bordercolor="#DCE6F5", lightcolor=C["accent"], darkcolor=C["accent"], thickness=5)

    def _center(self, w, h):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = min(w, int(sw * 0.94)), min(h, int(sh * 0.9))
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2 - 20)}")

    def _build(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self._build_header()

        body = tk.Frame(self.root, bg=C["bg"])
        body.grid(row=1, column=0, sticky="nsew", padx=24, pady=(20, 8))
        body.columnconfigure(0, weight=3, minsize=560)
        body.columnconfigure(1, weight=2, minsize=340)
        body.rowconfigure(0, weight=1)
        left = tk.Frame(body, bg=C["bg"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        right = tk.Frame(body, bg=C["bg"])
        right.grid(row=0, column=1, sticky="nsew")
        self._build_files(left)
        self._build_side(right)

        foot = tk.Frame(self.root, bg=C["bg"])
        foot.grid(row=2, column=0, sticky="ew", padx=26, pady=(0, 10))
        version = getattr(br, "SCRIPT_VERSION", "not loaded") if br else "not loaded"
        tk.Label(foot, text=f"Report engine: build_report.py  {version.split(' (')[0]}",
                 font=(FONT, 8), fg=C["faint"], bg=C["bg"]).pack(side="left")
        tk.Label(foot, text=("Drag & drop on" if HAS_DND else "Drag & drop off - click boxes to browse")
                 + "   ·   Ctrl+Enter = Generate", font=(FONT, 8), fg=C["faint"],
                 bg=C["bg"]).pack(side="right")

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C["header"])
        hdr.grid(row=0, column=0, sticky="ew")
        inner = tk.Frame(hdr, bg=C["header"])
        inner.pack(fill="x", padx=24, pady=11)
        logo = tk.Frame(inner, bg=C["accent"], width=40, height=40)
        logo.pack(side="left")
        logo.pack_propagate(False)
        tk.Label(logo, text="PR", bg=C["accent"], fg="white", font=(FONT_SB, 12)).pack(expand=True)
        titles = tk.Frame(inner, bg=C["header"])
        titles.pack(side="left", padx=14)
        tk.Label(titles, text="Portfolio Report Maker", font=(FONT_SB, 15), fg="white",
                 bg=C["header"]).pack(anchor="w")
        tk.Label(titles, text="Portfolio review report generator", font=(FONT, 9),
                 fg=C["header_sub"], bg=C["header"]).pack(anchor="w")
        for text, icon, cmd in (("Settings", "settings", self.choose_folder),
                                ("Reports folder", "folder", self.open_reports_folder)):
            IconButton(inner, text, cmd, icon=icon, bg=C["header_btn"], hover=C["header_btn_hover"],
                       font=(FONT, 10), padx=14, pady=8, icon_size=10).pack(side="right", padx=(8, 0))

    def _build_files(self, parent):
        parent.columnconfigure(0, weight=1)
        box = card(parent)
        box.grid(row=0, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)

        head = tk.Frame(box, bg=C["surface"])
        head.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 0))
        tk.Label(head, text="Client files", font=(FONT_SB, 13), fg=C["text"],
                 bg=C["surface"]).pack(side="left")
        Link(head, "Clear all", self.clear_all, fg=C["muted"]).pack(side="right")
        self.ready_pill = tk.Label(head, text="", font=(FONT_SB, 9), padx=10, pady=2)
        self.ready_pill.pack(side="right", padx=12)
        # the hint line and the message strip share one spot at the top of the card
        self.hint_line = tk.Label(box, text="Drag files onto a box or click to browse  ·  drop all files "
                                            "at once and each goes to its own box",
                                  font=(FONT, 9), fg=C["muted"], bg=C["surface"], anchor="w")
        self.hint_line.grid(row=1, column=0, sticky="ew", padx=20, pady=(3, 10))
        self.banner = Banner(box, on_show=lambda: self._compact(True),
                             on_hide=lambda: self._compact(False))
        self.banner.grid(row=1, column=0, sticky="ew", padx=20, pady=(10, 12))
        self.banner.hide()

        grid = tk.Frame(box, bg=C["surface"])
        grid.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 18))
        grid.columnconfigure((0, 1), weight=1, uniform="z")
        self.zones = {}
        for n, (key, *_rest) in enumerate(ZONES):
            z = DropZone(grid, self, key, n + 1)
            if key == "sip":
                z.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
            else:
                z.grid(row=n // 2, column=n % 2, sticky="nsew",
                       padx=(0, 5) if n % 2 == 0 else (5, 0), pady=(0 if n < 2 else 10, 0))
            self.zones[key] = z

        self.go = IconButton(parent, "Generate Report", self.on_generate, icon="play",
                             font=(FONT_SB, 13), pady=13, icon_size=12)
        self.go.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        self.progress = ttk.Progressbar(parent, mode="indeterminate",
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        self.progress.grid_remove()

    def _build_side(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        stats = tk.Frame(parent, bg=C["bg"])
        stats.grid(row=0, column=0, sticky="ew")
        stats.columnconfigure((0, 1, 2), weight=1, uniform="s")
        self.stat = {}
        for i, (key, label, colour) in enumerate((("total", "Reports made", C["text"]),
                                                  ("done", "Successful", C["ok"]),
                                                  ("failed", "Failed", C["bad"]))):
            c = card(stats)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0 if i == 2 else 6))
            num = tk.Label(c, text="0", font=(FONT_SB, 20), fg=colour, bg=C["surface"])
            num.pack(anchor="w", padx=16, pady=(12, 0))
            tk.Label(c, text=label, font=(FONT, 9), fg=C["muted"],
                     bg=C["surface"]).pack(anchor="w", padx=16, pady=(0, 12))
            self.stat[key] = num

        act = card(parent)
        act.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        act.columnconfigure(0, weight=1)
        act.rowconfigure(2, weight=1)
        tk.Label(act, text="Session activity", font=(FONT_SB, 12), fg=C["text"],
                 bg=C["surface"]).grid(row=0, column=0, sticky="w", padx=18, pady=(16, 0))
        tk.Label(act, text="Every report made since the app was opened. Double-click a row for details.",
                 font=(FONT, 9), fg=C["muted"], bg=C["surface"], anchor="w", justify="left",
                 wraplength=360).grid(row=1, column=0, sticky="ew", padx=18, pady=(2, 10))
        wrap = tk.Frame(act, bg=C["surface"])
        wrap.grid(row=2, column=0, sticky="nsew", padx=(10, 6), pady=(0, 12))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.log = ttk.Treeview(wrap, columns=("client", "time", "status"), show="headings",
                                style="Session.Treeview", selectmode="browse")
        for c, text, w, stretch in (("client", "CLIENT", 170, True), ("time", "TIME", 74, False),
                                    ("status", "STATUS", 96, False)):
            self.log.heading(c, text=text, anchor="w")
            self.log.column(c, width=w, minwidth=w if not stretch else 120, stretch=stretch, anchor="w")
        self.log.tag_configure("bad", foreground=C["bad"])
        self.log.tag_configure("odd", background="#FAFBFD")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.log.yview,
                           style="Slim.Vertical.TScrollbar")

        def autoscroll(lo, hi):                  # scrollbar only when the list is long
            if float(lo) <= 0.0 and float(hi) >= 1.0:
                sb.grid_remove()
            else:
                sb.grid()
            sb.set(lo, hi)
        self.log.configure(yscrollcommand=autoscroll)
        self.log.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        sb.grid_remove()
        self.log.bind("<Double-1>", self.on_log_double_click)
        self.empty = tk.Label(self.log, text="No reports yet.\nYour first report will appear here.",
                              font=(FONT, 9), fg=C["faint"], bg=C["surface"], justify="center")
        self.empty.place(relx=0.5, rely=0.45, anchor="center")

        out = card(parent)
        out.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        out.columnconfigure(1, weight=1)
        tk.Label(out, text=Icons.get("folder"), font=Icons.font(16), fg=C["accent"],
                 bg=C["surface"]).grid(row=0, column=0, rowspan=2, sticky="n", padx=(16, 10), pady=16)
        tk.Label(out, text="Reports are saved to", font=(FONT_SB, 10), fg=C["text"],
                 bg=C["surface"]).grid(row=0, column=1, sticky="w", pady=(14, 0))
        self.folder_label = tk.Label(out, text="", font=(FONT, 9), fg=C["muted"], bg=C["surface"],
                                     anchor="w", justify="left")
        self.folder_label.grid(row=1, column=1, sticky="ew", pady=(1, 14))
        links = tk.Frame(out, bg=C["surface"])
        links.grid(row=0, column=2, rowspan=2, sticky="e", padx=16)
        Link(links, "Open", self.open_reports_folder).pack(side="left", padx=(0, 12))
        Link(links, "Change", self.choose_folder).pack(side="left")

    # ---------- files ready? ----------
    def _missing(self):
        return [ZONE_BY_KEY[k][1] for k, *_ in ZONES if ZONE_BY_KEY[k][4] and not self.zones[k].paths]

    def on_zone_change(self):
        if not hasattr(self, "go"):
            return
        required = sum(1 for z in ZONES if z[4])
        have = required - len(self._missing())
        ready = have == required
        self.ready_pill.configure(text=f"{have} of {required} required files",
                                  fg=C["ok"] if ready else C["muted"],
                                  bg=C["ok_soft"] if ready else C["chip"])
        if not self.busy:
            self.go.style(bg=C["accent"] if ready else C["disabled"],
                          hover=C["accent_hover"] if ready else C["disabled"])

    def _compact(self, on):
        """While a message is showing, drop the small grey lines so everything fits on
        small laptop screens."""
        if on:
            self.hint_line.grid_remove()
        else:
            self.hint_line.grid()
        for z in getattr(self, "zones", {}).values():
            if on:
                z.meta.grid_remove()
                z.body.grid_configure(pady=(2, 10))
            else:
                z.meta.grid()
                z.body.grid_configure(pady=(2, 0))

    def clear_all(self):
        if self.busy:
            return
        for z in self.zones.values():
            z.paths = []
            z._paint()
        self.on_zone_change()

    # ---------- settings / folder ----------
    def reports_folder(self):
        f = self.settings.get("reports_folder")
        return f if f and os.path.isdir(f) else None

    def _refresh_folder_label(self):
        f = self.reports_folder()
        shown = f
        if f and len(f) > 44:                    # C:\Users\...\Final Reports
            parts = f.split(os.sep)
            shown = os.sep.join(parts[:2]) + os.sep + "…" + os.sep + os.sep.join(parts[-2:])
            if len(shown) > 52:
                shown = parts[0] + os.sep + "…" + os.sep + parts[-1]
        self.folder_label.configure(text=shown if f else "Not set yet - click Change",
                                    fg=C["muted"] if f else C["bad"])

    def choose_folder(self, first_time=False):
        if self.busy:
            return False
        if first_time:
            messagebox.showinfo("Where should reports go?",
                                "Which folder should final reports be saved to?\n\n"
                                "Pick it once - it will be remembered.\n"
                                "(You can change it later with Settings.)", parent=self.root)
        folder = filedialog.askdirectory(title="Which folder should final reports be saved to?",
                                         initialdir=self.reports_folder() or os.path.expanduser("~"),
                                         parent=self.root)
        if not folder:
            return False
        self.settings["reports_folder"] = os.path.abspath(folder)
        try:
            save_settings(self.settings)
        except OSError as e:
            messagebox.showerror("Settings", f"Could not save settings.json:\n{e}", parent=self.root)
        self._refresh_folder_label()
        return True

    def open_reports_folder(self):
        f = self.reports_folder()
        if f:
            os.startfile(f)
        else:
            self.choose_folder()

    def _startup_checks(self):
        if IMPORT_ERROR:
            self.banner.show("bad", "build_report.py not found",
                             f"Keep build_report.py in the same folder as this app, then reopen it:\n"
                             f"{APP_DIR}\n({IMPORT_ERROR})")
        elif not HAS_DND:
            self.banner.show("info", "Drag-and-drop is off",
                             "tkinterdnd2 is not installed - click a box to pick files. "
                             "Opening the app with run_app.bat installs it.")
        if not self.reports_folder():
            if self.settings.get("reports_folder"):
                messagebox.showwarning("Reports folder missing",
                                       f"The saved reports folder no longer exists:\n"
                                       f"{self.settings['reports_folder']}\n\nPlease choose it again.",
                                       parent=self.root)
            self.choose_folder(first_time=True)

    # ---------- several files dropped at once ----------
    def auto_sort(self, paths):
        unplaced = []
        for p in paths:
            low = os.path.basename(p).lower()
            target = None
            for key, _t, _h, exts, _r, words in ZONES:
                if low.endswith(exts) and any(w in low for w in words):
                    target = key
                    break
            if target is None:
                unplaced.append(os.path.basename(p))
            else:
                self.zones[target].set_paths([p])      # SIP box adds, others replace
        if unplaced:
            self.banner.show("warn", "Some files were not placed",
                             "Couldn't tell which box these belong to - drop them on the right box:\n"
                             + "\n".join(unplaced))

    # ---------- generate ----------
    def on_generate(self):
        if self.busy:
            return
        if br is None:
            self.banner.show("bad", "build_report.py not found",
                             f"Put build_report.py in this folder, then close and reopen the app:\n"
                             f"{APP_DIR}\n({IMPORT_ERROR})")
            return
        missing = self._missing()
        if missing:
            self.banner.show("warn", "Add the missing files first",
                             "Still empty:  " + "   ·   ".join(missing))
            return
        folder = self.reports_folder()
        if not folder:
            if not self.choose_folder(first_time=True):
                return
            folder = self.reports_folder()
        files = {k: list(z.paths) for k, z in self.zones.items()}
        self._set_busy(True)
        threading.Thread(target=self._worker, args=(files, folder), daemon=True).start()
        self.root.after(150, self._poll)

    def _worker(self, files, folder):
        try:
            res = generate_report(files, folder)
            res["move_problems"] = move_to_processed([p for ps in files.values() for p in ps])
            self.results.put(("ok", res, files))
        except Exception as e:
            if not isinstance(e, UserError):
                try:
                    with open(ERROR_LOG, "a", encoding="utf-8") as fh:
                        fh.write(f"\n==== {dt.datetime.now():%d-%m-%Y %H:%M:%S} ====\n")
                        fh.write(json.dumps(files, indent=1) + "\n")
                        fh.write(traceback.format_exc())
                except OSError:
                    pass
            self.results.put(("err", friendly_error(e), files))

    def _poll(self):
        try:
            kind, payload, files = self.results.get_nowait()
        except queue.Empty:
            self.root.after(150, self._poll)
            return
        self._set_busy(False)
        now = dt.datetime.now().strftime("%H:%M")
        if kind == "ok":
            self.done += 1
            path = payload["out_path"]
            details = payload["details"]
            if payload["move_problems"]:
                details += ("\n\nCould not move these to 'processed' (still in place):\n - "
                            + "\n - ".join(payload["move_problems"]))
            iid = self._add_log_row(payload["client"], now, "✔  Success", "ok")
            self.row_info[iid] = ("ok", path, details, payload["client"])
            for z in self.zones.values():
                z.paths = []
                z._paint()
            self.on_zone_change()
            notes = " · ".join(payload["notes"])
            actions = (("Open report", lambda: self._open(path)),
                       ("Open folder", self.open_reports_folder),
                       ("View details", lambda: self.show_details(payload["client"], details, path)))
            if payload.get("sip_errors"):
                self.banner.show("warn", f"Report saved - {payload['client']}  (check SIP files)",
                                 f"{os.path.basename(path)}{'  ·  ' + notes if notes else ''}\n"
                                 f"Some SIP files could not be used: "
                                 f"{payload['sip_errors'][0]}" +
                                 (f" (+{len(payload['sip_errors']) - 1} more)"
                                  if len(payload["sip_errors"]) > 1 else "") + " - see details.",
                                 actions)
            elif payload["move_problems"]:
                self.banner.show("warn", f"Report saved - {payload['client']}",
                                 f"{os.path.basename(path)}{'  ·  ' + notes if notes else ''}\n"
                                 f"Some input files could not be moved to 'processed' "
                                 f"(open in another program?) - see details.", actions)
            else:
                self.banner.show("ok", f"Report saved - {payload['client']}",
                                 f"{os.path.basename(path)}{'  ·  ' + notes if notes else ''}\n"
                                 f"Input files moved to 'processed'. Boxes cleared - ready for the "
                                 f"next client.", actions)
        else:
            self.failed += 1
            master = files.get("master") or []
            who = os.path.basename(master[0]) if master else "(no master file)"
            m = re.match(r"portfolio\s*valuation[\s_-]+(.+?)[\s_-]+\d", who, re.I)
            client = m.group(1).upper() if m else who
            iid = self._add_log_row(client, now, "✖  Failed", "bad")
            self.row_info[iid] = ("err", None, payload, client)
            lines = payload.strip().splitlines()
            short = "\n".join(lines[:3]) + ("\n..." if len(lines) > 3 else "")
            self.banner.show("bad", "Report not created - fix this and click Generate again",
                             short,
                             (("View details", lambda: self.show_details(client, payload)),))
        self._update_counter()

    def _set_busy(self, busy):
        self.busy = busy
        if busy:
            self.banner.hide()
            self.go.style(text="Generating report...", bg="#5B8FD9", cursor="watch")
            self.progress.grid()
            self.progress.start(12)
        else:
            self.progress.stop()
            self.progress.grid_remove()
            self.go.style(text="Generate Report", cursor="hand2")
            self.on_zone_change()

    # ---------- session list ----------
    def _add_log_row(self, client, when, status, tag):
        n = self.done + self.failed
        tags = (tag,) + (("odd",) if n % 2 == 0 else ())
        iid = self.log.insert("", 0, values=(f" {client}", f" {when}", f" {status}"), tags=tags)
        self.log.selection_set(iid)
        self.log.see(iid)
        self.empty.place_forget()
        return iid

    def _update_counter(self):
        self.stat["total"].configure(text=str(self.done + self.failed))
        self.stat["done"].configure(text=str(self.done))
        self.stat["failed"].configure(text=str(self.failed))

    def on_log_double_click(self, event):
        iid = self.log.identify_row(event.y)
        if iid not in self.row_info:
            return
        kind, path, details, client = self.row_info[iid]
        self.show_details(client, details, path if kind == "ok" else None)

    # ---------- details window ----------
    def _open(self, path):
        if path and os.path.exists(path):
            os.startfile(path)
        else:
            self.banner.show("warn", "Report file not found",
                             f"It may have been moved or renamed:\n{path}")

    def show_details(self, client, text, path=None):
        w = tk.Toplevel(self.root)
        w.title(f"Details - {client}")
        w.configure(bg=C["surface"])
        w.transient(self.root)
        w.minsize(560, 360)
        x = self.root.winfo_rootx() + 120
        y = self.root.winfo_rooty() + 80
        w.geometry(f"780x520+{x}+{y}")
        ok = path is not None
        head = tk.Frame(w, bg=C["header"])
        head.pack(fill="x")
        tk.Label(head, text=client, font=(FONT_SB, 13), fg="white",
                 bg=C["header"]).pack(anchor="w", padx=20, pady=(14, 0))
        tk.Label(head, text="Report created" if ok else "Report not created",
                 font=(FONT, 9), fg="#8FE3B5" if ok else "#FFB4AB",
                 bg=C["header"]).pack(anchor="w", padx=20, pady=(0, 14))
        bar = tk.Frame(w, bg=C["surface"])
        bar.pack(side="bottom", fill="x", padx=20, pady=14)
        IconButton(bar, "Close", w.destroy, bg=C["chip"], hover=C["border"], fg=C["text"],
                   font=(FONT_SB, 10)).pack(side="right")
        if ok:
            IconButton(bar, "Open report", lambda: self._open(path),
                       font=(FONT_SB, 10)).pack(side="right", padx=8)
        box = tk.Frame(w, bg=C["surface"])
        box.pack(fill="both", expand=True, padx=20, pady=(14, 0))
        txt = tk.Text(box, wrap="word", font=(FONT, 10), fg=C["text"], bg="#FAFBFD",
                      relief="flat", highlightthickness=1, highlightbackground=C["border"],
                      padx=12, pady=10)
        sb = ttk.Scrollbar(box, orient="vertical", command=txt.yview, style="Slim.Vertical.TScrollbar")
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        txt.insert("1.0", text)
        txt.configure(state="disabled")
        w.bind("<Escape>", lambda e: w.destroy())
        w.focus_set()

    def run(self):
        self.root.mainloop()


def main():
    if sys.platform == "win32":
        try:                                     # sharp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().run()


if __name__ == "__main__":
    main()
