# ============================================================
#  FlickerDrop
#  A quick-capture thought log that appends entries to an Excel
#  workbook (.xlsx) using openpyxl.
#
#  Features: multi-sheet tabs, recent-entries preview with
#  click-to-copy, undo, single-instance lock, dark/light theme.
#
#  Requirements:  customtkinter, openpyxl
# ============================================================



import sys
import os
import json
import queue
import socket
import threading
import time
import traceback
import atexit
import shutil
from datetime import datetime
from pathlib import Path
from tkinter import messagebox


CONFIG = {
    # ---------- EXCEL ENGINE ----------
    "engine": {
        "display_alerts":    False,
        "screen_updating":   False,
        "open_timeout_s":    30,
    },

    # ---------- FILE ----------
    "file": {
        # Overridden at startup: uses FlickerView's saved path when
        # present, otherwise falls back to Sample_Thoughts.xlsx next
        # to this script (see DYNAMIC FILE PATH RESOLUTION below).
        "path":               "Sample_Thoughts.xlsx",
        "template_sheet":     "Template",
        "default_sheet":      None,
        "index_column":       "A",
        "entry_column":       "B",
        "timestamp_column":   "C",
    },

    # ---------- KEY BINDINGS ----------
    "keys": {
        "submit":  ["<Return>", "<KP_Enter>"],
        "newline": ["<Control-Return>", "<Shift-Return>"],
    },

    # ---------- SAFE EDIT ----------
    "safe_edit": {
        "dry_run":              False,
        "verify_after_save":    True,
        "refuse_if_locked":     True,
        "check_file_lock":      True,
        "stale_lock_minutes":   5,
        "work_on_copy":         False,
        "copy_verify":          True,
        "copy_suffix":          "_working",
        "log_operations":       False,
        "log_file":             None,
    },

    # ---------- BACKUP ----------
    "backup": {
        "enabled":      False,
        "suffix":       "_backup",
        "keep_last_n":  20,
        "backup_dir":   None,
    },

    # ---------- INSERT BEHAVIOR ----------
    "insert": {
        "trim_whitespace":   True,
        "skip_empty":        True,
        "add_timestamp":     True,
        "timestamp_format":  "%Y-%m-%d %H:%M:%S",
        "mark_cells":        False,
        "mark_color":        "FFF9C4",
    },

    # ---------- TRANSFORM ----------
    "transform": {
        "case":         "none",     # none / lower / upper / title / capitalize
        "prefix":       "",
        "suffix":       "",
        "trim_chars":   None,
        "max_length":   0,
    },

    # ---------- SHEET TABS ----------
    "tabs": {
        "enabled":            True,
        "show_add_button":    True,
        "add_button_text":    "＋",
        "add_button_width":   34,
        "add_button_tooltip": "Create a new sheet from Template",
        "hide_template":      True,
        "tab_height":         25,
        "tab_width":          0,     # 0 = auto
        "tab_font_size":      14,
        "tab_corner_radius":  6,
        "selected_color":         ("#144870", "#0d3a6b"),   # light/dark tuple
        "selected_hover_color":   ("#144870", "#0d3a6b"),
        "unselected_color":       ("#3B8ED0", "#1F6AA5"),
        "unselected_hover_color": ("#5aa9e6", "#3B8ED0"),
        "new_sheet_dialog_title": "New Sheet",
        "new_sheet_dialog_text":  "Enter a name for the new sheet:",
    },

    # ---------- RECENT ENTRIES (shown below textbox) ----------
    "recent_entries": {
        "enabled":              True,
        "count":                20,       # read last N from the sheet
        "show_timestamp":       False,
        "show_index":           False,
        "separator":            "  ·  ",
        "max_chars_per_entry":  220,
        "font_size":            14,
        "label_text":           "Recent",
        "text_color":           "gray75",
        "label_color":          "gray50",
        "loading_text":         "Loading recent entries…",
        "empty_text":           "(no entries yet)",
        "refresh_button":       True,
        "wrap_width":           720,   # fallback only — labels auto-wrap to window width

        # --- Fixed-height scrollable container ---
        "height":               120,               # px — fixed section height
        "entry_gap":            4,                 # px gap between cards
        "card_bg":              ("gray92", "gray15"),  # light / dark tuple
        "card_radius":          6,
        "card_pad_x":           10,
        "card_pad_y":           6,
        "scrollbar_color":      ("gray72", "gray35"),
        "scrollbar_hover":      ("gray60", "gray45"),

        # --- hover / click-to-copy ---
        "card_hover_bg":        ("#dbe7ff", "#1f2b3d"),  # clearly visible tint
        "card_copied_bg":       ("#c8e6c9", "#1e4620"),  # brief green flash
        "card_click_copies":    True,                    # click → clipboard
        "card_cursor":          "hand2",                 # pointer cursor
    },

    # ---------- OPEN EXCEL BUTTON ----------
    "open_excel": {
        "enabled":                True,
        "button_text":            "📂  Open Excel",
        "activate_current_sheet": True,  # best-effort after open
    },

    # ---------- VIEW ENTRIES WINDOW ----------
    "view_entries": {
        "enabled":          True,
        "button_text":      "📋  Open FlickerView",
        "window_title":     "FlickerView",
        "window_size":      "820x600",
        "read_fresh":       True,
        "search_enabled":   True,
        "search_placeholder": "Search entries or timestamps…",
        "refresh_button":   True,
        "close_button":     True,
        "show_index":       True,
        "show_timestamp":   True,
        "font_size":        18,
        "empty_message":    "No entries yet.",
        "index_prefix":     "#",
        # --- where to open ---
        #   "internal" = built-in EntriesViewer popup (customtkinter)
        #   "external" = launch Entry_Viewer.pyw (PySide6 app)
        "target":           "external",
        "external_path":    None,   # None = sibling Entry_Viewer.pyw
        "external_wait":    False,  # block until external app closes?
    },

    # ---------- UNDO ----------
    "undo": {
        "enabled":        True,
        "show_button":    True,
        "button_text":    "↩  Undo",
        "stack_size":     50,
        "shortcut":       "<Control-z>",
    },

    # ---------- UI ----------
    "ui": {
        "appearance":            "dark",
        "color_theme":           "blue",
        "window_title":          "FlickerDrop",
        "window_size":           "780x790",
        "always_on_top":         False,

        "header_text":           "✨  FlickerDrop",
        "header_font_size":      24,
        "subtitle_text":         "Enter submits  •  Ctrl+Enter for a new line inside the entry",
        "subtitle_font_size":    12,

        "button_text":           "＋  Save Thought",
        "button_height":         42,
        "button_font_size":      16,

        "textbox_height":        320,
        "textbox_font_size":     28,
        "font_family":           None,

        "accent_color":          None,
        "show_status_bar":       True,
        "status_font_size":      14,
        "show_hint_bar":         True,
        "hint_font_size":        12,
        "hint_color":            "gray50",

        "autofocus_textbox":     True,
        "autofocus_delay_ms":    80,
        "select_all_on_focus":   False,

        "theme_toggle_shortcut": "<F2>",
    },

    # ---------- WINDOW MEMORY ----------
    "window": {
        "remember_position":    False,
        "position_file":        None,
        "center_on_screen":     True,
    },

    # ---------- NOTIFICATIONS ----------
    "notifications": {
        "use_popups":           False,
        "success_color":        "#6EE787",
        "error_color":          "#FF6B6B",
        "warning_color":        "orange",
        "info_color":           "gray60",
        "auto_clear_status_ms": 2500,   # ms before reverting to "Ready" (0 = off)
        "status_wraplength":    700,
        "popup_on_success":     False,
        "popup_on_error":       False,
    },

    # ---------- SOUND ----------
    "sound": {
        "enabled":          False,
        "success_sound":    "default",
        "error_sound":      "default",
        "warning_sound":    "default",
    },

    # ---------- HISTORY ----------
    "history": {
        "enabled":          False,
        "size":             10,
        "recall_shortcut":  "<Control-Up>",
    },

    # ---------- BEHAVIOR ----------
    "behavior": {
        "clear_after_add":           True,
        "refocus_after_add":         True,
        "cursor_to_start_after_add": True,
        "escape_closes":             True,
        "auto_close_enabled":        False,
        "auto_close_delay_ms":       1500,
        "auto_close_on_error":       False,
        "confirm_large_inserts":     False,
        "large_insert_threshold":    100,
    },

    # ---------- SINGLE INSTANCE ----------
    "instance": {
        "single_instance": True,
        "port":            51789,
        "focus_poll_ms":   300,
    },
}

# ============================================================
#  DYNAMIC FILE PATH RESOLUTION
# ============================================================

# Determine if running as .exe or .py
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys.executable).resolve().parent
else:
    _BASE_DIR = Path(__file__).resolve().parent

# 1. Check if FlickerView has saved a custom file path
_viewer_config = _BASE_DIR / ".entry_viewer_config.json"
_custom_path = None
if _viewer_config.exists():
    try:
        cfg_data = json.loads(_viewer_config.read_text(encoding="utf-8"))
        _custom_path = cfg_data.get("file_path")
    except Exception:
        pass

# 2. Set the path: User's choice -> Sample file fallback
if _custom_path and Path(_custom_path).exists():
    CONFIG["file"]["path"] = _custom_path
else:
    CONFIG["file"]["path"] = str(_BASE_DIR / "Sample_Thoughts.xlsx")

# ============================================================
#  IMPORTS & DEPENDENCY CHECK
# ============================================================

import customtkinter as ctk

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    msg = ("openpyxl is not installed.\n\n"
           "Install with:  pip install openpyxl")
    sys.stderr.write(f"\n[FATAL] {msg}\n")
    raise SystemExit(1)


# Reserved for a future Excel-COM / xlwings integration. Currently
# unused — openpyxl performs all reads and writes directly against the
# .xlsx file, so no serialization lock is needed today.
_EXCEL_LOCK = threading.RLock()


def _run_excel(func, path=None):
    """Future hook for Excel-COM operations; not implemented yet."""
    raise NotImplementedError("_run_excel is not implemented yet.")


# ============================================================
#  Helpers
# ============================================================
def _to_int(value, default=0):
    try:
        if value is None:
            return default
        return int(float(value))
    except (ValueError, TypeError):
        return default


def _is_file_locked_by_excel(path, retries=2, delay=0.15):
    """Return True only if there's an active write-lock (ignores stale ~$ files)."""
    p = Path(path)
    if not p.exists():
        return False
    for i in range(retries):
        try:
            with open(path, "r+b"):
                pass
            return False
        except PermissionError:
            if i < retries - 1:
                time.sleep(delay); continue
            return True
        except OSError as e:
            if getattr(e, "winerror", None) == 32:
                if i < retries - 1:
                    time.sleep(delay); continue
                return True
            return False
        except Exception:
            return False
    return False


def _play_sound(kind):
    s = CONFIG["sound"]
    if not s["enabled"]:
        return
    try:
        if os.name == "nt":
            import winsound
            name = s.get(f"{kind}_sound", "default")
            if name == "default":
                freq = {"success": 880, "error": 300, "warning": 600}.get(kind, 600)
                winsound.Beep(freq, 120)
            elif isinstance(name, str) and name.endswith(".wav"):
                winsound.PlaySound(name, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            sys.stdout.write("\a"); sys.stdout.flush()
    except Exception:
        pass


def _acquire_listener(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        if os.name == "nt":
            s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        else:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("127.0.0.1", port))
        s.listen(1)
        s.setblocking(False)
        return s
    except OSError:
        s.close()
        return None


def _ping_existing(port):
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.8) as c:
            c.sendall(b"focus")
        return True
    except Exception:
        return False


def _log(msg):
    if not CONFIG["safe_edit"].get("log_operations"):
        return
    lf = CONFIG["safe_edit"].get("log_file") or str(
        Path(__file__).with_name("thoughts_log.txt")
    )
    try:
        with open(lf, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat(timespec='seconds')}] {msg}\n")
    except Exception:
        pass


def _as_list(x):
    if isinstance(x, (list, tuple)):
        return list(x)
    return [x]


def _transform(value):
    t = CONFIG["transform"]
    s = str(value)

    case = t.get("case", "none")
    if case == "lower":
        s = s.lower()
    elif case == "upper":
        s = s.upper()
    elif case == "title":
        s = s.title()
    elif case == "capitalize":
        s = s.capitalize()

    tc = t.get("trim_chars")
    if tc:
        s = s.strip(tc)

    if t.get("prefix"):
        s = t["prefix"] + s
    if t.get("suffix"):
        s = s + t["suffix"]

    ml = _to_int(t.get("max_length", 0), 0)
    if ml > 0 and len(s) > ml:
        s = s[:ml]

    return s


# ============================================================
#  openpyxl Helpers
# ============================================================
def _find_last_row(ws, col_letter):
    col_idx = column_index_from_string(col_letter)
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=col_idx).value not in (None, ""):
            return row
    return 0

def _sheet_names(path=None):
    if path is None:
        path = CONFIG["file"]["path"]
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []


# ============================================================
#  Entries Viewer (read-only popup)
# ============================================================
class EntriesViewer(ctk.CTkToplevel):
    def __init__(self, parent, sheet_name):
        super().__init__(parent)
        cfg = CONFIG["view_entries"]
        self.title(cfg["window_title"])
        self.geometry(cfg["window_size"])
        self.transient(parent)
        self.sheet_name = sheet_name
        self._rows = []

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(10, 4))

        self.search_var = ctk.StringVar()
        if cfg["search_enabled"]:
            self.search_entry = ctk.CTkEntry(
                top, textvariable=self.search_var,
                placeholder_text=cfg["search_placeholder"],
            )
            self.search_entry.pack(side="left", fill="x", expand=True)
            self.search_entry.bind("<KeyRelease>", lambda e: self.render())

        if cfg["refresh_button"]:
            ctk.CTkButton(top, text="↻", width=40,
                          command=self.load).pack(side="left", padx=(6, 0))

        if cfg["close_button"]:
            ctk.CTkButton(top, text="Close", width=70,
                          command=self.destroy).pack(side="left", padx=(6, 0))

        self.sheet_label = ctk.CTkLabel(
            self, text=f"Sheet: {sheet_name}", anchor="w",
        )
        self.sheet_label.pack(fill="x", padx=12)

        self.textbox = ctk.CTkTextbox(
            self, font=ctk.CTkFont(size=cfg["font_size"])
        )
        self.textbox.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.textbox.configure(state="disabled")

        self.after(60, self.load)

    def load(self):
        fcfg = CONFIG["file"]
        path = fcfg["path"]
        self._rows = []

        if not Path(path).exists():
            self.render(); return

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            if self.sheet_name not in wb.sheetnames:
                wb.close()
                self.render(); return
            
            ws = wb[self.sheet_name]
            last = _find_last_row(ws, fcfg["index_column"])
            if last <= 0:
                wb.close()
                self.render(); return
                
            out = []
            c_idx = column_index_from_string(fcfg["index_column"])
            c_ent = column_index_from_string(fcfg["entry_column"])
            c_ts = column_index_from_string(fcfg["timestamp_column"])
            
            for r in range(1, last + 1):
                vals = (
                    ws.cell(row=r, column=c_idx).value,
                    ws.cell(row=r, column=c_ent).value,
                    ws.cell(row=r, column=c_ts).value
                )
                out.append(vals)
            wb.close()
            self._rows = out
        except Exception as e:
            _log(f"viewer load failed: {e}")
            self._rows = []

        self.render()

    def render(self):
        cfg = CONFIG["view_entries"]
        q = self.search_var.get().strip().lower() if cfg["search_enabled"] else ""

        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")

        count = 0
        for idx, entry, ts in self._rows:
            entry_s = "" if entry is None else str(entry)
            ts_s = "" if ts is None else str(ts)
            idx_s = "" if idx is None else str(idx)

            if q and (q not in entry_s.lower() and q not in ts_s.lower()):
                continue

            count += 1
            header_bits = []
            if cfg["show_index"] and idx_s:
                header_bits.append(f"{cfg['index_prefix']}{idx_s}")
            if cfg["show_timestamp"] and ts_s:
                header_bits.append(ts_s)
            if header_bits:
                self.textbox.insert("end", "   ".join(header_bits) + "\n")
            self.textbox.insert("end", entry_s + "\n\n")

        if count == 0:
            self.textbox.insert("end", cfg["empty_message"])

        self.textbox.configure(state="disabled")

    def refresh(self):
        self.load()


# ============================================================
#  Recent Entries strip (below the textbox)
# ============================================================
class RecentEntriesWidget(ctk.CTkFrame):
    """Fixed-height, scrollable preview of the last N entries.
    Labels auto-wrap to the current widget width — no more touching
    wrap_width on every resize. The outer widget is pinned via
    pack_propagate(False) so content overflow can never push the parent
    buttons off-screen."""

    def __init__(self, parent, on_refresh=None, on_copy=None):
        super().__init__(parent, fg_color="transparent")
        cfg = CONFIG["recent_entries"]
        self.cfg = cfg
        self._on_copy = on_copy

        # Track labels so we can re-wrap them on window resize
        self._labels = []
        self._last_width = 0
        self._resize_after_id = None
        self.bind("<Configure>", self._on_configure)

        HEADER_H = 26
        GAP = 4
        self._total_h = HEADER_H + GAP + cfg["height"]

        # --- Header row (fixed height, cannot grow) ---
        header = ctk.CTkFrame(self, fg_color="transparent", height=HEADER_H)
        header.pack(fill="x", pady=(0, GAP))
        header.pack_propagate(False)

        ctk.CTkLabel(
            header, text=cfg["label_text"],
            font=ctk.CTkFont(size=cfg["font_size"], weight="bold"),
            text_color=cfg["label_color"], anchor="w",
        ).pack(side="left")

        if cfg.get("refresh_button") and on_refresh:
            ctk.CTkButton(
                header, text="↻", width=26, height=20,
                font=ctk.CTkFont(size=cfg["font_size"] - 1),
                fg_color="transparent",
                hover_color=("gray70", "gray40"),   # clearly visible hover
                text_color=cfg["label_color"],
                command=on_refresh,
            ).pack(side="right")

        # --- Scrollable area, pinned to cfg["height"] ---
        self.scroll = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            height=cfg["height"],
            corner_radius=0,
            scrollbar_button_color=cfg["scrollbar_color"],
            scrollbar_button_hover_color=cfg["scrollbar_hover"],
        )
        self.scroll.pack(fill="both", expand=True)

        # PIN the outer widget: content below cannot grow it.
        self.configure(height=self._total_h)
        self.pack_propagate(False)

        self.set_loading()

    # ---------- responsive wrapping ----------
    def _calc_wrap_width(self):
        """Compute the label wrap width from the current widget width."""
        w = self.winfo_width()
        if w <= 100:  # not laid out yet — use config fallback
            return self.cfg.get("wrap_width", 720)
        # card_pad_x per side + scrollbar gutter buffer
        return max(180, w - self.cfg["card_pad_x"] * 2 - 30)

    def _on_configure(self, event):
        # Only react to width changes; ignore height-only events
        w = event.width
        if w == self._last_width:
            return
        self._last_width = w
        if self._resize_after_id is not None:
            try:
                self.after_cancel(self._resize_after_id)
            except Exception:
                pass
        # Debounce 80 ms so rapid resize doesn't thrash label reconfigure
        self._resize_after_id = self.after(80, self._update_wrap_widths)

    def _update_wrap_widths(self):
        self._resize_after_id = None
        ww = self._calc_wrap_width()
        for lbl in self._labels:
            try:
                lbl.configure(wraplength=ww)
            except Exception:
                pass

    # ---------- internal helpers ----------
    def _clear(self):
        self._labels.clear()
        for w in self.scroll.winfo_children():
            try: w.destroy()
            except Exception: pass

    def _make_row(self, text, color, raw_text=None):
        cfg = self.cfg
        row = ctk.CTkFrame(
            self.scroll,
            fg_color=cfg["card_bg"],
            corner_radius=cfg["card_radius"],
        )
        row.pack(fill="x", pady=(0, cfg["entry_gap"]))

        ww = self._calc_wrap_width()
        lbl = ctk.CTkLabel(
            row, text=text,
            font=ctk.CTkFont(size=cfg["font_size"]),
            text_color=color, anchor="w", justify="left",
            wraplength=ww,
        )
        lbl.pack(fill="x", padx=cfg["card_pad_x"], pady=cfg["card_pad_y"])
        self._labels.append(lbl)

        # Hover + click-to-copy only for real entries (raw_text provided)
        if raw_text is not None and cfg.get("card_click_copies", True):
            self._attach_row_behavior(row, lbl, raw_text)
        return row

    # ---------- hover + click ----------
    def _bind_recursive(self, widget, sequence, callback):
        """Bind an event to a widget and every descendant so clicks /
        hovers anywhere on the card trigger the same handler. CTkLabel
        wraps its text in a tkinter.Label — without recursive binding,
        clicks on the text never reach the parent frame."""
        try:
            widget.bind(sequence, callback, add="+")
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._bind_recursive(child, sequence, callback)
        except Exception:
            pass

    def _set_cursor_recursive(self, widget, cursor):
        try:
            widget.configure(cursor=cursor)
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._set_cursor_recursive(child, cursor)
        except Exception:
            pass

    def _attach_row_behavior(self, row, lbl, raw_text):
        cfg = self.cfg
        base = cfg["card_bg"]
        hover = cfg.get("card_hover_bg", ("gray80", "gray30"))
        cursor = cfg.get("card_cursor", "hand2")

        state = {"hovering": False}

        def on_enter(e):
            state["hovering"] = True
            try:
                row.configure(fg_color=hover)
            except Exception:
                pass

        def on_leave(e):
            state["hovering"] = False
            try:
                row.configure(fg_color=base)
            except Exception:
                pass

        def on_click(e):
            self._copy_entry(raw_text, row, base, hover, state)
            return "break"

        self._set_cursor_recursive(row, cursor)

        for seq, cb in (("<Enter>", on_enter),
                        ("<Leave>", on_leave),
                        ("<Button-1>", on_click)):
            self._bind_recursive(row, seq, cb)

    def _copy_entry(self, raw_text, row, base, hover, state):
        # 1. Copy to clipboard
        try:
            self.clipboard_clear()
            self.clipboard_append(raw_text)
            self.update_idletasks()
        except Exception:
            return

        # 2. Brief green flash for immediate feedback
        copied = self.cfg.get("card_copied_bg", ("#c8e6c9", "#1e4620"))

        def restore():
            try:
                row.configure(fg_color=hover if state["hovering"] else base)
            except Exception:
                pass

        try:
            row.configure(fg_color=copied)
            self.after(180, restore)
        except Exception:
            pass

        # 3. Notify the host app (status-bar toast)
        if self._on_copy:
            try:
                self._on_copy(raw_text)
            except Exception:
                pass

    # ---------- public API ----------
    def set_loading(self):
        self._clear()
        self._make_row(self.cfg["loading_text"], self.cfg["text_color"])

    def set_error(self, message):
        self._clear()
        self._make_row(
            f"⚠  {message}", CONFIG["notifications"]["error_color"]
        )

    def set_entries(self, rows):
        """rows = list of (idx, text, ts) tuples, oldest-first (sheet order)."""
        self._clear()
        if not rows:
            self._make_row(self.cfg["empty_text"], self.cfg["text_color"])
            return

        cfg = self.cfg
        rows = rows[-cfg["count"]:][::-1]   # newest first

        for idx, text, ts in rows:
            # Preserve the original text (with newlines) for clipboard copy
            raw_text = "" if text is None else str(text)

            # Build the truncated single-line version for display
            body = raw_text.replace("\n", " ").replace("\r", " ").strip()
            if len(body) > cfg["max_chars_per_entry"]:
                body = body[: cfg["max_chars_per_entry"] - 1] + "…"

            bits = []
            if cfg.get("show_index") and idx not in (None, ""):
                bits.append(f"#{idx}")
            if cfg.get("show_timestamp") and ts not in (None, ""):
                ts_s = str(ts)
                if " " in ts_s:
                    ts_s = ts_s.split(" ", 1)[1]   # time only
                bits.append(ts_s)
            bits.append(body)
            line = cfg["separator"].join(b for b in bits if b)
            self._make_row(line, cfg["text_color"], raw_text=raw_text)

        # Re-pin the scroll height
        try:
            self.scroll.configure(height=self.cfg["height"])
        except Exception:
            pass


# ============================================================
#  APP
# ============================================================
class ThoughtsApp(ctk.CTk):
    def __init__(self, listener_socket=None):
        super().__init__()

        ui = CONFIG["ui"]
        ctk.set_appearance_mode(ui["appearance"])
        ctk.set_default_color_theme(ui["color_theme"])

        # Show the filename in the title so the user knows where it's saving
        f_name = Path(CONFIG["file"]["path"]).name
        self.title(f"{ui['window_title']} — {f_name}")
        self._apply_window_geometry()
        try:
            self.iconbitmap("FlickerDrop.ico")
        except Exception:
            pass
        self.resizable(False, False)
        if ui["always_on_top"]:
            self.attributes("-topmost", True)

        # --- Header ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(pady=(18, 2))
        ctk.CTkLabel(
            header_frame, text=ui["header_text"],
            font=ctk.CTkFont(size=ui["header_font_size"], weight="bold"),
        ).pack(side="left")

        ctk.CTkLabel(
            self, text=ui["subtitle_text"], text_color="gray65",
            font=ctk.CTkFont(size=ui["subtitle_font_size"]),
        ).pack(pady=(0, 8))

        # Subtle divider under header for visual hierarchy
        ctk.CTkFrame(
            self, height=1, fg_color=("gray82", "gray24"),
        ).pack(fill="x", padx=32, pady=(0, 10))

        # --- Sheet tabs ---
        self.tab_frame = ctk.CTkFrame(self, fg_color="transparent")
        if CONFIG["tabs"]["enabled"]:
            self.tab_frame.pack(fill="x", padx=20, pady=(0, 8))
        self.tab_buttons = {}
        self.sheet_names = []
        self.current_sheet = None

        # --- Textbox ---
        box_font = ctk.CTkFont(size=ui["textbox_font_size"])
        if ui["font_family"]:
            box_font = ctk.CTkFont(
                family=ui["font_family"], size=ui["textbox_font_size"]
            )
        # Card-style wrapper for a more modern, framed look
        textbox_h = ui["textbox_height"]
        self.textbox_card = ctk.CTkFrame(
            self, corner_radius=12,
            fg_color=("gray94", "gray13"),
            border_width=1,
            border_color=("gray80", "gray28"),
        )
        self.textbox_card.pack(padx=28, pady=(0, 4), fill="x")
        self.textbox = ctk.CTkTextbox(
            self.textbox_card, width=688, height=textbox_h,
            font=box_font, corner_radius=8,
            border_width=0, fg_color=("gray100", "gray10"),
        )
        self.textbox.pack(padx=6, pady=6, fill="both", expand=True)

        # --- Recent entries strip ---
        if CONFIG["recent_entries"]["enabled"]:
            self.recent_widget = RecentEntriesWidget(
                self, on_refresh=self._reload_recent,
                on_copy=self._on_recent_copy,
            )
            # expand=False + fill="x" → horizontal fill, vertical fixed
            self.recent_widget.pack(fill="x", expand=False,
                                    padx=30, pady=(2, 6))
        else:
            self.recent_widget = None

        # --- Hint bar ---
        if ui["show_hint_bar"]:
            self.hint = ctk.CTkLabel(
                self, text="", text_color=ui["hint_color"],
                font=ctk.CTkFont(size=ui["hint_font_size"]),
            )
            self.hint.pack(pady=(0, 4))
        else:
            self.hint = None

        # --- Add button ---
        btn_kwargs = dict(
            text=ui["button_text"],
            height=ui["button_height"],
            font=ctk.CTkFont(size=ui["button_font_size"], weight="bold"),
            corner_radius=10,
            command=self.add_values,
        )
        if ui["accent_color"]:
            btn_kwargs["fg_color"] = ui["accent_color"]
            btn_kwargs["text_color"] = "black"
        self.add_btn = ctk.CTkButton(self, **btn_kwargs)
        self.add_btn.pack(pady=(8, 6), fill="x", padx=30)

        # --- Action row ---
        action_row = ctk.CTkFrame(self, fg_color="transparent")
        action_row.pack(fill="x", padx=30, pady=(0, 6))

        # Modern ghost-style secondary actions
        ghost_kwargs = dict(
            height=34, corner_radius=8,
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray30"),
            text_color=("gray15", "gray85"),
            hover_color=("gray88", "gray22"),
        )

        if CONFIG["open_excel"]["enabled"]:
            ctk.CTkButton(
                action_row, text=CONFIG["open_excel"]["button_text"],
                command=self.open_excel, **ghost_kwargs,
            ).pack(side="left", expand=True, fill="x", padx=(0, 4))

        if CONFIG["view_entries"]["enabled"]:
            ctk.CTkButton(
                action_row, text=CONFIG["view_entries"]["button_text"],
                command=self.view_entries, **ghost_kwargs,
            ).pack(side="left", expand=True, fill="x", padx=4)

        if CONFIG["undo"]["enabled"] and CONFIG["undo"]["show_button"]:
            ctk.CTkButton(
                action_row, text=CONFIG["undo"]["button_text"],
                command=self.undo_last, **ghost_kwargs,
            ).pack(side="left", expand=True, fill="x", padx=(4, 0))

        # --- Status ---
        if ui["show_status_bar"]:
            self.status = ctk.CTkLabel(
                self, text="Ready",
                text_color=CONFIG["notifications"]["info_color"],
                font=ctk.CTkFont(size=ui["status_font_size"]),
                wraplength=CONFIG["notifications"]["status_wraplength"],
                justify="center",
            )
            self.status.pack(pady=(4, 12))
        else:
            self.status = None

        # --- State ---
        self._undo_stack = []
        self._history = []
        self._viewer = None
        self._busy = False
        # Recent-entries reload control:
        #   _recent_load_gen increments on every load request; results with
        #   an older gen are discarded (fixes stale data when sheets switch
        #   mid-load).
        #   _recent_debounce_id holds the pending after() id so rapid sheet
        #   switches coalesce into one load.
        self._recent_load_gen = 0
        self._recent_debounce_id = None
        self._status_clear_id = None    # pending after() to reset status to "Ready"

        # --- Key bindings ---
        for key in _as_list(CONFIG["keys"]["submit"]):
            try:
                self.textbox.bind(key, self._on_submit_key)
            except Exception:
                pass
        for key in _as_list(CONFIG["keys"]["newline"]):
            try:
                self.textbox.bind(key, self._on_newline_key)
            except Exception:
                pass

        if CONFIG["behavior"]["escape_closes"]:
            self.bind("<Escape>", lambda e: self.destroy())
        if CONFIG["undo"]["enabled"] and CONFIG["undo"]["shortcut"]:
            self.bind(CONFIG["undo"]["shortcut"], lambda e: self.undo_last())
        if ui["theme_toggle_shortcut"]:
            self.bind(ui["theme_toggle_shortcut"], lambda e: self.toggle_theme())
        if CONFIG["history"]["enabled"]:
            self.bind(CONFIG["history"]["recall_shortcut"],
                      lambda e: self._recall_history())

        if ui["autofocus_textbox"]:
            self.after(ui["autofocus_delay_ms"], self._initial_focus)

        self.textbox.bind("<<Modified>>", self._on_text_modified)

        # --- Single instance ---
        self._focus_queue = queue.Queue()
        if listener_socket is not None:
            self._start_listener(listener_socket)
            self.after(CONFIG["instance"]["focus_poll_ms"], self._poll_focus_queue)

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- Load tabs ---
        if CONFIG["tabs"]["enabled"]:
            self._load_sheets_async()

    # ---------- lifecycle ----------
    def _on_close(self):
        self.destroy()

    # ---------- key handlers ----------
    def _on_submit_key(self, event=None):
        self.add_values()
        return "break"

    def _on_newline_key(self, event=None):
        try:
            self.textbox.insert("insert", "\n")
        except Exception:
            pass
        return "break"

    # ---------- window geometry ----------
    def _apply_window_geometry(self):
        ui = CONFIG["ui"]
        w_cfg = CONFIG["window"]
        self.geometry(ui["window_size"])

        if not w_cfg["remember_position"]:
            if w_cfg["center_on_screen"]:
                self.update_idletasks()
                try:
                    w, h = ui["window_size"].split("x")
                    w, h = _to_int(w, 780), _to_int(h, 660)
                except Exception:
                    w, h = 780, 660
                x = (self.winfo_screenwidth() - w) // 2
                y = (self.winfo_screenheight() - h) // 3
                self.geometry(f"+{x}+{y}")
            return

        pf = w_cfg["position_file"]
        if pf is None:
            base = os.getenv("LOCALAPPDATA") or str(Path.home())
            pf = str(Path(base) / "ThoughtsBuilder" / "pos.json")
        self._pos_file = pf
        try:
            with open(pf, "r", encoding="utf-8") as f:
                pos = json.load(f)
            self.geometry(f"+{_to_int(pos.get('x'))}+{_to_int(pos.get('y'))}")
        except Exception:
            pass
        self.bind("<Configure>", self._save_position)

    def _save_position(self, event=None):
        if event is not None and event.widget is not self:
            return
        try:
            Path(self._pos_file).parent.mkdir(parents=True, exist_ok=True)
            with open(self._pos_file, "w", encoding="utf-8") as f:
                json.dump({"x": self.winfo_x(), "y": self.winfo_y()}, f)
        except Exception:
            pass

    # ---------- theme ----------
    def toggle_theme(self):
        ctk.set_appearance_mode(
            "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        )

    # ---------- focus ----------
    def _initial_focus(self):
        try:
            self.deiconify(); self.lift(); self.focus_force()
        except Exception:
            pass
        try:
            self.textbox.focus_force()
            if CONFIG["ui"]["select_all_on_focus"]:
                self.textbox.tag_add("sel", "1.0", "end-1c")
            self.textbox.mark_set("insert", "end-1c")
            self.textbox.see("end")
        except Exception:
            pass

    # ---------- single instance ----------
    def _start_listener(self, sock):
        def loop():
            while True:
                try:
                    conn, _ = sock.accept()
                except BlockingIOError:
                    threading.Event().wait(0.15); continue
                except OSError:
                    break
                try:
                    with conn:
                        conn.settimeout(0.5)
                        try: conn.recv(32)
                        except Exception: pass
                        self._focus_queue.put(True)
                except Exception:
                    pass
        threading.Thread(target=loop, daemon=True).start()

    def _poll_focus_queue(self):
        try:
            while True:
                self._focus_queue.get_nowait()
                self._bring_to_front()
        except queue.Empty:
            pass
        self.after(CONFIG["instance"]["focus_poll_ms"], self._poll_focus_queue)

    def _bring_to_front(self):
        try:
            self.deiconify(); self.lift()
            self.attributes("-topmost", True)
            self.after(180, lambda: self.attributes(
                "-topmost", CONFIG["ui"]["always_on_top"]
            ))
            self.focus_force()
            if CONFIG["ui"]["autofocus_textbox"]:
                self.after(50, lambda: self.textbox.focus_force())
        except Exception:
            pass

    # ---------- notifications ----------
    def _notify(self, text, kind="info"):
        n = CONFIG["notifications"]
        color = {
            "success": n["success_color"],
            "error":   n["error_color"],
            "warning": n["warning_color"],
            "info":    n["info_color"],
        }.get(kind, n["info_color"])

        if self.status is not None:
            self.status.configure(text=text, text_color=color)

        if n["use_popups"]:
            if kind == "error" and n["popup_on_error"]:
                messagebox.showerror("Error", text)
            elif kind == "success" and n["popup_on_success"]:
                messagebox.showinfo("Success", text)

        if kind in ("success", "error", "warning"):
            _play_sound(kind)

        # Cancel any pending auto-clear so an old timer can't wipe the
        # message we're about to schedule (rapid copy / add / switch).
        if self._status_clear_id is not None:
            try:
                self.after_cancel(self._status_clear_id)
            except Exception:
                pass
            self._status_clear_id = None

        delay = n.get("auto_clear_status_ms", 0)
        # Skip auto-clear for errors (user needs to read them), for
        # "Ready" itself (would infinite-loop), and when disabled.
        if delay > 0 and kind != "error" and text != "Ready":
            self._status_clear_id = self.after(delay, self._reset_status)

    def _reset_status(self):
        self._status_clear_id = None
        if self.status is not None:
            try:
                self.status.configure(
                    text="Ready",
                    text_color=CONFIG["notifications"]["info_color"],
                )
            except Exception:
                pass

    def _maybe_auto_close(self, error=False):
        b = CONFIG["behavior"]
        if not b["auto_close_enabled"]:
            return
        if error and not b["auto_close_on_error"]:
            return
        self.after(b["auto_close_delay_ms"], self._on_close)

    # ---------- hint bar ----------
    def _on_text_modified(self, event=None):
        try:
            if self.textbox.edit_modified():
                self.textbox.edit_modified(False)
        except Exception:
            pass
        self._update_hint()

    def _update_hint(self):
        if self.hint is None:
            return
        raw = self.textbox.get("1.0", "end")
        if raw.endswith("\n"):
            raw = raw[:-1]
        if not raw.strip():
            self.hint.configure(text="")
            return
        n_chars = len(raw)
        n_lines = raw.count("\n") + 1
        self.hint.configure(
            text=f"{n_lines} line{'s' if n_lines != 1 else ''}  •  {n_chars} chars"
        )

    # ---------- history ----------
    def _recall_history(self):
        if not CONFIG["history"]["enabled"]:
            return
        if not self._history:
            self._notify("↺  No history yet.", "info"); return
        last = self._history[-1]
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", last)
        self._notify("↺  Recalled last entry.", "info")

    # ---------- recent entries ----------
    def _reload_recent(self):
        """Refresh button handler — force immediate reload."""
        if self.recent_widget is not None:
            self.recent_widget.set_loading()
            self._load_recent_async()

    def _on_recent_copy(self, text):
        """Called by RecentEntriesWidget after copying to clipboard."""
        preview = text.replace("\n", " ").strip()
        if len(preview) > 60:
            preview = preview[:57] + "…"
        self._notify(f"📋  Copied:  {preview}", "success")

    def _schedule_recent_load(self, delay_ms=120):
        """Debounced trigger. Rapid sheet switches fire one load, not N."""
        if self.recent_widget is None:
            return
        if self._recent_debounce_id is not None:
            try:
                self.after_cancel(self._recent_debounce_id)
            except Exception:
                pass
        self._recent_debounce_id = self.after(
            delay_ms, self._fire_recent_load
        )

    def _fire_recent_load(self):
        self._recent_debounce_id = None
        self._load_recent_async()

    def _load_recent_async(self):
        if self.recent_widget is None:
            return
        if not self.current_sheet:
            self.recent_widget.set_entries([])
            return

        self._recent_load_gen += 1
        my_gen = self._recent_load_gen
        sheet_name = self.current_sheet

        def worker():
            try:
                rows = self._read_recent(sheet_name)
                self.after(
                    0,
                    lambda r=rows, g=my_gen: self._on_recent_loaded(r, None, g),
                )
            except Exception as e:
                err = str(e)
                self.after(
                    0,
                    lambda err=err, g=my_gen: self._on_recent_loaded(None, err, g),
                )

        threading.Thread(target=worker, daemon=True).start()

    def _read_recent(self, sheet_name):
        fcfg = CONFIG["file"]
        path = fcfg["path"]
        if not Path(path).exists():
            return []
        count = _to_int(CONFIG["recent_entries"]["count"], 3)
        if count <= 0:
            return []

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            if sheet_name not in wb.sheetnames:
                wb.close(); return []
            ws = wb[sheet_name]
            last = _find_last_row(ws, fcfg["index_column"])
            if last <= 0:
                wb.close(); return []
            
            start = max(1, last - count + 1)
            c_idx = column_index_from_string(fcfg["index_column"])
            c_ent = column_index_from_string(fcfg["entry_column"])
            c_ts = column_index_from_string(fcfg["timestamp_column"])
            
            out = []
            for r in range(start, last + 1):
                out.append((
                    ws.cell(row=r, column=c_idx).value,
                    ws.cell(row=r, column=c_ent).value,
                    ws.cell(row=r, column=c_ts).value
                ))
            wb.close()
            return out
        except Exception:
            return []

    def _on_recent_loaded(self, rows, error, gen=None):
        if self.recent_widget is None:
            return
        # Discard stale results — user has switched sheets since we fired.
        if gen is not None and gen != self._recent_load_gen:
            return
        if error:
            msg = str(error)
            low = msg.lower()
            if (
                "remote procedure call" in low
                or "0x800706be" in low
                or "call was rejected" in low
            ):
                # Transient COM error — retry after a short delay.
                self._schedule_recent_load(delay_ms=400)
                return
            self.recent_widget.set_error(msg)
        else:
            self.recent_widget.set_entries(rows or [])

    # ---------- sheets / tabs ----------
    def _load_sheets_async(self):
        def worker():
            try:
                names = _sheet_names()
                self.after(0, lambda n=names: self._on_sheets_loaded(n, None))
            except Exception as e:
                err = f"{e}"
                self.after(0, lambda err=err: self._on_sheets_loaded(None, err))

        threading.Thread(target=worker, daemon=True).start()

    def _on_sheets_loaded(self, names, error):
        if error:
            self._notify(f"❌  Could not load sheets: {error}", "error")
            return
        self.sheet_names = list(names)
        self._build_tab_strip()

        template = CONFIG["file"]["template_sheet"]
        hide_tpl = CONFIG["tabs"]["hide_template"]
        default_cfg = CONFIG["file"]["default_sheet"]

        choice = None
        if default_cfg and default_cfg in self.sheet_names:
            choice = default_cfg
        else:
            for n in self.sheet_names:
                if hide_tpl and n == template:
                    continue
                choice = n
                break
            if choice is None and self.sheet_names:
                choice = self.sheet_names[0]

        self.current_sheet = choice
        self._refresh_tab_styles()
        if CONFIG["recent_entries"]["enabled"]:
            self._schedule_recent_load()

    def _build_tab_strip(self):
        for w in self.tab_frame.winfo_children():
            w.destroy()
        self.tab_buttons.clear()

        tabs = CONFIG["tabs"]
        template = CONFIG["file"]["template_sheet"]

        for name in self.sheet_names:
            if tabs["hide_template"] and name == template:
                continue
            btn_kwargs = dict(
                text=name,
                height=tabs["tab_height"],
                corner_radius=tabs["tab_corner_radius"],
                font=ctk.CTkFont(size=tabs["tab_font_size"]),
                command=lambda n=name: self._select_sheet(n),
            )
            if tabs["tab_width"] > 0:
                btn_kwargs["width"] = tabs["tab_width"]
            btn = ctk.CTkButton(self.tab_frame, **btn_kwargs)
            btn.pack(side="left", padx=2)
            self.tab_buttons[name] = btn

        if tabs["show_add_button"]:
            add_btn = ctk.CTkButton(
                self.tab_frame,
                text=tabs["add_button_text"],
                width=tabs["add_button_width"],
                height=tabs["tab_height"],
                corner_radius=tabs["tab_corner_radius"],
                font=ctk.CTkFont(size=tabs["tab_font_size"], weight="bold"),
                command=self.add_sheet,
            )
            add_btn.pack(side="left", padx=(8, 2))

    def _refresh_tab_styles(self):
        tabs = CONFIG["tabs"]
        sel = tabs.get("selected_color")
        unsel = tabs.get("unselected_color")
        sel_hover = tabs.get("selected_hover_color")
        unsel_hover = tabs.get("unselected_hover_color")

        for name, btn in self.tab_buttons.items():
            try:
                if name == self.current_sheet:
                    kwargs = {}
                    if sel:
                        kwargs["fg_color"] = sel
                        kwargs["hover_color"] = sel_hover or sel
                    if kwargs:
                        btn.configure(**kwargs)
                else:
                    kwargs = {}
                    if unsel:
                        kwargs["fg_color"] = unsel
                        kwargs["hover_color"] = unsel_hover or unsel
                    if kwargs:
                        btn.configure(**kwargs)
            except Exception:
                pass

    def _select_sheet(self, name):
        self.current_sheet = name
        self._refresh_tab_styles()
        self._notify(f"📑  Sheet: {name}", "info")
        if CONFIG["recent_entries"]["enabled"]:
            self._schedule_recent_load()

    def add_sheet(self):
        if self._busy:
            self._notify("⏳  Busy — please wait…", "info"); return

        tabs = CONFIG["tabs"]
        dlg = ctk.CTkInputDialog(
            text=tabs["new_sheet_dialog_text"],
            title=tabs["new_sheet_dialog_title"],
        )
        name = dlg.get_input()
        if not name:
            return
        name = name.strip()
        if not name:
            return

        invalid = set(r'\/?*[]:')
        if any(ch in invalid for ch in name):
            self._notify("❌  Sheet name contains invalid characters.", "error"); return
        if len(name) > 31:
            self._notify("❌  Sheet name too long (max 31 characters).", "error"); return

        path = CONFIG["file"]["path"]
        template = CONFIG["file"]["template_sheet"]

        self._busy = True
        self._notify(f"⏳  Creating sheet '{name}'…", "info")
        self.update_idletasks()
        try:
            wb = openpyxl.load_workbook(path)
            if name in wb.sheetnames:
                raise ValueError(f"Sheet '{name}' already exists.")
            
            ws = wb.create_sheet(name)
            
            # Copy header from Template if it exists
            if template in wb.sheetnames:
                tpl = wb[template]
                for col in range(1, tpl.max_column + 1):
                    ws.cell(row=1, column=col).value = tpl.cell(row=1, column=col).value
            
            wb.save(path)
            wb.close()
            self._notify(f"✅  Sheet '{name}' created.", "success")
            self._load_sheets_async()
            self.after(500, lambda n=name: self._select_sheet(n))
        except Exception as e:
            self._notify(f"❌  {e}", "error")
        finally:
            self._busy = False

    # ---------- open Excel ----------
    def open_excel(self):
        path = CONFIG["file"]["path"]
        if not Path(path).exists():
            self._notify(f"❌  File not found: {path}", "error"); return
        try:
            # Give the current sheet a chance to be selected
            if (CONFIG["open_excel"]["activate_current_sheet"]
                    and self.current_sheet):
                try:
                    def do(wb):
                        names = [s.name for s in wb.sheets]
                        if self.current_sheet in names:
                            wb.sheets[self.current_sheet].activate()
                    # Sheet activation is a future hook — not currently
                    # wired up. The workbook is opened via the OS handler.
                except Exception:
                    pass

            if os.name == "nt":
                os.startfile(path)   # opens with default app (Excel)
            else:
                import subprocess
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, path])

            self._notify(
                "📂  Excel opened — close it before adding again.",
                "info",
            )
        except Exception as e:
            self._notify(f"❌  Could not open Excel: {e}", "error")

    # ---------- view entries ----------
    def view_entries(self):
        cfg = CONFIG["view_entries"]
        target = cfg.get("target", "internal")

        if target == "external":
            self._open_external_viewer(cfg)
            return

        # internal (existing behavior)
        if self._viewer is not None:
            try:
                if self._viewer.winfo_exists():
                    self._viewer.lift()
                    self._viewer.refresh()
                    return
            except Exception:
                pass
        if not self.current_sheet:
            self._notify("⚠  No sheet selected.", "warning"); return
        self._viewer = EntriesViewer(self, self.current_sheet)

    def _open_external_viewer(self, cfg):
        path = cfg.get("external_path")
        if not path:
            path = str(Path(__file__).with_name("FlickerView.exe"))
        if not Path(path).exists():
            self._notify(f"❌  External viewer not found: {path}", "error")
            return
        try:
            import subprocess
            exe = sys.executable
            if os.name == "nt":
                cand = str(Path(exe).with_name("pythonw.exe"))
                if Path(cand).exists():
                    exe = cand
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) \
                if os.name == "nt" else 0
            proc = subprocess.Popen([exe, path], creationflags=flags)

            if cfg.get("external_wait"):
                proc.wait()
                self._notify("📋  Viewer closed.", "info")
            else:
                self._notify("📋  Opened external viewer.", "info")
        except Exception as e:
            self._notify(f"❌  Could not launch viewer: {e}", "error")

    # ---------- undo ----------
    def undo_last(self):
        if not CONFIG["undo"]["enabled"]:
            return
        if self._busy:
            self._notify("⏳  Busy — please wait…", "info"); return
        if not self._undo_stack:
            self._notify("↩  Nothing to undo.", "info"); return

        item = self._undo_stack[-1]
        sheet_name = item["sheet"]
        row = item["row"]
        fcfg = CONFIG["file"]

        self._busy = True
        try:
            wb = openpyxl.load_workbook(fcfg["path"])
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' no longer exists.")
            ws = wb[sheet_name]
            for col in (fcfg["index_column"], fcfg["entry_column"], fcfg["timestamp_column"]):
                ws[f"{col}{row}"].value = None
            wb.save(fcfg["path"])
            wb.close()
            self._undo_stack.pop()
            self._notify(
                f"↩  Undid entry from '{sheet_name}' (row {row}).", "info"
            )
        except Exception as e:
            self._notify(f"❌  Undo failed: {e}", "error")
        finally:
            self._busy = False

    # ---------- main logic ----------
    def add_values(self):
        if self._busy:
            self._notify("⏳  Busy — please wait…", "info"); return

        raw = self.textbox.get("1.0", "end")
        if raw.endswith("\n"):
            raw = raw[:-1]

        ins = CONFIG["insert"]
        if ins["trim_whitespace"]:
            raw = raw.strip()

        if (not raw) or (ins["skip_empty"] and not raw.strip()):
            self._notify("⚠  Nothing to add.", "warning"); return

        if not self.current_sheet:
            self._notify("❌  No sheet selected.", "error"); return

        # Transform
        t = CONFIG["transform"]
        if any((t.get("case") and t["case"] != "none", t.get("prefix"),
                t.get("suffix"), t.get("trim_chars"), t.get("max_length"))):
            raw = _transform(raw)

        fcfg = CONFIG["file"]
        safety = CONFIG["safe_edit"]
        beh = CONFIG["behavior"]
        path = fcfg["path"]

        if not Path(path).exists():
            self._notify(f"❌  File not found: {path}", "error")
            self._maybe_auto_close(error=True); return

        if safety["refuse_if_locked"] and safety.get("check_file_lock", True):
            if _is_file_locked_by_excel(path):
                self._notify(
                    "❌  File is locked. Close Excel (or restart the app).",
                    "error",
                )
                self._maybe_auto_close(error=True); return

        # Work-on-copy
        working_path = path
        if safety["work_on_copy"]:
            try:
                src = Path(path)
                work = src.with_name(f"{src.stem}{safety['copy_suffix']}{src.suffix}")
                shutil.copy2(path, work)
                working_path = str(work)
            except Exception as e:
                self._notify(f"❌  Copy failed: {e}", "error")
                self._maybe_auto_close(error=True); return

        # Backup
        if CONFIG["backup"]["enabled"]:
            try:
                b = Path(path)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                bdir = CONFIG["backup"]["backup_dir"]
                target_dir = Path(bdir) if bdir else b.parent
                target_dir.mkdir(parents=True, exist_ok=True)
                bk = target_dir / f"{b.stem}_{ts}{b.suffix}"
                shutil.copy2(path, bk)
                keep = _to_int(CONFIG["backup"]["keep_last_n"], 20)
                pattern = f"{b.stem}_*{b.suffix}"
                olds = sorted(target_dir.glob(pattern),
                              key=lambda x: x.stat().st_mtime, reverse=True)
                for old in olds[keep:]:
                    try: old.unlink()
                    except Exception: pass
            except Exception as e:
                _log(f"backup failed: {e}")

        sheet_name = self.current_sheet
        ts_str = datetime.now().strftime(ins["timestamp_format"])
        idx_col = fcfg["index_column"]
        ent_col = fcfg["entry_column"]
        ts_col  = fcfg["timestamp_column"]
        result = {"row": 0}

        def do():
            wb = openpyxl.load_workbook(working_path)
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found.")
            ws = wb[sheet_name]

            last = _find_last_row(ws, idx_col)
            next_row = last + 1
            result["row"] = next_row

            ws[f"{idx_col}{next_row}"].value = next_row
            
            entry_cell = ws[f"{ent_col}{next_row}"]
            if isinstance(raw, str) and raw[:1] in ("=", "+", "-", "@"):
                entry_cell.number_format = "@"
                entry_cell.value = "'" + raw
            else:
                entry_cell.value = raw
                
            if ins["add_timestamp"]:
                ws[f"{ts_col}{next_row}"].value = ts_str

            if ins["mark_cells"]:
                for c in (idx_col, ent_col, ts_col):
                    try:
                        ws[f"{c}{next_row}"].fill.start_color.index = ins["mark_color"]
                    except Exception:
                        pass

            if not safety["dry_run"]:
                wb.save(working_path)
            wb.close()

        textbox_snapshot = self.textbox.get("1.0", "end-1c")

        self._busy = True
        self._notify("⏳  Saving…", "info")
        self.update_idletasks()
        try:
            do()
        except Exception as e:
            self._notify(f"❌  Excel error: {e}", "error")
            _log(traceback.format_exc())
            self._maybe_auto_close(error=True); return
        finally:
            self._busy = False

        next_row = result["row"]

        # Work-on-copy swap
        if safety["work_on_copy"] and not safety["dry_run"] and working_path != path:
            ok = True
            if safety["copy_verify"]:
                try:
                    def _probe(wb): return True
                    _run_excel(_probe, path=working_path)
                except Exception as e:
                    ok = False
                    _log(f"copy verify failed: {e}")
            if ok:
                try:
                    shutil.copy2(working_path, path)
                except Exception as e:
                    self._notify(f"❌  Swap failed: {e}", "error")
                    self._maybe_auto_close(error=True); return
            try:
                os.remove(working_path)
            except Exception:
                pass

        # Verify
        if (safety["verify_after_save"] and not safety["dry_run"] and next_row > 0):
            try:
                wb_v = openpyxl.load_workbook(path, data_only=True)
                ws_v = wb_v[sheet_name]
                v = ws_v[f"{ent_col}{next_row}"].value
                wb_v.close()
                if v is None or str(v).strip() == "":
                    self._notify(
                        "⚠  Saved but verification could not find the value.",
                        "warning",
                    )
            except Exception as e:
                _log(f"verify failed: {e}")

        # Undo / history
        if CONFIG["undo"]["enabled"] and not safety["dry_run"] and next_row > 0:
            self._undo_stack.append({"sheet": sheet_name, "row": next_row})
            if len(self._undo_stack) > _to_int(CONFIG["undo"]["stack_size"], 50):
                self._undo_stack.pop(0)

        if CONFIG["history"]["enabled"]:
            self._history.append(raw)
            if len(self._history) > _to_int(CONFIG["history"]["size"], 10):
                self._history.pop(0)

        # UI update
        if beh["clear_after_add"]:
            # Only clear if the textbox still holds exactly what we saved.
            # If the user typed more during the (blocking) save, leave it.
            current = self.textbox.get("1.0", "end-1c")
            if current == textbox_snapshot:
                self.textbox.delete("1.0", "end")
        if beh["refocus_after_add"]:
            self.textbox.focus_set()
            try:
                if beh.get("cursor_to_start_after_add", True):
                    self.textbox.mark_set("insert", "1.0")
                    self.textbox.see("1.0")
                else:
                    self.textbox.mark_set("insert", "end-1c")
                    self.textbox.see("end")
            except Exception:
                pass
        if self.hint is not None:
            self.hint.configure(text="")

        # Refresh recent entries strip so the new entry appears immediately
        if CONFIG["recent_entries"]["enabled"] and not safety["dry_run"]:
            self._schedule_recent_load(delay_ms=80)

        prefix = "🧪 [DRY RUN] " if safety["dry_run"] else "✅  "
        msg = f"{prefix}Added 1 entry to '{sheet_name}' at row {next_row}."
        self._notify(msg, "success")
        _log(msg)
        self._maybe_auto_close(error=False)


# ============================================================
#  ENTRY POINT
# ============================================================
if __name__ == "__main__":
    inst = CONFIG["instance"]
    listener = None
    if inst["single_instance"]:
        listener = _acquire_listener(inst["port"])
        if listener is None:
            _ping_existing(inst["port"])
            sys.exit(0)

    app = ThoughtsApp(listener_socket=listener)
    try:
        app.mainloop()
    finally:
        if listener is not None:
            try: listener.close()
            except Exception: pass