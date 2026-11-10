#!/usr/bin/env python3
"""
FlickerView — a fast, themeable viewer/editor for Excel (.xlsx/.xlsm)
workbooks, built with PySide6.

Features: 11 themes, per-tab custom colors, live search, undo/redo,
multi-column sorting, editable cells, and safe atomic saves.

Requires: PySide6, openpyxl
"""
import sys, os, json, copy, datetime, shutil, traceback
from pathlib import Path

if getattr(sys, 'frozen', False):
    HERE = Path(sys.executable).resolve().parent
else:
    HERE = Path(__file__).resolve().parent
LOG = HERE / "Entry_Viewer.log"

def log_exc():
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"\n=== {datetime.datetime.now().isoformat()} ===\n")
            traceback.print_exc(file=f)
    except Exception:
        pass

try:
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
    from PySide6.QtCore import (Qt, QAbstractTableModel, QModelIndex,
                                QSortFilterProxyModel, QTimer, QEvent, QRect)
    from PySide6.QtGui import (QFont, QColor, QBrush, QShortcut, QKeySequence,
                               QFontDatabase, QAction, QPen, QPainter,
                               QPainterPath, QIcon)
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QTableView, QTabBar, QWidget, QVBoxLayout,
        QHBoxLayout, QLineEdit, QPushButton, QLabel, QStatusBar, QMessageBox,
        QFileDialog, QDialog, QFormLayout, QSpinBox, QCheckBox, QComboBox,
        QToolBar, QHeaderView, QDialogButtonBox, QAbstractItemView,
        QSizePolicy, QMenu, QColorDialog, QInputDialog, QSlider, QSizeGrip,
        QStyledItemDelegate, QStyle, QStyleOptionSlider,
        QTabWidget, QScrollArea, QFrame, QTextEdit, QMenuBar)
except Exception:
    log_exc()
    try: print("Import failed. See", LOG, file=sys.stderr)
    finally: raise
    
class JumpSlider(QSlider):
    """QSlider that jumps to the clicked position instantly (like Excel)."""
    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            opt = QStyleOptionSlider()
            self.initStyleOption(opt)
            style = self.style()
            groove = style.subControlRect(QStyle.CC_Slider, opt, QStyle.SC_SliderGroove, self)
            handle = style.subControlRect(QStyle.CC_Slider, opt, QStyle.SC_SliderHandle, self)
            
            if self.orientation() == Qt.Horizontal:
                pos = ev.position().x() - (handle.width() / 2)
                groove_len = groove.width() - handle.width()
                val = QStyle.sliderValueFromPosition(self.minimum(), self.maximum(), pos - groove.x(), groove_len)
                self.setValue(val)
            else:
                pos = ev.position().y() - (handle.height() / 2)
                groove_len = groove.height() - handle.height()
                val = QStyle.sliderValueFromPosition(self.minimum(), self.maximum(), pos - groove.y(), groove_len, upsideDown=True)
                self.setValue(val)
            self.setSliderDown(True)
            ev.accept()
            return
        super().mousePressEvent(ev)

# ============================================================
# CONFIG
# ============================================================
CONFIG = {
    "file_path": str(HERE / "Sample_Thoughts.xlsx"),
    "font_family": "Segoe UI",
    "font_size": 18,
    "row_height": 38,
    "theme": "midnight",              # midnight | darker | amoled | slate |
                                      # nord | gruvbox | ocean | forest |
                                      # sunset | rose | light
    "accent_color": "#3b82f6",
    "row_coloring": "cycle",
    "first_row_is_header": True,
    "wrap_text": False,
    "frozen_first_col": False,
    "auto_backup": False,
    "max_col_width": 420,
    "min_col_width": 70,
    "window_width": 1400,
    "window_height": 900,
    "delete_row_shortcut": "Ctrl+Shift+Delete",
    "show_row_numbers": True,
    "show_col_letters": True,
    "show_gridlines": True,
    "grid_line_color": "",              # hex color or "" = theme default
    "grid_line_width": 1,               # 1 = thin, 2 = medium, 3 = thick
    "apply_to_all_sheets": True,
    "col_widths": {},
    "tab_colors": {},                 # UI-only tab text tints, keyed "filename::sheet"
    "tab_bg_colors": {},              # UI-only tab background tints
    "tab_accents": {},                # UI-only tab accent bar colors
    "tab_bg_alpha": 30,                # 0 = auto (theme-aware, 145–215); 1–255 = explicit
    "tab_accent_height": 2,           # px height of the accent bar (1–8)


    # --- Selection style ---
    "select_bg_override": "",         # hex or "" = theme default
    "select_fg_override": "",         # hex or "" = theme default (#ffffff)
    "select_border": True,            # draw a visible border on selected cells
    "select_border_color": "",        # hex or "" = use accent_color
    "select_border_width": 2,         # px

    # --- Row number (index) column ---
    "row_header_match_row": False,    # row numbers take the row's palette color
    "row_header_text_color": "",      # hex or "" = theme header_text
    "row_header_width": 40,           # px width of the row number strip

    # --- Table layout ---
    "header_row_height": 0,           # px; 0 = same as row_height
    # min_col_width already defined above (70)

    # --- UI chrome ---
    "show_toolbar": True,
    "show_statusbar": True,

    # --- Behavior ---
    "edit_trigger": "both",           # "double" | "any" | "both"
    "confirm_delete": False,          # confirm before deleting a row
    "save_zoom": True,                # remember zoom across sessions
    "default_sort_col": -1,           # -1 = unsorted on open
    "default_sort_desc": False,

    # --- Appearance extras ---
    "header_bold": True,              # bold column header row
    "header_text_color_override": "", # hex or "" = theme header_text

    # --- Table layout extras ---
    "cell_v_align": "middle",         # "top" | "middle" | "bottom"
    "autofit_on_open": True,          # autofit column widths when no saved widths
    "remember_col_widths": True,      # persist custom column widths

    # --- Behavior extras ---
    "confirm_exit_unsaved": True,     # prompt to save before closing with changes
}
CONFIG_FILE = HERE / ".entry_viewer_config.json"

THEMES = {
    "midnight": {
        "bg": "#0f1115", "bg_alt": "#161a21", "panel": "#1a1f27",
        "border": "#2a3038", "text": "#e6e8eb", "text_dim": "#8a9199",
        "header_bg": "#1f242d", "header_text": "#c9ced6",
        "gridline": "#232830", "select_bg": "#26406b",
        "row_alt1": "#14171d", "row_alt2": "#0f1115",
    },
    "darker": {
        "bg": "#08090c", "bg_alt": "#0e1014", "panel": "#12151a",
        "border": "#1e232a", "text": "#dfe2e6", "text_dim": "#6d747c",
        "header_bg": "#141820", "header_text": "#b8bec7",
        "gridline": "#181c22", "select_bg": "#1f3557",
        "row_alt1": "#0b0d11", "row_alt2": "#08090c",
    },
    "amoled": {
        "bg": "#000000", "bg_alt": "#0a0a0a", "panel": "#111111",
        "border": "#1f1f1f", "text": "#e8e8e8", "text_dim": "#808080",
        "header_bg": "#141414", "header_text": "#c9ced6",
        "gridline": "#1a1a1a", "select_bg": "#1a2f52",
        "row_alt1": "#050505", "row_alt2": "#000000",
    },
    "slate": {
        "bg": "#151821", "bg_alt": "#1b1f2a", "panel": "#202634",
        "border": "#2b3240", "text": "#e8ebef", "text_dim": "#8f97a3",
        "header_bg": "#242b3a", "header_text": "#cfd5de",
        "gridline": "#2a3040", "select_bg": "#2e4a7a",
        "row_alt1": "#181c26", "row_alt2": "#151821",
    },
    "nord": {
        "bg": "#2e3440", "bg_alt": "#3b4252", "panel": "#434c5e",
        "border": "#4c566a", "text": "#eceff4", "text_dim": "#8893a3",
        "header_bg": "#434c5e", "header_text": "#e5e9f0",
        "gridline": "#3b4252", "select_bg": "#5e81ac",
        "row_alt1": "#3b4252", "row_alt2": "#2e3440",
    },
    "gruvbox": {
        "bg": "#282828", "bg_alt": "#32302f", "panel": "#3c3836",
        "border": "#504945", "text": "#ebdbb2", "text_dim": "#a89984",
        "header_bg": "#3c3836", "header_text": "#fbf1c7",
        "gridline": "#3c3836", "select_bg": "#458588",
        "row_alt1": "#32302f", "row_alt2": "#282828",
    },
    "ocean": {
        "bg": "#0a1929", "bg_alt": "#0f2233", "panel": "#12263a",
        "border": "#1c3b57", "text": "#d9e6f2", "text_dim": "#7a91a8",
        "header_bg": "#142c46", "header_text": "#b8d4ec",
        "gridline": "#173550", "select_bg": "#1f5278",
        "row_alt1": "#0f2233", "row_alt2": "#0a1929",
    },
    "forest": {
        "bg": "#0f1a14", "bg_alt": "#14221a", "panel": "#18291f",
        "border": "#22392b", "text": "#dcead9", "text_dim": "#7d9880",
        "header_bg": "#1a2f22", "header_text": "#c7e0c0",
        "gridline": "#21382a", "select_bg": "#2e5a3c",
        "row_alt1": "#132018", "row_alt2": "#0f1a14",
    },
    "sunset": {
        "bg": "#1a100c", "bg_alt": "#241710", "panel": "#2d1e15",
        "border": "#3d2b1e", "text": "#f2e4d5", "text_dim": "#a08a75",
        "header_bg": "#2f1f15", "header_text": "#e8d1b6",
        "gridline": "#33241a", "select_bg": "#7a4a1f",
        "row_alt1": "#221610", "row_alt2": "#1a100c",
    },
    "rose": {
        "bg": "#170e14", "bg_alt": "#1e141c", "panel": "#251a23",
        "border": "#342432", "text": "#ecdee8", "text_dim": "#9d8894",
        "header_bg": "#261b24", "header_text": "#e0c8d8",
        "gridline": "#2a1d28", "select_bg": "#5c2e4d",
        "row_alt1": "#1e141c", "row_alt2": "#170e14",
    },
    "light": {
        "bg": "#ffffff", "bg_alt": "#f5f6f8", "panel": "#eef0f3",
        "border": "#d1d5db", "text": "#1f2328", "text_dim": "#57606a",
        "header_bg": "#e8ebf0", "header_text": "#1f2328",
        "gridline": "#e5e7eb", "select_bg": "#c7dbff",
        "row_alt1": "#f9fafb", "row_alt2": "#ffffff",
    },
}
ROW_PALETTE = ["#1a1f27", "#1d1a26", "#192430", "#1a241c",
               "#26201a", "#261a1d", "#182525", "#201a26"]
ROW_PALETTE_LIGHT = ["#ffffff", "#f7f9fc", "#fbf7ff", "#f6fff9",
                     "#fffaf3", "#fff5f5", "#f4fbfb", "#f9f6ff"]


def _rgba(hex_color, alpha):
    """Convert a hex color to an rgba() CSS string with the given alpha."""
    c = QColor(hex_color)
    return f"rgba({c.red()},{c.green()},{c.blue()},{alpha:.3f})"


def load_config():
    cfg = copy.deepcopy(CONFIG)
    if CONFIG_FILE.exists():
        try: cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except Exception: log_exc()
    return cfg


def save_config(cfg):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2, default=str),
                               encoding="utf-8")
    except Exception:
        log_exc()


# ============================================================
# STYLE HELPERS
# ============================================================
def _capture_style(cell):
    try:
        nf = cell.number_format
        f = cell.font; fill = cell.fill; b = cell.border
        is_default = (
            nf == "General"
            and not f.b and not f.i and not f.u
            and (f.sz is None or f.sz in (11, 11.0))
            and (f.color is None or f.color.rgb in (None, "FF000000"))
            and (fill.patternType is None)
            and (b.left.style is None and b.right.style is None
                 and b.top.style is None and b.bottom.style is None))
        if is_default:
            return None
        return {
            "font": copy.copy(f), "fill": copy.copy(fill),
            "border": copy.copy(b), "alignment": copy.copy(cell.alignment),
            "number_format": nf, "protection": copy.copy(cell.protection),
        }
    except Exception:
        return None


def _apply_style(cell, st):
    if not st: return
    try:
        cell.font = copy.copy(st["font"])
        cell.fill = copy.copy(st["fill"])
        cell.border = copy.copy(st["border"])
        cell.alignment = copy.copy(st["alignment"])
        cell.number_format = st["number_format"]
        cell.protection = copy.copy(st["protection"])
    except Exception:
        log_exc()


def _copy_style_row(row):
    return [dict(s) if s else None for s in row]


# ============================================================
# DATA
# ============================================================
class SheetData:
    __slots__ = ("name", "rows", "styles", "hidden_cols", "hidden",
                 "tab_color", "tab_bg_color", "tab_accent", "overrides")
    def __init__(self, name):
        self.name = name
        self.rows = []
        self.styles = []
        self.hidden_cols = set()
        self.hidden = False
        self.tab_color = None
        self.tab_bg_color = None
        self.tab_accent = None
        self.overrides = {}

    def max_cols(self):
        return max((len(r) for r in self.rows), default=0)

    def sync_styles_len(self):
        while len(self.styles) < len(self.rows):
            self.styles.append([])
        while len(self.styles) > len(self.rows):
            self.styles.pop()


class SheetModel(QAbstractTableModel):
    def __init__(self, app, sheet_data):
        super().__init__()
        self.app = app
        self.sheet = sheet_data

    def cfg(self, k, d=None): return self.app.config.get(k, d)
    def theme(self): return THEMES.get(self.cfg("theme", "midnight"), THEMES["midnight"])

    def _font_size_eff(self):
        base = self.sheet.overrides.get("font_size",
                                        self.app.config.get("font_size", 18))
        return max(8, int(round(base * getattr(self.app, "_zoom", 1.0))))

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.sheet.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.sheet.max_cols()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return (get_column_letter(section + 1) if orientation == Qt.Horizontal
                    else str(section + 1))
        if role == Qt.TextAlignmentRole:
            return int(Qt.AlignCenter)
        return None

    def _is_header_row(self, r):
        return self.cfg("first_row_is_header", True) and r == 0

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
        r, c = index.row(), index.column()
        if r >= len(self.sheet.rows): return None
        row = self.sheet.rows[r]
        val = row[c] if c < len(row) else None
        theme = self.theme()
        is_hdr = self._is_header_row(r)

        if role in (Qt.DisplayRole, Qt.EditRole):
            if val is None: return ""
            if isinstance(val, datetime.datetime):
                return (val.strftime("%Y-%m-%d %H:%M")
                        if (val.hour or val.minute) else val.strftime("%Y-%m-%d"))
            if isinstance(val, datetime.date):
                return val.strftime("%Y-%m-%d")
            if isinstance(val, float) and val.is_integer():
                return str(int(val))
            return str(val)

        if role == Qt.TextAlignmentRole:
            v = {
                "top":    Qt.AlignTop,
                "middle": Qt.AlignVCenter,
                "bottom": Qt.AlignBottom,
            }.get(self.cfg("cell_v_align", "middle"), Qt.AlignVCenter)
            if isinstance(val, bool): return int(Qt.AlignCenter | v)
            if isinstance(val, (int, float)):
                return int(Qt.AlignRight | v)
            if isinstance(val, (datetime.datetime, datetime.date)):
                return int(Qt.AlignCenter | v)
            return int(Qt.AlignLeft | v)

        if role == Qt.FontRole:
            f = QFont(self.cfg("font_family", "Segoe UI"))
            f.setPixelSize(self._font_size_eff())
            if is_hdr and self.cfg("header_bold", True):
                f.setBold(True)
            return f

        if role == Qt.ForegroundRole and is_hdr:
            override = (self.cfg("header_text_color_override", "") or "").strip()
            color = override or theme["header_text"]
            return QBrush(QColor(color))

        if role == Qt.BackgroundRole:
            if is_hdr:
                return QBrush(QColor(theme["header_bg"]))

            dr = r - (1 if self.cfg("first_row_is_header", True) else 0)
            mode = self.cfg("row_coloring", "cycle")

            if mode == "alternate":
                return QBrush(QColor(theme["row_alt1"] if dr % 2 == 0
                                     else theme["row_alt2"]))
            if mode == "cycle":
                palette = (ROW_PALETTE_LIGHT if self.cfg("theme") == "light"
                           else ROW_PALETTE)
                return QBrush(QColor(palette[dr % len(palette)]))
            return None
        return None

    def flags(self, index):
        if not index.isValid(): return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid(): return False
        r, c = index.row(), index.column()
        while len(self.sheet.rows) <= r: self.sheet.rows.append([])
        row = self.sheet.rows[r]
        while len(row) <= c: row.append(None)
        old = row[c]
        parsed = self._parse(value, old, c)
        if old == parsed: return False
        self.app.push_undo()
        row[c] = parsed
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        self.app.mark_dirty()
        # Real-time sort: if we're sorting by this exact column, re-apply
        # the sort so the row jumps into its correct position immediately.
        self.app._after_cell_edit(c)
        return True

    def _parse(self, value, old, col):
        if value is None: return None
        s = str(value).strip()
        if s == "": return None
        if isinstance(old, (datetime.datetime, datetime.date)):
            for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d",
                        "%d/%m/%Y %H:%M", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    dt = datetime.datetime.strptime(s, fmt)
                    if isinstance(old, datetime.date) and not isinstance(old, datetime.datetime):
                        return dt.date()
                    return dt
                except ValueError: continue
            return s
        if isinstance(old, (int, float)) and not isinstance(old, bool):
            try: return float(s) if ("." in s or "e" in s.lower()) else int(s)
            except ValueError: pass
        if old is None:
            vals = [row[col] for row in self.sheet.rows
                    if col < len(row) and row[col] is not None]
            if vals:
                dh = sum(1 for v in vals
                         if isinstance(v, (datetime.datetime, datetime.date)))
                if dh >= len(vals) * 0.5:
                    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d",
                                "%d/%m/%Y %H:%M", "%d/%m/%Y", "%m/%d/%Y"):
                        try: return datetime.datetime.strptime(s, fmt)
                        except ValueError: continue
        return s

    def _clone_row_styles(self, at):
        n = len(self.sheet.rows)
        if at is None or at >= n:
            src = n - 1
        else:
            src = at - 1
            if src >= 0 and self._is_header_row(src):
                src = at if at < n else None
        if src is None or src < 0 or src >= len(self.sheet.styles):
            return []
        return _copy_style_row(self.sheet.styles[src])

    def add_column(self):
        self.app.push_undo()
        self.beginResetModel()
        for row in self.sheet.rows:
            row.append(None)
        for styles in self.sheet.styles:
            styles.append(None)
        self.endResetModel()
        self.app.mark_dirty()

    def insert_column(self, at):
        self.app.push_undo()
        self.beginResetModel()
        for row in self.sheet.rows:
            row.insert(at, None)
        for styles in self.sheet.styles:
            styles.insert(at, None)
        new_hidden = set()
        for c in self.sheet.hidden_cols:
            new_hidden.add(c + 1 if c >= at else c)
        self.sheet.hidden_cols = new_hidden
        self.endResetModel()
        self.app.mark_dirty()

    def duplicate_column(self, col):
        self.app.push_undo()
        self.beginResetModel()
        for row in self.sheet.rows:
            val = row[col] if col < len(row) else None
            row.insert(col + 1, val)
        for styles in self.sheet.styles:
            st = styles[col] if col < len(styles) else None
            styles.insert(col + 1, _copy_style_row([st])[0] if st else None)
        new_hidden = set()
        for c in self.sheet.hidden_cols:
            new_hidden.add(c + 1 if c > col else c)
        self.sheet.hidden_cols = new_hidden
        self.endResetModel()
        self.app.mark_dirty()

    def delete_column(self, col):
        self.app.push_undo()
        self.beginResetModel()
        for row in self.sheet.rows:
            if col < len(row):
                del row[col]
        for styles in self.sheet.styles:
            if col < len(styles):
                del styles[col]
        new_hidden = set()
        for c in self.sheet.hidden_cols:
            if c == col: continue
            new_hidden.add(c - 1 if c > col else c)
        self.sheet.hidden_cols = new_hidden
        self.endResetModel()
        self.app.mark_dirty()

    def add_row(self, at=None):
        self.app.push_undo()
        self.beginResetModel()
        cols = max(self.columnCount(), 1)
        new_row = [None] * cols
        new_styles = self._clone_row_styles(at)
        if len(new_styles) < cols:
            new_styles += [None] * (cols - len(new_styles))
        if at is None or at >= len(self.sheet.rows):
            self.sheet.rows.append(new_row)
            self.sheet.styles.append(new_styles)
        else:
            self.sheet.rows.insert(at, new_row)
            self.sheet.styles.insert(at, new_styles)
        self.endResetModel()
        self.app.mark_dirty()

    def duplicate_row(self, r):
        if 0 <= r < len(self.sheet.rows):
            self.app.push_undo()
            self.beginResetModel()
            self.sheet.rows.insert(r + 1, list(self.sheet.rows[r]))
            src_styles = (self.sheet.styles[r]
                          if r < len(self.sheet.styles) else [])
            self.sheet.styles.insert(r + 1, _copy_style_row(src_styles))
            self.endResetModel()
            self.app.mark_dirty()

    def delete_row(self, r):
        if 0 <= r < len(self.sheet.rows) and not self._is_header_row(r):
            self.app.push_undo()
            self.beginResetModel()
            del self.sheet.rows[r]
            if r < len(self.sheet.styles):
                del self.sheet.styles[r]
            self.endResetModel()
            self.app.mark_dirty()

    def refresh_style(self):
        if self.sheet.rows and self.columnCount() > 0:
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(len(self.sheet.rows) - 1, self.columnCount() - 1))


class FilterProxy(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.filter_text = ""
        self._sort_order = Qt.AscendingOrder
        self._sort_disabled = True    # start in original order
        self.setSortCaseSensitivity(Qt.CaseInsensitive)
        self.setDynamicSortFilter(False)

    def set_filter(self, text):
        self.filter_text = (text or "").strip().lower()
        self.invalidateFilter()

    def sort(self, column, order=Qt.AscendingOrder):
        self._sort_order = order
        self._sort_disabled = (column < 0)
        super().sort(column, order)

    def clear_sort(self):
        self._sort_disabled = True
        # column -1 restores source order in Qt 6; also invalidate as a safety net
        try:
            super().sort(-1, Qt.AscendingOrder)
        except Exception:
            pass
        self.invalidate()

    def filterAcceptsRow(self, source_row, source_parent):
        if not self.filter_text: return True
        m = self.sourceModel()
        if m is None: return True
        if m.cfg("first_row_is_header", True) and source_row == 0:
            return True
        for c in range(m.columnCount()):
            txt = m.data(m.index(source_row, c, source_parent), Qt.DisplayRole)
            if txt and self.filter_text in str(txt).lower(): return True
        return False

    def lessThan(self, left, right):
        m = self.sourceModel()
        if m is None: return False
        # Keep header row at top regardless of sort order
        if m.cfg("first_row_is_header", True):
            if left.row() == 0 and right.row() != 0:
                return True
            if right.row() == 0 and left.row() != 0:
                return False
        # "Original order" sort — sort by source row index
        if self._sort_disabled:
            return left.row() < right.row()
        lv = m.data(left, Qt.DisplayRole) or ""
        rv = m.data(right, Qt.DisplayRole) or ""
        # Numeric if both sides parse as numbers
        try:
            return float(lv) < float(rv)
        except (ValueError, TypeError):
            pass
        # Date-aware string comparison: try ISO parse
        try:
            ld = datetime.datetime.fromisoformat(str(lv))
            rd = datetime.datetime.fromisoformat(str(rv))
            return ld < rd
        except (ValueError, TypeError):
            pass
        return str(lv).lower() < str(rv).lower()


# ============================================================
# TOAST
# ============================================================
class Toast(QLabel):
    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.hide()
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def show_msg(self, text, ms=1800, color="#2d8a4e"):
        self.setText(text)
        self.setStyleSheet(
            f"QLabel {{ background:{color}; color:white; padding:10px 20px;"
            f" border-radius:10px; font-size:15px; font-weight:600; }}")
        self.adjustSize()
        p = self.parent()
        if p:
            self.move(max((p.width() - self.width()) // 2, 10),
                      max(p.height() - self.height() - 55, 10))
        self.show(); self.raise_()
        self._timer.start(ms)


# ============================================================
# SETTINGS DIALOG
# ============================================================
class SettingsDialog(QDialog):
    """Tabbed settings dialog — Appearance · Table Layout · Selection ·
    Behavior · Advanced. Each tab is independently scrollable so the
    dialog never grows past the screen even as more settings are added."""

    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumSize(880, 690)
        self._cfg = cfg
        self._reset_requested = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 10, 10, 10)
        outer.setSpacing(10)

        self.tabs = QTabWidget()
        outer.addWidget(self.tabs, 1)

        self._build_appearance_tab()
        self._build_table_tab()
        self._build_selection_tab()
        self._build_behavior_tab()
        self._build_advanced_tab()

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        outer.addWidget(bb)

    # ---- helpers ----
    def _add_tab(self, title):
        """Creates a scrollable tab with a QFormLayout and returns the layout."""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        w = QWidget()
        form = QFormLayout(w)
        form.setContentsMargins(20, 20, 20, 20)
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        scroll.setWidget(w)
        self.tabs.addTab(scroll, title)
        return form

    def _section(self, form, title):
        """Bold section heading inside a form."""
        lbl = QLabel(f"<b>{title}</b>")
        lbl.setStyleSheet("margin-top: 6px;")
        form.addRow(lbl)

    def _color_row(self, initial_hex, fallback_hex):
        """Returns (layout, lineedit) for a color picker row."""
        edit = QLineEdit(initial_hex or "")
        edit.setPlaceholderText("(default)")
        btn = QPushButton("Pick…")
        btn.clicked.connect(
            lambda: self._pick_color(edit, fallback_hex))
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(edit, 1)
        row.addWidget(btn)
        return row, edit

    def _pick_color(self, target_field, fallback_hex):
        cur = target_field.text().strip() or fallback_hex
        c = QColorDialog.getColor(QColor(cur), self, "Pick color")
        if c.isValid():
            target_field.setText(c.name())

    def _pick_accent(self):
        c = QColorDialog.getColor(
            QColor(self.accent.text() or "#3b82f6"), self, "Accent color")
        if c.isValid():
            self.accent.setText(c.name())

    def _pick_grid_color(self):
        cur = self.grid_color.text().strip() or "#2a3038"
        c = QColorDialog.getColor(QColor(cur), self, "Grid line color")
        if c.isValid():
            self.grid_color.setText(c.name())

    def _browse(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select file", "", "Excel Files (*.xlsx *.xlsm)")
        if p:
            self.file_path.setText(p)

    # ================================================================
    # TAB 1 — Appearance
    # ================================================================
    def _build_appearance_tab(self):
        cfg = self._cfg
        form = self._add_tab("Appearance")

        self._section(form, "Theme")

        self.theme = QComboBox()
        self.theme.addItems(list(THEMES.keys()))
        self.theme.setCurrentText(cfg.get("theme", "midnight"))
        form.addRow("Theme:", self.theme)

        self.accent = QLineEdit(cfg.get("accent_color", "#3b82f6"))
        acc_btn = QPushButton("Pick…")
        acc_btn.clicked.connect(self._pick_accent)
        acc_row = QHBoxLayout()
        acc_row.setContentsMargins(0, 0, 0, 0)
        acc_row.addWidget(self.accent, 1)
        acc_row.addWidget(acc_btn)
        form.addRow("Accent color:", acc_row)

        self._section(form, "Header row")

        self.header_bold = QCheckBox()
        self.header_bold.setChecked(cfg.get("header_bold", True))
        self.header_bold.setToolTip(
            "Bold the first row (column headers). Off = same weight as data.")
        form.addRow("Header row bold:", self.header_bold)

        ht_row, self.header_text_color = self._color_row(
            cfg.get("header_text_color_override", ""), "#c9ced6")
        form.addRow("Header text color:", ht_row)

        self._section(form, "Typography")

        self.font_family = QComboBox()
        self.font_family.addItems(sorted(QFontDatabase.families()))
        self.font_family.setCurrentText(cfg.get("font_family", "Segoe UI"))
        form.addRow("Font family:", self.font_family)

        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(int(cfg.get("font_size", 18)))
        form.addRow("Font size:", self.font_size)

        self._section(form, "Rows")

        self.row_height = QSpinBox()
        self.row_height.setRange(16, 300)
        self.row_height.setValue(int(cfg.get("row_height", 38)))
        form.addRow("Row height:", self.row_height)

        self.header_row_height = QSpinBox()
        self.header_row_height.setRange(0, 300)
        self.header_row_height.setValue(
            int(cfg.get("header_row_height", 0)))
        self.header_row_height.setSpecialValueText("(same as rows)")
        self.header_row_height.setToolTip(
            "Height of the header row (row 0). 0 = same as regular rows.")
        form.addRow("Header row height:", self.header_row_height)

        self._section(form, "Row coloring")

        self.row_coloring = QComboBox()
        self.row_coloring.addItems(["off", "alternate", "cycle"])
        self.row_coloring.setCurrentText(cfg.get("row_coloring", "cycle"))
        form.addRow("Row coloring:", self.row_coloring)

        self.apply_all = QCheckBox()
        self.apply_all.setChecked(cfg.get("apply_to_all_sheets", True))
        self.apply_all.setToolTip(
            "ON = appearance settings affect every sheet.\n"
            "OFF = only the current sheet (stored as overrides).")
        form.addRow("Apply to all sheets:", self.apply_all)

    # ================================================================
    # TAB 2 — Table Layout
    # ================================================================
    def _build_table_tab(self):
        cfg = self._cfg
        form = self._add_tab("Table Layout")

        self._section(form, "Structure")

        self.first_hdr = QCheckBox()
        self.first_hdr.setChecked(cfg.get("first_row_is_header", True))
        form.addRow("First row is header:", self.first_hdr)

        self.wrap = QCheckBox()
        self.wrap.setChecked(cfg.get("wrap_text", False))
        form.addRow("Wrap text:", self.wrap)

        self.frozen = QCheckBox()
        self.frozen.setChecked(cfg.get("frozen_first_col", False))
        form.addRow("Freeze first column:", self.frozen)

        self.cell_v_align = QComboBox()
        self.cell_v_align.addItems(["top", "middle", "bottom"])
        self.cell_v_align.setCurrentText(cfg.get("cell_v_align", "middle"))
        self.cell_v_align.setToolTip(
            "Vertical alignment of text inside each cell.")
        form.addRow("Cell text alignment:", self.cell_v_align)

        self._section(form, "Headers")

        self.show_rows = QCheckBox()
        self.show_rows.setChecked(cfg.get("show_row_numbers", True))
        form.addRow("Show row numbers:", self.show_rows)

        self.row_header_width = QSpinBox()
        self.row_header_width.setRange(20, 200)
        self.row_header_width.setValue(
            int(cfg.get("row_header_width", 40)))
        form.addRow("Row header width:", self.row_header_width)

        self.show_cols = QCheckBox()
        self.show_cols.setChecked(cfg.get("show_col_letters", True))
        form.addRow("Show column letters:", self.show_cols)

        self._section(form, "Grid")

        self.show_grid = QCheckBox()
        self.show_grid.setChecked(cfg.get("show_gridlines", True))
        form.addRow("Show gridlines:", self.show_grid)

        gc_row, self.grid_color = self._color_row(
            cfg.get("grid_line_color", ""), "#2a3038")
        # replace the default picker to keep our custom grid picker
        gc_row.itemAt(1).widget().clicked.disconnect()
        gc_row.itemAt(1).widget().clicked.connect(self._pick_grid_color)
        form.addRow("Grid line color:", gc_row)

        self.grid_width = QSpinBox()
        self.grid_width.setRange(0, 5)
        self.grid_width.setValue(int(cfg.get("grid_line_width", 1)))
        self.grid_width.setToolTip("0 = hidden, 1 = thin, 2–5 = thicker")
        form.addRow("Grid line width:", self.grid_width)

        self._section(form, "Columns")

        self.max_w = QSpinBox()
        self.max_w.setRange(100, 2000)
        self.max_w.setValue(int(cfg.get("max_col_width", 420)))
        form.addRow("Max column width:", self.max_w)

        self.min_w = QSpinBox()
        self.min_w.setRange(20, 400)
        self.min_w.setValue(int(cfg.get("min_col_width", 70)))
        form.addRow("Min column width:", self.min_w)

        self.autofit_on_open = QCheckBox()
        self.autofit_on_open.setChecked(
            cfg.get("autofit_on_open", True))
        self.autofit_on_open.setToolTip(
            "Auto-fit columns when a sheet has no saved width yet.\n"
            "Off = keep the file's own widths on first open.")
        form.addRow("Autofit columns on open:", self.autofit_on_open)

        self.remember_col_widths = QCheckBox()
        self.remember_col_widths.setChecked(
            cfg.get("remember_col_widths", True))
        self.remember_col_widths.setToolTip(
            "Persist column widths you drag between sessions.\n"
            "Off = widths reset to the file's defaults each open.")
        form.addRow("Remember column widths:", self.remember_col_widths)

    # ================================================================
    # TAB 3 — Selection
    # ================================================================
    def _build_selection_tab(self):
        cfg = self._cfg
        form = self._add_tab("Selection")

        self._section(form, "Selection fill")

        sb_row, self.select_bg = self._color_row(
            cfg.get("select_bg_override", ""), "#26406b")
        form.addRow("Selection background:", sb_row)

        sf_row, self.select_fg = self._color_row(
            cfg.get("select_fg_override", ""), "#ffffff")
        form.addRow("Selection text color:", sf_row)

        self._section(form, "Selection border")

        self.select_border = QCheckBox()
        self.select_border.setChecked(cfg.get("select_border", True))
        self.select_border.setToolTip(
            "Adds a visible border around selected cells — helps when\n"
            "selection fill happens to match a row color.")
        form.addRow("Draw selection border:", self.select_border)

        sbc_row, self.select_border_color = self._color_row(
            cfg.get("select_border_color", ""), "#3b82f6")
        form.addRow("Border color:", sbc_row)

        self.select_border_width = QSpinBox()
        self.select_border_width.setRange(0, 6)
        self.select_border_width.setValue(
            int(cfg.get("select_border_width", 2)))
        form.addRow("Border width:", self.select_border_width)

        self._section(form, "Row number colors")

        self.row_header_match = QCheckBox()
        self.row_header_match.setChecked(
            cfg.get("row_header_match_row", False))
        self.row_header_match.setToolTip(
            "Row numbers on the left take the row's own color.\n"
            "Helps trace a row across the table when scrolling.")
        form.addRow("Row numbers match row color:", self.row_header_match)

        rht_row, self.row_header_text = self._color_row(
            cfg.get("row_header_text_color", ""), "#8a9199")
        form.addRow("Row number text color:", rht_row)

    # ================================================================
    # TAB 4 — Behavior
    # ================================================================
    def _build_behavior_tab(self):
        cfg = self._cfg
        form = self._add_tab("Behavior")

        self._section(form, "Editing")

        self.edit_trigger = QComboBox()
        self.edit_trigger.addItems(["double", "any", "both"])
        self.edit_trigger.setCurrentText(cfg.get("edit_trigger", "both"))
        self.edit_trigger.setToolTip(
            "double  = double-click a cell to edit\n"
            "any     = start typing on a focused cell\n"
            "both    = either works")
        form.addRow("Edit trigger:", self.edit_trigger)

        self.confirm_delete = QCheckBox()
        self.confirm_delete.setChecked(cfg.get("confirm_delete", False))
        form.addRow("Confirm before deleting row:", self.confirm_delete)

        self._section(form, "Session")

        self.save_zoom = QCheckBox()
        self.save_zoom.setChecked(cfg.get("save_zoom", True))
        self.save_zoom.setToolTip(
            "Persist zoom level between sessions. Turn off to always\n"
            "start at 100%.")
        form.addRow("Remember zoom on close:", self.save_zoom)

        self._section(form, "On open")

        self.default_sort_col = QSpinBox()
        self.default_sort_col.setRange(-1, 999)
        self.default_sort_col.setValue(
            int(cfg.get("default_sort_col", -1)))
        self.default_sort_col.setSpecialValueText("(none)")
        self.default_sort_col.setToolTip(
            "Column to sort by when a file opens.\n"
            "-1 = unsorted; 0 = A; 1 = B; 2 = C; …\n"
            "Users can still click headers to sort manually.")
        form.addRow("Default sort column:", self.default_sort_col)

        self.default_sort_desc = QCheckBox()
        self.default_sort_desc.setChecked(
            cfg.get("default_sort_desc", False))
        form.addRow("Default sort descending:", self.default_sort_desc)

        self._section(form, "Safety")

        self.confirm_exit_unsaved = QCheckBox()
        self.confirm_exit_unsaved.setChecked(
            cfg.get("confirm_exit_unsaved", True))
        self.confirm_exit_unsaved.setToolTip(
            "Prompt to save before closing with unsaved changes.\n"
            "Off = close silently (unsaved changes are lost).")
        form.addRow("Confirm exit with unsaved changes:",
                    self.confirm_exit_unsaved)

    # ================================================================
    # TAB 5 — Advanced
    # ================================================================
    def _build_advanced_tab(self):
        cfg = self._cfg
        form = self._add_tab("Advanced")

        self._section(form, "Chrome visibility")

        self.show_toolbar = QCheckBox()
        self.show_toolbar.setChecked(cfg.get("show_toolbar", True))
        self.show_toolbar.setToolTip(
            "Hide the top toolbar. Settings remains accessible via\n"
            "the menu (View → Settings) or Ctrl+,")
        form.addRow("Show toolbar:", self.show_toolbar)

        self.show_statusbar = QCheckBox()
        self.show_statusbar.setChecked(cfg.get("show_statusbar", True))
        form.addRow("Show status bar:", self.show_statusbar)

        self._section(form, "Saving")

        self.backup = QCheckBox()
        self.backup.setChecked(cfg.get("auto_backup", False))
        self.backup.setToolTip(
            "Copies the file to <name>.bak next to it before every save.")
        form.addRow("Auto-backup on save:", self.backup)

        self._section(form, "File")

        self.file_path = QLineEdit(cfg.get("file_path", ""))
        b = QPushButton("Browse")
        b.clicked.connect(self._browse)
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(self.file_path, 1)
        row.addWidget(b)
        form.addRow("Default file:", row)

        self._section(form, "Maintenance")

        reset_btn = QPushButton("Reset all settings to defaults…")
        reset_btn.setToolTip(
            "Restores every option on every tab to its factory default.\n"
            "Your XLSX files are not touched — only the app's settings.")
        reset_btn.clicked.connect(self._reset_to_defaults)
        form.addRow(reset_btn)

    def _reset_to_defaults(self):
        r = QMessageBox.question(
            self, "Reset to defaults",
            "Reset ALL settings to their factory defaults?\n\n"
            "This cannot be undone. Your XLSX files are not affected.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if r == QMessageBox.Yes:
            self._reset_requested = True
            self.accept()

    # ---- collect values ----
    def values(self):
        return {
            # Appearance
            "theme": self.theme.currentText(),
            "accent_color": self.accent.text().strip() or "#3b82f6",
            "font_family": self.font_family.currentText(),
            "font_size": self.font_size.value(),
            "row_height": self.row_height.value(),
            "header_row_height": self.header_row_height.value(),
            "row_coloring": self.row_coloring.currentText(),
            "apply_to_all_sheets": self.apply_all.isChecked(),
            # Table layout
            "first_row_is_header": self.first_hdr.isChecked(),
            "wrap_text": self.wrap.isChecked(),
            "frozen_first_col": self.frozen.isChecked(),
            "show_row_numbers": self.show_rows.isChecked(),
            "row_header_width": self.row_header_width.value(),
            "show_col_letters": self.show_cols.isChecked(),
            "show_gridlines": self.show_grid.isChecked(),
            "grid_line_color": self.grid_color.text().strip(),
            "grid_line_width": self.grid_width.value(),
            "max_col_width": self.max_w.value(),
            "min_col_width": self.min_w.value(),
            # Selection
            "select_bg_override": self.select_bg.text().strip(),
            "select_fg_override": self.select_fg.text().strip(),
            "select_border": self.select_border.isChecked(),
            "select_border_color": self.select_border_color.text().strip(),
            "select_border_width": self.select_border_width.value(),
            "row_header_match_row": self.row_header_match.isChecked(),
            "row_header_text_color": self.row_header_text.text().strip(),
            # Behavior
            "edit_trigger": self.edit_trigger.currentText(),
            "confirm_delete": self.confirm_delete.isChecked(),
            "save_zoom": self.save_zoom.isChecked(),
            "default_sort_col": self.default_sort_col.value(),
            "default_sort_desc": self.default_sort_desc.isChecked(),
            "confirm_exit_unsaved": self.confirm_exit_unsaved.isChecked(),
            # Appearance extras
            "header_bold": self.header_bold.isChecked(),
            "header_text_color_override":
                self.header_text_color.text().strip(),
            # Table layout extras
            "cell_v_align": self.cell_v_align.currentText(),
            "autofit_on_open": self.autofit_on_open.isChecked(),
            "remember_col_widths": self.remember_col_widths.isChecked(),
            # Advanced
            "show_toolbar": self.show_toolbar.isChecked(),
            "show_statusbar": self.show_statusbar.isChecked(),
            "auto_backup": self.backup.isChecked(),
            "file_path": self.file_path.text(),
        }

# ============================================================
# ROW HEADER — optionally paints the row's own background color
# behind the row number so you can trace a row across the table.
# ============================================================
class RowHeader(QHeaderView):
    """Vertical row-number header.
    Font size is explicitly controlled via set_font_px() so it stays
    consistent regardless of QSS reapplication on Settings OK."""

    def __init__(self, parent, get_paint_data):
        super().__init__(Qt.Vertical, parent)
        self._get_paint_data = get_paint_data
        self._font_px = 16

    def set_font_px(self, px):
        self._font_px = max(8, int(px))
        f = self.font()
        f.setPixelSize(self._font_px)
        self.setFont(f)
        self.viewport().update()

    def paintSection(self, painter, rect, logicalIndex):
        info = None
        try:
            info = self._get_paint_data(logicalIndex)
        except Exception:
            info = None

        if not info:
            super().paintSection(painter, rect, logicalIndex)
            return

        bg, fg, border = info
        painter.save()
        painter.fillRect(rect, bg)
        if border is not None:
            painter.setPen(border)
            painter.drawLine(rect.bottomLeft(), rect.bottomRight())
            painter.drawLine(rect.topRight(), rect.bottomRight())
        # Always use the explicit font size — never inherit from stale state
        f = painter.font()
        f.setPixelSize(self._font_px)
        painter.setFont(f)
        if fg is not None:
            painter.setPen(fg)
        painter.drawText(rect, Qt.AlignCenter, str(logicalIndex + 1))
        painter.restore()

class GridDelegate(QStyledItemDelegate):
    """Draws gridlines + optional selection border on top of standard
    cell painting. super().paint() handles model BackgroundRole (row
    colors), selection fill, text, alignment etc. exactly as QTableView
    normally would."""

    def __init__(self, get_grid_cfg, get_sel_cfg):
        super().__init__()
        self._get_grid_cfg = get_grid_cfg
        self._get_sel_cfg = get_sel_cfg

    def createEditor(self, parent, option, index):
        editor = QTextEdit(parent)
        editor.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        editor.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        return editor

    def setEditorData(self, editor, index):
        text = index.model().data(index, Qt.EditRole)
        editor.setPlainText(str(text) if text else "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.EditRole)

    def eventFilter(self, editor, event):
        if event.type() == QEvent.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                if event.modifiers() & Qt.ShiftModifier:
                    event.accept()
                    editor.insertPlainText("\n")
                    return True
                else:
                    self.commitData.emit(editor)
                    self.closeEditor.emit(editor, QAbstractItemDelegate.NoHint)
                    return True
        return super().eventFilter(editor, event)

    def paint(self, painter, option, index):
        super().paint(painter, option, index)

        # ---- grid lines ----
        try:
            visible, width, color_str = self._get_grid_cfg()
        except Exception:
            visible = False
        if visible and width > 0:
            c = QColor(color_str)
            if c.isValid():
                painter.save()
                pen = QPen(c)
                pen.setWidth(int(width))
                painter.setPen(pen)
                r = option.rect
                painter.drawLine(r.topRight(), r.bottomRight())
                painter.drawLine(r.bottomLeft(), r.bottomRight())
                painter.restore()

        # ---- selection border ----
        try:
            sel_cfg = self._get_sel_cfg()
        except Exception:
            sel_cfg = None
        if (sel_cfg and sel_cfg["border"]
                and (option.state & QStyle.State_Selected)):
            c = QColor(sel_cfg["color"])
            if c.isValid():
                painter.save()
                pen = QPen(c)
                pen.setWidth(int(sel_cfg["width"]))
                painter.setPen(pen)
                r = option.rect.adjusted(0, 0, -1, -1)
                painter.drawRect(r)
                painter.restore()

class ColoredTabBar(QTabBar):
    """Fully custom-painted QTabBar.
    Calls super().paintEvent(ev) so Qt calculates layout, but tells super() 
    to draw transparent text. Then we draw our backgrounds, accent, and 
    text exactly once — no double-rendering artifacts.
    """

    def __init__(self, parent=None,
                 get_default_accent=None,
                 get_default_border=None,
                 get_theme=None,
                 get_bg_alpha=None,
                 get_accent_height=None):
        super().__init__(parent)
        self._tab_bg = {}
        self._tab_accent = {}
        self._tab_text = {}
        self._get_default_accent = get_default_accent
        self._get_default_border = get_default_border
        self._get_theme = get_theme
        self._get_bg_alpha = get_bg_alpha
        self._get_accent_height = get_accent_height
        self._font_px = 12

    # ---------- transparency helpers ----------
    def _effective_bg_alpha(self):
        """Return the alpha for the custom background tint.
        - Config override (1–255) wins when set.
        - Otherwise auto-tune to the theme: light surfaces need ~84% opacity
          to read as a real tint; dark surfaces only need ~60%."""
        if self._get_bg_alpha:
            try:
                a = int(self._get_bg_alpha())
                if a > 0:
                    return max(1, min(255, a))
            except (TypeError, ValueError):
                pass
        t = self._get_theme() if self._get_theme else THEMES["midnight"]
        c = QColor(t["bg_alt"])
        lum = (0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue()) / 255.0
        return int(145 + 70 * lum)   # 145 (amoled) … 215 (light)

    def _effective_accent_height(self):
        if self._get_accent_height:
            try:
                return max(1, min(8, int(self._get_accent_height())))
            except (TypeError, ValueError):
                pass
        return 3

    def set_font_px(self, px):
        self._font_px = max(8, int(px))
        self.update()

    def set_tab_bg(self, idx, color_hex):
        if color_hex:
            self._tab_bg[idx] = QColor(color_hex)
        else:
            self._tab_bg.pop(idx, None)
        self.update()

    def set_tab_accent(self, idx, color_hex):
        if color_hex:
            self._tab_accent[idx] = QColor(color_hex)
        else:
            self._tab_accent.pop(idx, None)
        self.update()

    def set_tab_text_color(self, idx, color):
        self._tab_text[idx] = color
        # Make QTabBar's default text painter transparent so super() doesn't draw it
        self.setTabTextColor(idx, QColor(0, 0, 0, 0))
        self.update()

    def clear_all_custom(self):
        self._tab_bg.clear()
        self._tab_accent.clear()
        self._tab_text.clear()
        for i in range(self.count()):
            self.setTabTextColor(i, QColor(0, 0, 0, 0))
        self.update()


    def paintEvent(self, ev):
        # 1. Layout + default QSS background + transparent text
        super().paintEvent(ev)

        if not self._tab_bg and not self._tab_accent and not self._tab_text:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        f = painter.font()
        f.setPixelSize(self._font_px)
        painter.setFont(f)

        cur_idx = self.currentIndex()

        # Layouts reparent widgets to their container, so self.parent()
        # is tab_container — NOT ExcelViewer. Use the injected callback.
        t = self._get_theme() if self._get_theme else THEMES["midnight"]
        default_text = QColor(t['text'])
        default_text_dim = QColor(t['text_dim'])

        # Theme-aware transparency + configurable accent thickness
        bg_alpha = self._effective_bg_alpha()
        accent_h = self._effective_accent_height()

        for i in range(self.count()):
            rect = self.tabRect(i)
            if rect.isEmpty():
                continue

            # tabRect() includes the QSS margin-right. Shrink it so the 
            # custom background tint and accent bar don't bleed into the gap.
            rect.setRight(rect.right() - 2)

            is_sel = (i == cur_idx)

            # ---- 2. Custom Background Tint ----
            bg = self._tab_bg.get(i)
            if bg:
                bg = QColor(bg)
                bg.setAlpha(bg_alpha)   # theme-aware (light themes get more)

                path = QPainterPath()
                r = 10
                path.moveTo(rect.left(), rect.bottom())
                path.lineTo(rect.left(), rect.top() + r)
                path.quadTo(rect.left(), rect.top(), rect.left() + r, rect.top())
                path.lineTo(rect.right() - r, rect.top())
                path.quadTo(rect.right(), rect.top(), rect.right(), rect.top() + r)
                path.lineTo(rect.right(), rect.bottom())
                path.closeSubpath()

                painter.fillPath(path, bg)

            # ---- 3. Accent Bar (selected tab only) ----
            # The accent bar is painted ONLY on the current tab. A per-tab
            # custom accent (Right-click → Set tab colors) overrides the
            # theme accent for that tab; otherwise the theme accent_color
            # is used. Non-selected tabs never draw an accent bar, so the
            # only "current tab" cue is this single underline.
            if is_sel:
                accent = self._tab_accent.get(i)
                if accent is None and self._get_default_accent:
                    accent = self._get_default_accent()

                if accent:
                    painter.save()
                    path = QPainterPath()
                    r = 10
                    path.moveTo(rect.left(), rect.bottom())
                    path.lineTo(rect.left(), rect.top() + r)
                    path.quadTo(rect.left(), rect.top(), rect.left() + r, rect.top())
                    path.lineTo(rect.right() - r, rect.top())
                    path.quadTo(rect.right(), rect.top(), rect.right(), rect.top() + r)
                    path.lineTo(rect.right(), rect.bottom())
                    path.closeSubpath()
                    painter.setClipPath(path)

                    painter.fillRect(
                        QRect(rect.left(), rect.bottom() - accent_h,
                              rect.width(), accent_h),
                        accent)
                    painter.restore()

            # ---- 4. Draw Text (Since super() drew transparent) ----
            text = self.tabText(i)
            if text:
                tc = self._tab_text.get(i)
                if not tc:
                    tc = default_text if is_sel else default_text_dim
                painter.setPen(QPen(tc))
                painter.drawText(rect, Qt.AlignCenter, text)

        painter.end()




# ============================================================
# MAIN WINDOW
# ============================================================
class ExcelViewer(QMainWindow):
    def __init__(self, cfg):
        super().__init__()
        self.config = cfg
        self._zoom = float(cfg.get("zoom", 1.0))                       # session-only, not persisted
        self.file_path = cfg.get("file_path") or ""
        self.workbook = None
        self.sheets = {}
        self.sheet_order = []
        self.models = {}
        self.proxies = {}
        self.current_sheet = None
        self.dirty = False
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo = 20
        self._restoring = False
        self._autofitting = False
        self._syncing = False
        self._last_save = ""
        # sort state: -1 = original order, else column index
        self._sort_col = -1
        self._sort_desc = False
        # session state (Windows shutdown handshake)
        self._session_ending = False
        # pending column-width changes: {(sheet, col_idx): width}
        self._pending_col_sizes = {}
        # debounce timers for expensive operations
        self._zoom_apply_timer = QTimer(self)
        self._zoom_apply_timer.setSingleShot(True)
        self._zoom_apply_timer.timeout.connect(self._apply_zoom)
        self._zoom_save_timer = QTimer(self)
        self._zoom_save_timer.setSingleShot(True)
        self._zoom_save_timer.timeout.connect(self._persist_zoom)
        self._col_save_timer = QTimer(self)
        self._col_save_timer.setSingleShot(True)
        self._col_save_timer.timeout.connect(self._flush_col_widths)

        self.setWindowTitle("FlickerView")
        self.resize(cfg.get("window_width", 1400), cfg.get("window_height", 900))
        self._build_ui(); self._build_menu(); self._apply_theme()
        self._bind_shortcuts()
        # reflect saved zoom in the slider before any file loads
        if self._zoom != 1.0:
            self.zoom_slider.blockSignals(True)
            self.zoom_slider.setValue(int(round(self._zoom * 100)))
            self.zoom_slider.blockSignals(False)
            self.zoom_label.setText(f"{int(round(self._zoom * 100))}%")
        QTimer.singleShot(10, self._initial_file_load)
        # Windows sends WM_QUERYENDSESSION before shutdown/logout.
        # Qt surfaces this as commitDataRequest — use it to save
        # cleanly so Windows never has to force-kill the process.
        app = QApplication.instance()
        if app is not None:
            try:
                app.commitDataRequest.connect(self._on_commit_data)
            except Exception:
                log_exc()
    def _initial_file_load(self):
        if self.file_path and Path(self.file_path).exists():
            self._open_file(self.file_path)
        elif self.file_path:
            # Auto-generate a sample file if it doesn't exist
            try:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Thoughts"
                ws.append(["ID", "Thought", "Timestamp"])
                ws.append([1, "This is a sample entry. Click 'Add Row' to try it!", "2023-10-27 10:00:00"])
                ws.append([2, "You can right-click the tabs to change colors.", "2023-10-27 10:05:00"])
                wb.save(self.file_path)
                self._toast.show_msg("Created Sample File ✓")
                self._open_file(self.file_path)
            except Exception as e:
                self._toast.show_msg(f"Could not create sample file: {e}", color="#a33")
        self._update_status()

    # ---------------------------------------------------------- session end
    def _on_commit_data(self, manager):
        """Windows is asking us to save before shutdown/logout.
        This fires during WM_QUERYENDSESSION — BEFORE closeEvent — so
        if we try to save here and set a flag, closeEvent will skip the
        'Save?' modal (which would otherwise block shutdown and make
        Windows force-kill the process, risking NTFS unclean state)."""
        self._session_ending = True
        try:
            # Flush pending debounced writes first
            if self._zoom_save_timer.isActive():
                self._zoom_save_timer.stop()
                self._persist_zoom()
            if self._col_save_timer.isActive():
                self._col_save_timer.stop()
                self._flush_col_widths()
            if self.dirty:
                self._save_file()
        except Exception:
            log_exc()

    # ---------------------------------------------------------- UI
    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        central.setObjectName("centralWidget")
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        # --- Menu Bar ---
        m = self.menuBar()
        m.setContextMenuPolicy(Qt.CustomContextMenu)
        m.customContextMenuRequested.connect(self._global_context_menu)

        # --- Toolbar (Ribbon) ---
        tb = QToolBar(); tb.setMovable(False)
        self._toolbar = tb
        self.addToolBar(tb)
        tb.setContextMenuPolicy(Qt.CustomContextMenu)
        tb.customContextMenuRequested.connect(self._global_context_menu)
        
        self.open_btn = QPushButton("📂 Open")
        self.save_btn = QPushButton("💾 Save")
        self.add_btn  = QPushButton("＋ Add Row")
        self.ins_btn  = QPushButton("⤵ Insert Row")
        self.dup_btn  = QPushButton("⧉ Duplicate")
        self.del_btn  = QPushButton("🗑 Delete Row")
        self.color_btn = QPushButton("🎨 Colors")
        self.settings_btn = QPushButton("⚙ Settings")
        for b in (self.open_btn, self.save_btn, self.add_btn, self.ins_btn,
                  self.dup_btn, self.del_btn, self.color_btn, self.settings_btn):
            b.setCursor(Qt.PointingHandCursor)
        tb.addWidget(self.open_btn); tb.addWidget(self.save_btn); tb.addSeparator()
        tb.addWidget(self.add_btn); tb.addWidget(self.ins_btn)
        tb.addWidget(self.dup_btn); tb.addWidget(self.del_btn); tb.addSeparator()
        tb.addWidget(self.color_btn)
        tb.addWidget(self.settings_btn)

        self._search_wrap = QWidget()
        sr = QHBoxLayout(self._search_wrap)
        sr.setContentsMargins(10, 8, 10, 8)
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍 Search all columns…   (Esc to hide)")
        self.search.setClearButtonEnabled(True)
        sr.addWidget(self.search)
        self._search_wrap.setVisible(False)
        root.addWidget(self._search_wrap)

        self.tabbar = ColoredTabBar(
            self,
            get_default_accent=self._tab_default_accent,
            get_default_border=self._tab_default_border,
            get_theme=self._tab_theme,
            get_bg_alpha=self._tab_bg_alpha,
            get_accent_height=self._tab_accent_height,
        )
        self.tabbar.setExpanding(False)
        self.tabbar.setDrawBase(False)
        self.tabbar.setFixedHeight(38)
        self.tabbar.setContextMenuPolicy(Qt.CustomContextMenu)
        tab_container = QWidget()
        tab_container.setContextMenuPolicy(Qt.CustomContextMenu)
        tab_container.customContextMenuRequested.connect(self._global_context_menu)
        tab_lay = QHBoxLayout(tab_container)
        tab_lay.setContentsMargins(0, 8, 0, 0)
        tab_lay.setSpacing(0)
        tab_lay.addWidget(self.tabbar)
        self.add_sheet_btn = QPushButton("＋")
        self.add_sheet_btn.setObjectName("addSheetBtn")
        self.add_sheet_btn.setFixedSize(38, 38)
        self.add_sheet_btn.setCursor(Qt.PointingHandCursor)
        self.add_sheet_btn.clicked.connect(self._new_sheet)
        tab_lay.addWidget(self.add_sheet_btn, 0, Qt.AlignBottom)
        tab_lay.addStretch(1)
        root.addWidget(tab_container)

        self.frozen_view = QTableView()
        self.frozen_view.setVisible(False)
        self.frozen_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.frozen_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.frozen_view.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed)
        self.frozen_view.verticalHeader().hide()
        self.frozen_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.view = QTableView()
        self.view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.view.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed)
        self.view.setSortingEnabled(True)
        self.view.setShowGrid(True)
        hh = self.view.horizontalHeader()
        hh.setStretchLastSection(False)
        hh.setSectionsClickable(True)
        hh.setHighlightSections(False)
        hh.setContextMenuPolicy(Qt.CustomContextMenu)
        self.frozen_view.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)

        # Custom row header so it can optionally match row colors
        old_vh = self.view.verticalHeader()
        old_vh.setParent(None)
        old_vh.deleteLater()
        self._row_header = RowHeader(self.view, self._row_header_paint_data)
        self.view.setVerticalHeader(self._row_header)
        self._row_header.setContextMenuPolicy(Qt.CustomContextMenu)
        self._row_header.customContextMenuRequested.connect(self._row_header_context_menu)

        # Manual 3-state sort (overrides Qt's auto-sorting)
        self.view.setSortingEnabled(False)
        # Kill Qt's built-in "click a header → select the whole column"
        # behavior. QTableView connects sectionPressed/sectionEntered to
        # its own selectColumn slot; disconnecting those makes header
        # clicks do ONLY sorting, nothing else.
        for _sig_name in ("sectionPressed", "sectionEntered"):
            try:
                getattr(hh, _sig_name).disconnect()
            except Exception:
                pass
        hh.sectionClicked.connect(self._on_header_clicked)

        self.view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._cell_context_menu)

        tc = QWidget(); tl = QHBoxLayout(tc)
        tl.setContentsMargins(10, 0, 10, 10); tl.setSpacing(0)
        tl.addWidget(self.frozen_view); tl.addWidget(self.view, 1)
        root.addWidget(tc, 1)

        sb = QStatusBar(); self.setStatusBar(sb)
        sb.setSizeGripEnabled(False)
        self.status_label = QLabel("Ready"); sb.addWidget(self.status_label, 1)

        # --- zoom bar (Excel-style, bottom right) ---
        self.zoom_out_btn = QPushButton("−")
        self.zoom_out_btn.setObjectName("zoomBtn")
        self.zoom_out_btn.setFixedSize(22, 22)
        self.zoom_out_btn.setToolTip("Zoom out (Ctrl+-)")
        self.zoom_out_btn.clicked.connect(self.zoom_out)

        self.zoom_slider = JumpSlider(Qt.Horizontal)
        self.zoom_slider.setObjectName("zoomSlider")
        self.zoom_slider.setRange(50, 250)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setFixedWidth(110)
        self.zoom_slider.setFixedHeight(22)
        self.zoom_slider.setToolTip("Zoom 50–250%")
        
        # Ensure continuous, smooth dragging
        self.zoom_slider.setSingleStep(1)
        self.zoom_slider.setPageStep(5)
        self.zoom_slider.setTracking(True)
        self.zoom_slider.setTickPosition(QSlider.NoTicks) # Remove default OS chunks/ticks
        
        self.zoom_slider.valueChanged.connect(self._on_zoom_slider)

        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setObjectName("zoomBtn")
        self.zoom_in_btn.setFixedSize(22, 22)
        self.zoom_in_btn.setToolTip("Zoom in (Ctrl+=)")
        self.zoom_in_btn.clicked.connect(self.zoom_in)

        self.zoom_label = QLabel("100%")
        self.zoom_label.setObjectName("zoomPercent")
        self.zoom_label.setMinimumWidth(56)
        self.zoom_label.setFixedHeight(22)
        self.zoom_label.setAlignment(Qt.AlignCenter)

        self.save_label = QLabel("")
        self._zoom_group = QWidget()
        self._zoom_group.setObjectName("zoomGroup")
        self._zoom_group.setFixedHeight(24)
        zg_lay = QHBoxLayout(self._zoom_group)
        zg_lay.setContentsMargins(6, 1, 6, 1)
        zg_lay.setSpacing(4)
        zg_lay.setAlignment(Qt.AlignVCenter)
        zg_lay.addWidget(self.zoom_out_btn, 0, Qt.AlignVCenter)
        zg_lay.addWidget(self.zoom_slider, 0, Qt.AlignVCenter)
        zg_lay.addWidget(self.zoom_in_btn, 0, Qt.AlignVCenter)
        zg_lay.addWidget(self.zoom_label, 0, Qt.AlignVCenter)
        self.sort_label = QLabel("")
        self.sort_label.setObjectName("sortIndicator")
        self.sort_label.setVisible(False)
        sb.addPermanentWidget(self.sort_label)
        sb.addPermanentWidget(self.save_label)
        sb.addPermanentWidget(self._zoom_group)

        self._toast = Toast(self)

        self.open_btn.clicked.connect(self._open_dialog)
        self.save_btn.clicked.connect(self._save_file)
        self.add_btn.clicked.connect(self._add_row)
        self.ins_btn.clicked.connect(self._insert_row)
        self.dup_btn.clicked.connect(self._duplicate_row)
        self.del_btn.clicked.connect(self._delete_row)
        self.color_btn.clicked.connect(self._cycle_row_coloring)
        self.settings_btn.clicked.connect(self._open_settings)
        self.search.textChanged.connect(self._on_search)
        self.tabbar.currentChanged.connect(self._on_tab_changed)
        self.tabbar.customContextMenuRequested.connect(self._tab_context_menu)
        hh.customContextMenuRequested.connect(self._col_header_context_menu)
        self.frozen_view.horizontalHeader().customContextMenuRequested.connect(
            self._col_header_context_menu)
        hh.sectionResized.connect(self._on_col_resized)
        self.frozen_view.horizontalHeader().sectionResized.connect(
            self._on_frozen_col_resized)

        self.view.verticalScrollBar().valueChanged.connect(
            self.frozen_view.verticalScrollBar().setValue)
        self.frozen_view.verticalScrollBar().valueChanged.connect(
            self.view.verticalScrollBar().setValue)

        # Grid + selection border rendered by a delegate so model
        # BackgroundRole (row colors) is preserved. Native grid is off.
        self.view.setShowGrid(False)
        self.frozen_view.setShowGrid(False)
        self._grid_delegate = GridDelegate(self._grid_cfg, self._sel_cfg)
        self.view.setItemDelegate(self._grid_delegate)
        self.frozen_view.setItemDelegate(self._grid_delegate)

        self.view.installEventFilter(self)
        self.view.viewport().installEventFilter(self)
        self.view.horizontalHeader().installEventFilter(self)
        self.frozen_view.installEventFilter(self)
        self.frozen_view.viewport().installEventFilter(self)
        self.frozen_view.horizontalHeader().installEventFilter(self)
        self.search.installEventFilter(self)

    def _build_menu(self):
        m = self.menuBar()
        f = m.addMenu("&File")
        a = QAction("Open…", self); a.setShortcut("Ctrl+O")
        a.triggered.connect(self._open_dialog); f.addAction(a)
        a = QAction("Save", self); a.setShortcut("Ctrl+S")
        a.triggered.connect(self._save_file); f.addAction(a)
        a = QAction("Reload from disk", self)
        a.triggered.connect(self._reload_file); f.addAction(a)
        f.addSeparator()
        a = QAction("Exit", self); a.setShortcut("Ctrl+Q")
        a.triggered.connect(self.close); f.addAction(a)

        e = m.addMenu("&Edit")
        a = QAction("Undo", self); a.setShortcut("Ctrl+Z")
        a.triggered.connect(self.undo); e.addAction(a)
        a = QAction("Redo", self); a.setShortcut("Ctrl+Y")
        a.triggered.connect(self.redo); e.addAction(a)
        e.addSeparator()
        a = QAction("Copy", self); a.setShortcut("Ctrl+C")
        a.triggered.connect(self._on_copy); e.addAction(a)
        a = QAction("Cut", self); a.setShortcut("Ctrl+X")
        a.triggered.connect(self._on_cut); e.addAction(a)
        a = QAction("Paste", self); a.setShortcut("Ctrl+V")
        a.triggered.connect(self._on_paste); e.addAction(a)
        e.addSeparator()
        a = QAction("Add Row", self); a.setShortcut("Ctrl+N")
        a.triggered.connect(self._add_row); e.addAction(a)
        a = QAction("Insert Row Above", self)
        a.triggered.connect(self._insert_row); e.addAction(a)
        a = QAction("Duplicate Row", self)
        a.triggered.connect(self._duplicate_row); e.addAction(a)
        a = QAction("Delete Row", self); a.setShortcut("Ctrl+Shift+Delete")
        a.triggered.connect(self._delete_row); e.addAction(a)

        v = m.addMenu("&View")

        # Direct visibility toggles
        self._act_show_ribbon = QAction("Show Ribbon", self)
        self._act_show_ribbon.setCheckable(True)
        self._act_show_ribbon.setChecked(
            self.config.get("show_toolbar", True))
        self._act_show_ribbon.toggled.connect(self._toggle_ribbon)
        v.addAction(self._act_show_ribbon)

        self._act_show_menubar = QAction("Show Menu Bar", self)
        self._act_show_menubar.setCheckable(True)
        self._act_show_menubar.setChecked(True)
        self._act_show_menubar.toggled.connect(self._toggle_menubar)
        v.addAction(self._act_show_menubar)

        self._act_show_statusbar = QAction("Show Status Bar", self)
        self._act_show_statusbar.setCheckable(True)
        self._act_show_statusbar.setChecked(
            self.config.get("show_statusbar", True))
        self._act_show_statusbar.toggled.connect(self._toggle_statusbar)
        v.addAction(self._act_show_statusbar)

        v.addSeparator()

        sort_menu = v.addMenu("Sort")
        a = QAction("Reset Sort (original order)", self)
        a.triggered.connect(lambda: self._set_sort(-1, False))
        sort_menu.addAction(a)
        sort_menu.addSeparator()
        # Common column shortcuts — these are the A/B/C columns typically
        # used by add_thought (index / entry / timestamp). Click a header
        # for any other column.
        a = QAction("By Index ↑", self)
        a.triggered.connect(lambda: self._set_sort(0, False))
        sort_menu.addAction(a)
        a = QAction("By Index ↓", self)
        a.triggered.connect(lambda: self._set_sort(0, True))
        sort_menu.addAction(a)
        sort_menu.addSeparator()
        a = QAction("By Entry ↑", self)
        a.triggered.connect(lambda: self._set_sort(1, False))
        sort_menu.addAction(a)
        a = QAction("By Entry ↓", self)
        a.triggered.connect(lambda: self._set_sort(1, True))
        sort_menu.addAction(a)
        sort_menu.addSeparator()
        a = QAction("By Date ↑", self)
        a.triggered.connect(lambda: self._set_sort(2, False))
        sort_menu.addAction(a)
        a = QAction("By Date ↓", self)
        a.triggered.connect(lambda: self._set_sort(2, True))
        sort_menu.addAction(a)

        v.addSeparator()
        a = QAction("Zoom In", self); a.setShortcut("Ctrl+=")
        a.triggered.connect(self.zoom_in); v.addAction(a)
        a = QAction("Zoom Out", self); a.setShortcut("Ctrl+-")
        a.triggered.connect(self.zoom_out); v.addAction(a)
        a = QAction("Reset Zoom", self); a.setShortcut("Ctrl+0")
        a.triggered.connect(self.zoom_reset); v.addAction(a)
        v.addSeparator()
        a = QAction("Cycle Row Coloring", self)
        a.triggered.connect(self._cycle_row_coloring); v.addAction(a)
        a = QAction("Settings…", self)
        a.triggered.connect(self._open_settings); v.addAction(a)

    def _bind_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+F"), self, activated=self._toggle_search)
        QShortcut(QKeySequence("Ctrl+,"), self, activated=self._open_settings)
        QShortcut(QKeySequence("Ctrl+Z"), self, activated=self.undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, activated=self.redo)
        QShortcut(QKeySequence("Ctrl+Shift+Z"), self, activated=self.redo)
        QShortcut(QKeySequence(self.config.get("delete_row_shortcut",
                                               "Ctrl+Shift+Delete")),
                  self, activated=self._delete_row)
        QShortcut(QKeySequence("Ctrl+="), self, activated=self.zoom_in)
        QShortcut(QKeySequence("Ctrl++"), self, activated=self.zoom_in)
        QShortcut(QKeySequence("Ctrl+Shift+="), self, activated=self.zoom_in)
        QShortcut(QKeySequence("Ctrl+-"), self, activated=self.zoom_out)
        QShortcut(QKeySequence("Ctrl+0"), self, activated=self.zoom_reset)
        QShortcut(QKeySequence("Ctrl+Home"), self, activated=self._go_first_cell)
        QShortcut(QKeySequence("Ctrl+End"), self, activated=self._go_last_cell)

    def eventFilter(self, obj, ev):
        # Resolve widget refs safely — eventFilter can be called before
        # _build_ui finishes creating them (e.g. on the mini-bar's Polish event).
        view = getattr(self, "view", None)
        frozen_view = getattr(self, "frozen_view", None)
        search = getattr(self, "search", None)




        # Double-click on horizontal header (to add column past the end)
        if view is not None and (obj is view.horizontalHeader() or obj is frozen_view.horizontalHeader()):
            if ev.type() == QEvent.MouseButtonDblClick:
                pos = ev.position().toPoint()
                if view.horizontalHeader().logicalIndexAt(pos) < 0:
                    self._add_col_at_end()
                    return True

        if view is not None and (obj is view or obj is frozen_view or obj is view.viewport() or obj is frozen_view.viewport()):
            if ev.type() == QEvent.MouseButtonDblClick:
                pos = ev.position().toPoint()
                idx = view.indexAt(pos)
                if not idx.isValid() and view.model() is not None:
                    model = view.model()
                    
                    # Check if click is below the last row
                    if model.rowCount() > 0:
                        last_row_rect = view.visualRect(model.index(model.rowCount() - 1, 0))
                        if pos.y() > last_row_rect.bottom():
                            self._add_row_at_end()
                            return True
                            
                    # Check if click is right of the last column
                    if model.columnCount() > 0:
                        last_col_rect = view.visualRect(model.index(0, model.columnCount() - 1))
                        if pos.x() > last_col_rect.right():
                            self._add_col_at_end()
                            return True

            if ev.type() == QEvent.KeyPress:
                mods = ev.modifiers(); key = ev.key()
                if mods & Qt.ControlModifier:
                    if key == Qt.Key_C: self._on_copy(); return True
                    if key == Qt.Key_X: self._on_cut(); return True
                    if key == Qt.Key_V: self._on_paste(); return True
                if key == Qt.Key_Escape:
                    if self.view.state() != QAbstractItemView.EditingState:
                        self.close(); return True
        elif search is not None and obj is search:
            if ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Escape:
                if search.text():
                    search.clear()
                    return True
                self._hide_search()
                return True
        return super().eventFilter(obj, ev)

    # ---------------------------------------------------------- grid
    def _grid_cfg(self):
        """Returns (visible, width, color) for the grid delegate."""
        t = THEMES.get(self.config.get("theme", "midnight"), THEMES["midnight"])
        visible = bool(self.config.get("show_gridlines", True))
        width = int(self.config.get("grid_line_width", 1))
        color_str = (self.config.get("grid_line_color", "") or "").strip()
        if not color_str:
            color_str = t["gridline"]
        return (visible, width, color_str)
        
    # ---------------------------------------------------------- sort
    def _on_header_clicked(self, section):
        """3-state sort: none → asc → desc → none."""
        if self._sort_col == section:
            if not self._sort_desc:
                self._sort_desc = True   # asc → desc
            else:
                self._sort_col = -1      # desc → none
                self._sort_desc = False
        else:
            self._sort_col = section
            self._sort_desc = False      # first click on a column = asc
        self._apply_sort()

    def _set_sort(self, col, desc):
        self._sort_col = col
        self._sort_desc = desc
        self._apply_sort()

    def _apply_sort(self):
        proxy = self.view.model()
        if proxy is None:
            return
        if self._sort_col < 0:
            proxy.clear_sort()
        else:
            order = Qt.DescendingOrder if self._sort_desc else Qt.AscendingOrder
            proxy.sort(self._sort_col, order)
        self._update_sort_indicator()
        self._update_status()

    def _after_cell_edit(self, edited_col):
        """Called by SheetModel.setData after every successful edit.
        - If wrap text is on, recompute row heights so wrapped content fits
        - If sorting is active on the edited column, re-sort the proxy."""
        # Wrap text: recompute row heights so the new content fits
        if self.config.get("wrap_text", False):
            m = self.view.model()
            if m is not None and m.rowCount() < 500:
                self.view.resizeRowsToContents()
                self.frozen_view.resizeRowsToContents()
        # Sorting: re-apply only when the edited cell is in the sort column
        if self._sort_col < 0:
            return
        if edited_col != self._sort_col:
            return
        proxy = self.view.model()
        if proxy is None:
            return
        cur = self.view.currentIndex()
        was_valid = cur.isValid()
        cur_row = cur.row() if was_valid else -1
        self._apply_sort()
        if was_valid and 0 <= cur_row < proxy.rowCount():
            self.view.setCurrentIndex(proxy.index(cur_row, cur.column()))

    def _update_sort_indicator(self):
        if not hasattr(self, "sort_label"):
            return
        if self._sort_col < 0:
            self.sort_label.setText("")
            self.sort_label.setVisible(False)
            return
        arrow = "↓" if self._sort_desc else "↑"
        letter = get_column_letter(self._sort_col + 1)
        self.sort_label.setText(f"  Sort: {arrow} {letter}  ")
        self.sort_label.setVisible(True)

    def _apply_default_sort(self):
        """Apply the user's configured default sort on file open."""
        col = int(self.config.get("default_sort_col", -1))
        desc = bool(self.config.get("default_sort_desc", False))
        self._sort_col = col if col >= 0 else -1
        self._sort_desc = desc
        proxy = self.view.model()
        if proxy is not None:
            if self._sort_col < 0:
                proxy.clear_sort()
            else:
                order = Qt.DescendingOrder if self._sort_desc \
                        else Qt.AscendingOrder
                proxy.sort(self._sort_col, order)
        self._update_sort_indicator()

    # ---------------------------------------------------------- selection style
    def _sel_cfg(self):
        t = THEMES.get(self.config.get("theme", "midnight"), THEMES["midnight"])
        color = (self.config.get("select_border_color", "") or "").strip() \
                or self.config.get("accent_color", "#3b82f6")
        return {
            "border": bool(self.config.get("select_border", True)),
            "color": color,
            "width": int(self.config.get("select_border_width", 2)),
        }

    def _row_header_paint_data(self, logical_index):
        """Return (bg, fg, border) for the row header, or None for default."""
        if not self.config.get("row_header_match_row", False):
            return None
        model = self._row_header.model()
        if model is None:
            return None
        idx = model.index(logical_index, 0)
        bg_brush = model.data(idx, Qt.BackgroundRole)
        if bg_brush is None or not isinstance(bg_brush, QBrush):
            return None
        bg = bg_brush.color()
        t = THEMES.get(self.config.get("theme", "midnight"), THEMES["midnight"])
        fg_hex = (self.config.get("row_header_text_color", "") or "").strip() \
                 or t["text_dim"]
        border = QColor(t["border"])
        return bg, QColor(fg_hex), border

    # ---------------------------------------------------------- theme
    def _apply_theme(self):
        t = THEMES.get(self.config.get("theme", "midnight"), THEMES["midnight"])
        acc = self.config.get("accent_color", "#3b82f6")
        fs = self.config.get("font_size", 18)
        ff = self.config.get("font_family", "Segoe UI")
        small = max(fs - 5, 11); med = max(fs - 2, 12)

        # Row header must use the same font size as the main cell text so
        # the numbers match the data visually. Explicit size prevents QSS 
        # reapplication (or clicking OK on Settings) from resetting it.
        if hasattr(self, "_row_header"):
            self._row_header.set_font_px(fs)
            
        # Tab bar custom paint must use the same font size as the QSS
        if hasattr(self, "tabbar"):
            self.tabbar.set_font_px(med)

        window_bg  = t['bg']
        panel_qss  = t['panel']
        alt_qss    = t['bg_alt']
        header_qss = t['header_bg']
        table_bg   = t['bg']
        table_alt  = t['bg_alt']
        btn_solid = t['bg_alt']   # buttons and inputs stay solid for comfort

        # --- selection colors (overridable) ---
        sel_bg = (self.config.get("select_bg_override", "") or "").strip() \
                 or t['select_bg']
        sel_fg = (self.config.get("select_fg_override", "") or "").strip() \
                 or "#ffffff"

        # Grid is drawn by GridDelegate (installed in _build_ui).
        # Do NOT add QTableView::item QSS — QStyleSheetStyle would then
        # override the model's BackgroundRole and kill row colors.

        app = QApplication.instance()
        f = QFont(ff); f.setPixelSize(fs); app.setFont(f)

        self.setStyleSheet(f"""
            QMainWindow {{ background:{window_bg}; color:{t['text']}; }}
            QWidget {{ background:{window_bg}; color:{t['text']}; }}
            QMenuBar {{ background:{t['panel']}; color:{t['text']}; border-bottom: 1px solid {t['border']}; padding: 4px 0px; }}
            QMenuBar::item {{ background: transparent; padding: 6px 14px; border-radius: 4px; }}
            QMenuBar::item:selected {{ background:{t['header_bg']}; }}
            QMenu {{ background:{t['panel']}; color:{t['text']};
                     border:1px solid {t['border']}; padding: 6px 0px; }}
            QMenu::item {{ padding: 6px 24px 6px 24px; margin: 0px 2px; }}
            QMenu::item:selected {{ background:{t['select_bg']}; color:#ffffff;
                                    border-radius: 4px; }}
            QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 10px; }}
            QToolBar {{ background:{panel_qss}; border:none; padding:8px; spacing:6px; }}
            QToolBar::separator {{ background:{t['border']}; width:1px; margin:4px 6px; }}
            QPushButton {{ background:{btn_solid}; color:{t['text']};
                border:1px solid {t['border']}; border-radius:8px;
                padding:8px 16px; font-size:{med}px; }}
            QPushButton:hover {{ background:{t['header_bg']}; border-color:{t['text_dim']}; }}
            QPushButton:pressed {{ background:{t['text_dim']}; color:{t['bg']}; }}
            QLineEdit {{ background:{btn_solid}; color:{t['text']};
                border:1px solid {t['border']}; border-radius:8px;
                padding:8px 12px; font-size:{med}px;
                selection-background-color:{acc}; }}
            QLineEdit:focus {{ border-color:{acc}; }}
            QComboBox, QSpinBox {{ background:{btn_solid}; color:{t['text']};
                border:1px solid {t['border']}; border-radius:6px;
                padding:6px 10px; font-size:{med}px; }}
            QCheckBox {{ color:{t['text']}; font-size:{med}px; }}
            QTextEdit {{ background:{btn_solid}; color:{t['text']}; 
                border:2px solid {acc}; font-size:{med}px; padding: 2px; }}
                
            QTabBar {{ padding: 0px; margin: 0px; border: none; }}
            QTabBar::tab {{ background: {alt_qss};
                padding:0px 22px; border:none;
                border-top-left-radius:10px;
                border-top-right-radius:10px; margin-right:2px; margin-top:0px; margin-bottom:0px; margin-left:0px;
                min-width: 60px; min-height: 38px; max-height: 38px;
                font-size:{med}px; }}
            QTabBar::tab:selected {{ background: {panel_qss}; }}
            QTabBar::tab:hover:!selected {{ background: {t['header_bg']}; }}
            QTableView {{ background:{table_bg}; alternate-background-color:{table_alt};
                color:{t['text']}; gridline-color:transparent;
                selection-background-color:{sel_bg}; selection-color:{sel_fg};
                border:1px solid {t['border']}; border-radius:8px; }}
            QHeaderView::section {{ background:{header_qss}; color:{t['header_text']};
                border:none; border-right:1px solid {t['border']};
                border-bottom:1px solid {t['border']};
                padding:8px; font-weight:700; font-size:{med}px; }}
            QStatusBar {{ background:{panel_qss}; color:{t['text_dim']};
                border-top:1px solid {t['border']}; }}
            QStatusBar::item {{ border: none; }}
            QStatusBar QLabel {{ color:{t['text_dim']}; padding:0 6px; font-size:{small}px; }}
            QScrollBar:vertical, QScrollBar:horizontal {{ background:{t['bg']};
                border:none; width:13px; height:13px; }}
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
                background:{t['border']}; border-radius:6px;
                min-height:24px; min-width:24px; }}
            QScrollBar::handle:hover {{ background:{t['text_dim']}; }}
            QScrollBar::add-line, QScrollBar::sub-line {{ background:none; height:0; }}
            QSlider {{ border: none; background: transparent; }}
            QSlider::groove:horizontal {{ background:{t['border']}; height:4px;
                border-radius:2px; margin: 0px 0; }}
            QSlider::sub-page:horizontal {{ background:{t['text_dim']}; height:4px;
                border-radius:2px; }}
            QSlider::add-page:horizontal {{ background:{t['border']}; height:4px;
                border-radius:2px; }}
            QSlider::handle:horizontal {{ background:{t['text']}; width:14px; height:14px;
                margin:-5px 0; border-radius:7px; }}
            QSlider::handle:horizontal:hover {{ background:{acc}; }}

            #zoomGroup {{ background: transparent; border: none; }}
            #zoomGroup QWidget {{ background: transparent; }}
            #zoomBtn {{ background: transparent; color: {t['text_dim']};
                border: none; padding: 0; margin: 0;
                font-size: {med}px; font-weight: 700; }}
            #zoomBtn:hover {{ color: {t['text']}; }}
            #zoomBtn:pressed {{ background: {t['text_dim']};
                color: {t['bg']}; border-radius: 4px; }}
            #zoomPercent {{ color: {t['text_dim']}; font-size: {small}px;
                font-weight: 600; background: transparent;
                padding: 0; margin: 0; }}
            #zoomSlider {{ background: transparent; padding: 0; margin: 0; }}
            #zoomSlider::groove:horizontal {{ background: {t['border']};
                height: 3px; border-radius: 1px; }}
            #zoomSlider::sub-page:horizontal {{ background: {t['text_dim']};
                height: 3px; border-radius: 1px; }}
            #zoomSlider::handle:horizontal {{ background: {t['text_dim']};
                width: 12px; height: 12px; margin: -5px 0; border-radius: 6px; }}
            #zoomSlider::handle:horizontal:hover {{ background: {t['text']}; }}
            #addSheetBtn {{ background: {alt_qss}; color: {t['text_dim']};
                border: none;
                border-top-left-radius: 10px; border-top-right-radius: 10px;
                font-size: {med}px; font-weight: 700; padding: 0;
                min-width: 38px; max-width: 38px; min-height: 38px; max-height: 38px;
                margin-top: 0px; margin-bottom: 0px; margin-right: 0px; margin-left: 0px; }}
            #addSheetBtn:hover {{ background: {t['header_bg']}; color: {t['text']}; }}
            #sortIndicator {{ color: {t['text_dim']}; font-size: {small}px;
                font-weight: 600; padding: 0 4px;
                background: {alt_qss}; border-radius: 8px; }}
        """)
        for m in self.models.values(): m.refresh_style()
        self._apply_headers_visibility()
        self._apply_tab_colors()
        self._update_sort_indicator()
        # Repaint the row header when colors change
        if hasattr(self, "_row_header"):
            self._row_header.viewport().update()

    def _apply_headers_visibility(self):
        show_rows = self.config.get("show_row_numbers", True)
        show_cols = self.config.get("show_col_letters", True)
        self.view.verticalHeader().setVisible(show_rows)
        self.view.horizontalHeader().setVisible(show_cols)
        self.frozen_view.verticalHeader().setVisible(False)
        self.frozen_view.horizontalHeader().setVisible(show_cols)
        # Native grid stays disabled — the GridDelegate draws custom lines
        self.view.setShowGrid(False)
        self.frozen_view.setShowGrid(False)

        # Row header width
        if hasattr(self, "_row_header"):
            rw = int(self.config.get("row_header_width", 40))
            self._row_header.setFixedWidth(rw)

        # Chrome visibility
        if hasattr(self, "_toolbar"):
            self._toolbar.setVisible(
                self.config.get("show_toolbar", True))
        sb = self.statusBar()
        if sb is not None:
            sb.setVisible(self.config.get("show_statusbar", True))

        # Edit trigger mode
        mode = self.config.get("edit_trigger", "both")
        trig = QAbstractItemView.NoEditTriggers
        if mode in ("double", "both"):
            trig = trig | QAbstractItemView.DoubleClicked
        if mode in ("any", "both"):
            trig = (trig | QAbstractItemView.EditKeyPressed
                    | QAbstractItemView.AnyKeyPressed)
        self.view.setEditTriggers(trig)
        self.frozen_view.setEditTriggers(trig)

        # Repaint so delegate picks up any grid config change immediately
        if hasattr(self, "view"):
            self.view.viewport().update()
            self.frozen_view.viewport().update()
            
    def _update_tab_text_colors(self):
        # Kept as an alias so existing callers keep working
        self._apply_tab_colors()

    # ---------------------------------------------------------- file ops
    def _open_dialog(self):
        start = str(Path(self.file_path).parent) if self.file_path else ""
        p, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", start, "Excel Files (*.xlsx *.xlsm)")
        if p: self._open_file(p)

    def _reload_file(self):
        if not self.file_path: return
        if self.dirty:
            r = QMessageBox.question(
                self, "Discard changes?",
                "You have unsaved changes. Reload anyway?",
                QMessageBox.Yes | QMessageBox.No)
            if r != QMessageBox.Yes: return
        self._open_file(self.file_path)

    def _open_file(self, path):
        if not path: return
        p = Path(path)
        if not p.exists():
            QMessageBox.warning(self, "File not found", f"Not found:\n{path}"); return
        
        # Remember this path so the same file reopens next launch.
        self.config["file_path"] = str(p)
        save_config(self.config)

        try:
            kwargs = {"data_only": False}
            if p.suffix.lower() == ".xlsm": kwargs["keep_vba"] = True
            wb = openpyxl.load_workbook(path, **kwargs)
        except PermissionError:
            QMessageBox.critical(self, "File locked",
                "The file is open in Excel.\n\nClose it first and try again."); return
        except Exception as e:
            QMessageBox.critical(self, "Could not open", str(e)); return

        self.workbook = wb
        self.file_path = path
        self.sheets.clear(); self.sheet_order.clear()
        self.models.clear(); self.proxies.clear()
        self.undo_stack.clear(); self.redo_stack.clear()
        self.dirty = False

        for ws in wb.worksheets:
            sd = SheetData(ws.title)
            for row in ws.iter_rows():
                vals, styles = [], []
                for cell in row:
                    vals.append(cell.value)
                    styles.append(_capture_style(cell))
                sd.rows.append(vals)
                sd.styles.append(styles)
            for col_letter, dim in ws.column_dimensions.items():
                if getattr(dim, "hidden", False):
                    try: sd.hidden_cols.add(column_index_from_string(col_letter) - 1)
                    except (ValueError, TypeError): pass
            sd.hidden = (ws.sheet_state == "hidden")
            # Tab colors: read from CONFIG (UI-only) — never from the file.
            sd.tab_color = self._get_tab_color(ws.title)
            sd.tab_bg_color = self._get_tab_bg_color(ws.title)
            sd.tab_accent = self._get_tab_accent_color(ws.title)
            self.sheets[ws.title] = sd
            if not sd.hidden:
                self.sheet_order.append(ws.title)

        if not self.sheet_order and self.sheets:
            for n, sd in self.sheets.items():
                sd.hidden = False
                self.sheet_order.append(n)

        self._rebuild_tabbar()
        self._update_title()
        if self.sheet_order:
            self.tabbar.setCurrentIndex(0)
            self._switch_sheet(0)
            self._apply_default_sort()
        self._update_status()

    def _save_file(self):
        if not self.workbook or not self.file_path:
            self._toast.show_msg("No file open", color="#a33"); return
        if not self.dirty:
            self._toast.show_msg("No changes to save", color="#666"); return
        if self.config.get("auto_backup", False):
            try: shutil.copy2(self.file_path, str(self.file_path) + ".bak")
            except Exception: log_exc()
        try:
            for name, sd in self.sheets.items():
                ws = self.workbook[name]
                sd.sync_styles_len()
                nrows = len(sd.rows)
                ncols = max((len(r) for r in sd.rows), default=0)
                for r_i in range(nrows):
                    row_vals = sd.rows[r_i]
                    row_styles = sd.styles[r_i] if r_i < len(sd.styles) else []
                    for c_i in range(ncols):
                        cell = ws.cell(row=r_i + 1, column=c_i + 1)
                        cell.value = row_vals[c_i] if c_i < len(row_vals) else None
                        if c_i < len(row_styles) and row_styles[c_i]:
                            _apply_style(cell, row_styles[c_i])
                if nrows > 0 and ws.max_row > nrows:
                    ws.delete_rows(nrows + 1, ws.max_row - nrows)
                if ncols > 0 and ws.max_column > ncols:
                    ws.delete_cols(ncols + 1, ws.max_column - ncols)
                for col_letter, dim in ws.column_dimensions.items():
                    try: idx = column_index_from_string(col_letter) - 1
                    except (ValueError, TypeError): continue
                    dim.hidden = idx in sd.hidden_cols
                for idx in sd.hidden_cols:
                    ws.column_dimensions[get_column_letter(idx + 1)].hidden = True
                ws.sheet_state = "hidden" if sd.hidden else "visible"
                # NOTE: tab color is intentionally NOT written to the file
            # Atomic save: write to a temp file in the same folder, then
            # rename into place. If the process is killed mid-write
            # (Windows shutdown, crash, power loss), the ORIGINAL file
            # is untouched instead of being left truncated/corrupted.
            tgt = Path(self.file_path)
            tmp = tgt.with_name(tgt.stem + ".__saving__" + tgt.suffix)
            try:
                self.workbook.save(str(tmp))
                # os.replace is atomic on the same volume; retry briefly
                # in case Windows AV / Search Indexer holds the file.
                for _attempt in range(3):
                    try:
                        os.replace(str(tmp), str(tgt))
                        break
                    except PermissionError:
                        if _attempt == 2:
                            raise
                        import time as _t
                        _t.sleep(0.15)
            except Exception:
                try: tmp.unlink()
                except Exception: pass
                raise
            self.dirty = False
            self._last_save = datetime.datetime.now().strftime("%H:%M:%S")
            self._toast.show_msg("Saved ✓")
            self._update_title(); self._update_status()
        except PermissionError:
            QMessageBox.critical(self, "Cannot save",
                "The file is open in Excel.\n\nClose it first and try again.")
        except Exception as e:
            log_exc()
            QMessageBox.critical(self, "Save error", str(e))

    def closeEvent(self, ev):
        # Flush any pending debounced config writes before we exit
        try:
            if self._zoom_save_timer.isActive():
                self._zoom_save_timer.stop()
                self._persist_zoom()
            if self._col_save_timer.isActive():
                self._col_save_timer.stop()
                self._flush_col_widths()
        except Exception:
            log_exc()
        # During Windows shutdown / logout: don't show a modal — we
        # already saved in commitDataRequest. Showing one here makes
        # Windows force-kill the app (mid-write) which risks data loss
        # and dirty NTFS state.
        if self._session_ending:
            ev.accept()
            return
        if self.dirty and self.config.get("confirm_exit_unsaved", True):
            r = QMessageBox.question(
                self, "Unsaved changes",
                "You have unsaved changes. Save before closing?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            if r == QMessageBox.Save:
                self._save_file()
                if self.dirty: ev.ignore(); return
            elif r == QMessageBox.Cancel:
                ev.ignore(); return
        ev.accept()

    # ---------------------------------------------------------- tab colors (UI only)
    def _tab_default_bg(self, is_selected):
        t = THEMES.get(self.config.get("theme", "midnight"),
                       THEMES["midnight"])
        return QColor(t["panel"] if is_selected else t["bg_alt"])

    def _tab_default_accent(self):
        return QColor(self.config.get("accent_color", "#3b82f6"))

    def _tab_theme(self):
        return THEMES.get(self.config.get("theme", "midnight"),
                          THEMES["midnight"])

    def _tab_bg_alpha(self):
        return int(self.config.get("tab_bg_alpha", 0) or 0)

    def _tab_accent_height(self):
        return int(self.config.get("tab_accent_height", 3) or 3)

    def _tab_default_border(self):
        t = THEMES.get(self.config.get("theme", "midnight"),
                       THEMES["midnight"])
        return QColor(t["border"])

    def _tab_default_text(self, is_selected):
        t = THEMES.get(self.config.get("theme", "midnight"),
                       THEMES["midnight"])
        return QColor(t["text"] if is_selected else t["text_dim"])

    def _tab_color_key(self, sheet_name):
        if not self.file_path: return None
        return f"{Path(self.file_path).name}::{sheet_name}"

    def _get_tab_color(self, sheet_name):
        k = self._tab_color_key(sheet_name)
        if not k: return None
        return self.config.get("tab_colors", {}).get(k)

    def _store_tab_color(self, sheet_name, color):
        k = self._tab_color_key(sheet_name)
        if not k: return
        tc = self.config.setdefault("tab_colors", {})
        if color: tc[k] = color
        else: tc.pop(k, None)
        save_config(self.config)

    def _get_tab_bg_color(self, sheet_name):
        k = self._tab_color_key(sheet_name)
        if not k: return None
        return self.config.get("tab_bg_colors", {}).get(k)

    def _store_tab_bg_color(self, sheet_name, color):
        k = self._tab_color_key(sheet_name)
        if not k: return
        tc = self.config.setdefault("tab_bg_colors", {})
        if color: tc[k] = color
        else: tc.pop(k, None)
        save_config(self.config)

    def _get_tab_accent_color(self, sheet_name):
        k = self._tab_color_key(sheet_name)
        if not k: return None
        return self.config.get("tab_accents", {}).get(k)

    def _store_tab_accent_color(self, sheet_name, color):
        k = self._tab_color_key(sheet_name)
        if not k: return
        tc = self.config.setdefault("tab_accents", {})
        if color: tc[k] = color
        else: tc.pop(k, None)
        save_config(self.config)

    # ---------------------------------------------------------- sheet switching
    def _on_tab_changed(self, idx):
        if 0 <= idx < len(self.sheet_order): self._switch_sheet(idx)
        self._update_tab_text_colors()

    def _switch_sheet(self, idx):
        name = self.sheet_order[idx]
        self.current_sheet = name
        if name not in self.models:
            m = SheetModel(self, self.sheets[name])
            p = FilterProxy(); p.setSourceModel(m)
            self.models[name] = m; self.proxies[name] = p
        proxy = self.proxies[name]
        proxy.set_filter(self.search.text())

        self._restoring = True
        self.view.setModel(proxy)
        self.frozen_view.setModel(proxy)
        self._restoring = False

        try: self.view.selectionModel().currentChanged.disconnect()
        except Exception: pass
        try: self.frozen_view.selectionModel().currentChanged.disconnect()
        except Exception: pass
        try: self.view.selectionModel().selectionChanged.disconnect()
        except Exception: pass
        try: self.frozen_view.selectionModel().selectionChanged.disconnect()
        except Exception: pass
        self.view.selectionModel().currentChanged.connect(self._sync_from_main)
        self.frozen_view.selectionModel().currentChanged.connect(self._sync_from_frozen)
        self.view.selectionModel().selectionChanged.connect(
            self._on_selection_changed)
        self.frozen_view.selectionModel().selectionChanged.connect(
            self._on_selection_changed)

        self._apply_col_widths()
        self._apply_row_height()
        self._apply_wrap()
        self._apply_column_visibility()
        self._apply_headers_visibility()
        # Sync sort state with this sheet's own proxy (sorting is per-sheet)
        sc = proxy.sortColumn()
        self._sort_col = sc if sc is not None and sc >= 0 else -1
        self._sort_desc = (proxy.sortOrder() == Qt.DescendingOrder)
        self._update_sort_indicator()
        self._update_status()

    def _sync_from_main(self, cur, prev):
        if self._syncing or not cur.isValid() or not self.frozen_view.isVisible():
            return
        self._syncing = True
        fi = self.frozen_view.model().index(cur.row(), 0)
        if fi.isValid(): self.frozen_view.setCurrentIndex(fi)
        self._syncing = False

    def _sync_from_frozen(self, cur, prev):
        if self._syncing or not cur.isValid() or not self.frozen_view.isVisible():
            return
        self._syncing = True
        proxy = self.view.model()
        for c in range(proxy.columnCount()):
            if not self.view.isColumnHidden(c):
                self.view.setCurrentIndex(proxy.index(cur.row(), c)); break
        self._syncing = False

    def _on_selection_changed(self, selected=None, deselected=None):
        """Force full repaint on every selection change. Without this,
        fast keyboard navigation can leave stale selection borders behind
        because the custom GridDelegate draws them on top of the default
        cell paint, and Qt's internal repaint batching doesn't always
        cover the previous cell before the next one is painted."""
        if hasattr(self, "view"):
            self.view.viewport().update()
            self.frozen_view.viewport().update()
        if hasattr(self, "_row_header"):
            self._row_header.viewport().update()

    # ---------------------------------------------------------- column widths
    def _apply_col_widths(self):
        proxy = self.view.model()
        if proxy is None: return
        self._autofitting = True
        key = f"{Path(self.file_path).name}::{self.current_sheet}" if self.file_path else ""
        saved = self.config.get("col_widths", {}).get(key, {})
        if saved:
            for cs, w in saved.items():
                c = int(cs)
                if c < proxy.columnCount():
                    self.view.setColumnWidth(c, int(w))
        elif self.config.get("autofit_on_open", True):
            self._autofit()
        # else: leave the file's own default widths (Qt sizes from content)
        self._autofitting = False

    def _autofit(self):
        proxy = self.view.model()
        if proxy is None: return
        fm = self.view.fontMetrics()
        max_w = self.config.get("max_col_width", 420)
        min_w = self.config.get("min_col_width", 70)
        nrows = min(proxy.rowCount(), 100)
        for c in range(proxy.columnCount()):
            widest = min_w
            for r in range(nrows):
                txt = proxy.data(proxy.index(r, c), Qt.DisplayRole) or ""
                w = fm.horizontalAdvance(str(txt)) + 28
                if w > widest: widest = w
            hdr = fm.horizontalAdvance(get_column_letter(c + 1)) + 34
            widest = min(max(widest, hdr), max_w)
            self.view.setColumnWidth(c, widest)

    def _save_col_width(self, idx, w, persist=True):
        if not self.file_path or not self.current_sheet: return
        if not self.config.get("remember_col_widths", True):
            return
        key = f"{Path(self.file_path).name}::{self.current_sheet}"
        self.config.setdefault("col_widths", {}).setdefault(key, {})[str(idx)] = w
        if persist:
            save_config(self.config)

    def _on_col_resized(self, idx, old, new):
        if self._autofitting: return
        if not self.current_sheet: return
        # Key by (sheet, col) so the pending batch is correct even if
        # the user switches tabs before the debounce timer fires.
        self._pending_col_sizes[(self.current_sheet, idx)] = new
        self._col_save_timer.start(400)

    def _flush_col_widths(self):
        """Write all pending column widths in a single JSON update.
        Dragging a column fires sectionResized once per pixel, so the
        old per-tick save_config() was writing the entire config file
        hundreds of times per drag."""
        if not self._pending_col_sizes:
            return
        if not self.file_path or not self.config.get("remember_col_widths", True):
            self._pending_col_sizes.clear()
            return
        fname = Path(self.file_path).name
        by_sheet = {}
        for (sheet, idx), w in self._pending_col_sizes.items():
            by_sheet.setdefault(sheet, {})[str(idx)] = w
        for sheet, cols in by_sheet.items():
            key = f"{fname}::{sheet}"
            d = self.config.setdefault("col_widths", {}).setdefault(key, {})
            d.update(cols)
        self._pending_col_sizes.clear()
        save_config(self.config)

    def _on_frozen_col_resized(self, idx, old, new):
        if idx != 0: return
        self.view.setColumnWidth(0, new)
        self.frozen_view.setFixedWidth(
            new + self.frozen_view.verticalHeader().width() + 4)
        if self._autofitting: return
        if not self.current_sheet: return
        self._pending_col_sizes[(self.current_sheet, 0)] = new
        self._col_save_timer.start(400)

    # ---------------------------------------------------------- row height / wrap
    def _row_height_eff(self):
        sd = self.sheets.get(self.current_sheet) if self.current_sheet else None
        base = (sd.overrides.get("row_height", self.config.get("row_height", 38))
                if sd else self.config.get("row_height", 38))
        return max(16, int(round(base * self._zoom)))

    def _apply_row_height(self):
        rh = self._row_height_eff()
        hrh_raw = int(self.config.get("header_row_height", 0))
        hrh = max(16, int(round(hrh_raw * self._zoom))) \
            if hrh_raw > 0 else rh
        is_hdr = self.config.get("first_row_is_header", True)
        for v in (self.view, self.frozen_view):
            vh = v.verticalHeader()
            vh.setSectionResizeMode(QHeaderView.Fixed)
            vh.setDefaultSectionSize(rh)
            proxy = v.model()
            if proxy:
                for r in range(proxy.rowCount()):
                    height = hrh if (r == 0 and is_hdr) else rh
                    v.setRowHeight(r, height)

    def _apply_wrap(self):
        wrap = self.config.get("wrap_text", False)
        self.view.setWordWrap(wrap)
        self.frozen_view.setWordWrap(wrap)
        if not self.view.model():
            return
        # Apply baseline row heights first
        self._apply_row_height()
        if wrap and self.view.model().rowCount() < 500:
            # Let rows grow to fit wrapped content
            self.view.resizeRowsToContents()
            self.frozen_view.resizeRowsToContents()

    # ---------------------------------------------------------- column visibility
    def _apply_column_visibility(self):
        proxy = self.view.model()
        if proxy is None: return
        sd = self.sheets.get(self.current_sheet)
        if sd is None: return
        frozen_on = self.config.get("frozen_first_col", False)
        col0_hidden = 0 in sd.hidden_cols
        show_frozen = (frozen_on and not col0_hidden and proxy.columnCount() > 0)
        if show_frozen:
            self.frozen_view.show()
            self.frozen_view.setColumnHidden(0, False)
            w = self.view.columnWidth(0) or 200
            self.frozen_view.setColumnWidth(0, w)
            self.frozen_view.setFixedWidth(
                w + self.frozen_view.verticalHeader().width() + 4)
        else:
            self.frozen_view.hide()
        for c in range(proxy.columnCount()):
            hide = (c in sd.hidden_cols) or (show_frozen and c == 0)
            self.view.setColumnHidden(c, hide)

    # ---------------------------------------------------------- search
    def _on_search(self, text):
        proxy = self.view.model()
        if proxy and hasattr(proxy, "set_filter"):
            proxy.set_filter(text)
        self._update_status()

    # ---------------------------------------------------------- context menus
    def _cell_context_menu(self, pos):
        menu = QMenu(self)
        # Edit group
        menu.addAction("Cut", self._on_cut)
        menu.addAction("Copy", self._on_copy)
        menu.addAction("Paste", self._on_paste)
        menu.addSeparator()
        # Cell group
        menu.addAction("Clear cell content", self._clear_selection)
        menu.addSeparator()
        # Row group
        menu.addAction("Add row at end", self._add_row)
        menu.addAction("Insert row above", self._insert_row)
        menu.addAction("Duplicate row", self._duplicate_row)
        menu.addAction("Delete row", self._delete_row)
        menu.addSeparator()
        # View group
        ribbon_label = ("Hide Ribbon"
                        if self.config.get("show_toolbar", True)
                        else "Show Ribbon")
        menu.addAction(
            ribbon_label,
            lambda: self._toggle_ribbon(
                not self.config.get("show_toolbar", True)),
        )
        sb_label = ("Hide Status Bar"
                    if self.config.get("show_statusbar", True)
                    else "Show Status Bar")
        menu.addAction(
            sb_label,
            lambda: self._toggle_statusbar(
                not self.config.get("show_statusbar", True)),
        )
        menu.exec(self.view.viewport().mapToGlobal(pos))

    def _col_header_context_menu(self, pos):
        proxy = self.view.model()
        if proxy is None: return
        hh = self.sender()
        if not isinstance(hh, QHeaderView):
            hh = self.view.horizontalHeader()
        col = hh.logicalIndexAt(pos)
        if col < 0: return
        sd = self.sheets.get(self.current_sheet)
        if sd is None: return

        menu = QMenu(self)
        
        # Insert / Duplicate / Delete
        menu.addAction("Insert Column Left", lambda: self._insert_col(col))
        menu.addAction("Insert Column Right", lambda: self._insert_col(col + 1))
        menu.addAction("Duplicate Column", lambda: self._duplicate_col(col))
        menu.addAction("Delete Column", lambda: self._delete_col(col))
        menu.addSeparator()
        
        # Width Adjustments
        menu.addAction("Auto-Fit Column Width", lambda: self._autofit_col(col))
        menu.addAction("Set Column Width...", lambda: self._set_col_width(col))
        menu.addSeparator()
        
        # Hide/Unhide
        menu.addAction(f"Hide column {get_column_letter(col + 1)}",
                       lambda: self._hide_columns([col]))
        sel = self.view.selectionModel().selectedColumns()
        cols = sorted({s.column() for s in sel})
        if len(cols) > 1:
            menu.addAction(f"Hide {len(cols)} selected columns",
                           lambda: self._hide_columns(cols))
        if sd.hidden_cols:
            menu.addSeparator()
            menu.addAction(f"Show all hidden columns ({len(sd.hidden_cols)})",
                           self._show_all_columns)
            for hc in sorted(sd.hidden_cols):
                menu.addAction(f"Unhide column {get_column_letter(hc + 1)}",
                               lambda checked=False, c=hc: self._unhide_columns([c]))
        menu.exec(hh.mapToGlobal(pos))

    def _row_header_context_menu(self, pos):
        proxy = self.view.model()
        if proxy is None: return
        row = self._row_header.logicalIndexAt(pos)
        if row < 0: return
        
        # We need to convert visual row (if sorted) to source row
        # Since row headers map 1:1 with the view, we use the proxy
        proxy_idx = proxy.index(row, 0)
        src_idx = proxy.mapToSource(proxy_idx)
        src_row = src_idx.row()

        m = proxy.sourceModel()
        is_hdr = m._is_header_row(src_row)

        menu = QMenu(self)
        if not is_hdr:
            menu.addAction("Insert Row Above", lambda: self._insert_row_at(src_row))
            menu.addAction("Insert Row Below", lambda: self._insert_row_at(src_row + 1))
            menu.addAction("Duplicate Row", lambda: self._duplicate_row_at(src_row))
            menu.addAction("Delete Row", lambda: self._delete_row_at(src_row))
            menu.addSeparator()
        
        # Add an option to delete empty rows below or something if desired, 
        # but for now, just standard Excel-like actions.
        menu.exec(self._row_header.mapToGlobal(pos))

    def _tab_context_menu(self, pos):
        idx = self.tabbar.tabAt(pos)
        menu = QMenu(self)
        if 0 <= idx < len(self.sheet_order):
            name = self.sheet_order[idx]
            # Identity group
            menu.addAction("Rename sheet…",
                           lambda: self._rename_sheet(name))
            menu.addAction("Duplicate sheet",
                           lambda: self._duplicate_sheet(name))
            menu.addSeparator()
            # Appearance group
            menu.addAction("Set tab colors…",
                           lambda: self._set_tab_color(name))
            menu.addAction("Clear tab colors",
                           lambda: self._clear_tab_color(name))
            menu.addSeparator()
            # Lifecycle group
            menu.addAction("Hide sheet",
                           lambda: self._hide_sheet(name))
            menu.addAction("Delete sheet",
                           lambda: self._delete_sheet(name))
            menu.addSeparator()
        # Creation group
        menu.addAction("New sheet…", self._new_sheet)
        menu.addAction("Unhide sheets…", self._unhide_sheet_dialog)
        menu.exec(self.tabbar.mapToGlobal(pos))

    # ---------------------------------------------------------- hide/unhide columns
    def _hide_columns(self, cols):
        sd = self.sheets.get(self.current_sheet)
        if sd is None or not cols: return
        self.push_undo()
        for c in cols: sd.hidden_cols.add(int(c))
        self.mark_dirty()
        self._apply_column_visibility()
        self._toast.show_msg(f"Hid {len(cols)} column(s)")

    def _unhide_columns(self, cols):
        sd = self.sheets.get(self.current_sheet)
        if sd is None: return
        self.push_undo()
        for c in cols: sd.hidden_cols.discard(int(c))
        self.mark_dirty()
        self._apply_column_visibility()
        self._toast.show_msg(f"Unhid {len(cols)} column(s)")

    def _show_all_columns(self):
        sd = self.sheets.get(self.current_sheet)
        if sd is None or not sd.hidden_cols:
            self._toast.show_msg("No hidden columns", color="#555"); return
        n = len(sd.hidden_cols)
        self.push_undo()
        sd.hidden_cols.clear()
        self.mark_dirty()
        self._apply_column_visibility()
        self._toast.show_msg(f"Showed {n} column(s)")

    # ---------------------------------------------------------- sheet ops
    def _rebuild_tabbar(self):
        current = self.current_sheet
        self.tabbar.blockSignals(True)
        while self.tabbar.count() > 0:
            self.tabbar.removeTab(0)
        for name in self.sheet_order:
            self.tabbar.addTab(name)
        self.tabbar.blockSignals(False)
        if current in self.sheet_order:
            idx = self.sheet_order.index(current)
            self.tabbar.setCurrentIndex(idx); self._switch_sheet(idx)
        elif self.sheet_order:
            self.tabbar.setCurrentIndex(0); self._switch_sheet(0)
        else:
            self.view.setModel(None); self.frozen_view.setModel(None)
            self.current_sheet = None; self._update_status()
        self._apply_tab_colors()

    def _apply_tab_colors(self):
        """Push per-tab custom colors to the ColoredTabBar.
        Text color: user's custom color, else theme default — auto-contrast
        against a custom background happens inside paintEvent."""
        if not hasattr(self, "tabbar") or self.tabbar is None:
            return
        if hasattr(self.tabbar, "clear_all_custom"):
            self.tabbar.clear_all_custom()
        t = THEMES.get(self.config.get("theme", "midnight"), THEMES["midnight"])
        cur_idx = self.tabbar.currentIndex()
        for i in range(self.tabbar.count()):
            if i >= len(self.sheet_order):
                break
            name = self.sheet_order[i]
            sd = self.sheets.get(name)
            if not sd:
                continue
            # If user picked a custom text color, ALWAYS use it exactly.
            # Otherwise fall back to theme default.
            if sd.tab_color:
                tc = QColor(sd.tab_color)
            else:
                color_hex = t["text"] if i == cur_idx else t["text_dim"]
                tc = QColor(color_hex)
                
            if hasattr(self.tabbar, "set_tab_text_color"):
                self.tabbar.set_tab_text_color(i, tc)
                
            if hasattr(self.tabbar, "set_tab_bg"):
                self.tabbar.set_tab_bg(i, sd.tab_bg_color)
                self.tabbar.set_tab_accent(i, sd.tab_accent)

    def _hide_sheet(self, name):
        if name not in self.sheet_order: return
        if len(self.sheet_order) <= 1:
            self._toast.show_msg("Cannot hide the last visible sheet",
                                 color="#a33"); return
        self.push_undo()
        self.sheets[name].hidden = True
        self.sheet_order.remove(name)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._toast.show_msg("Sheet hidden")

    def _unhide_sheet_dialog(self):
        hidden = [n for n, sd in self.sheets.items() if sd.hidden]
        if not hidden:
            self._toast.show_msg("No hidden sheets", color="#555"); return
        dlg = QDialog(self)
        dlg.setWindowTitle("Unhide Sheets")
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel("Select sheets to unhide:"))
        checks = []
        for n in hidden:
            cb = QCheckBox(n); cb.setChecked(True)
            lay.addWidget(cb); checks.append((n, cb))
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        lay.addWidget(bb)
        if dlg.exec() != QDialog.Accepted: return
        wb_order = [ws.title for ws in self.workbook.worksheets]
        changed = False
        for n, cb in checks:
            if not cb.isChecked(): continue
            self.sheets[n].hidden = False
            if n not in self.sheet_order:
                try:
                    target = wb_order.index(n)
                    pos = sum(1 for j in range(target)
                              if wb_order[j] in self.sheet_order)
                except ValueError:
                    pos = len(self.sheet_order)
                self.sheet_order.insert(pos, n)
                changed = True
        if changed:
            self.push_undo()
            self.mark_dirty(); self._rebuild_tabbar()
            self._toast.show_msg("Sheets unhidden")

    def _rename_sheet(self, old_name):
        if old_name not in self.sheets: return
        new_name, ok = QInputDialog.getText(
            self, "Rename sheet", "New name:", text=old_name)
        if not ok: return
        new_name = new_name.strip()
        if not new_name or new_name == old_name: return
        if new_name in self.sheets:
            QMessageBox.warning(self, "Name in use",
                                "That sheet name already exists."); return
        self.push_undo()
        sd = self.sheets.pop(old_name)
        sd.name = new_name
        self.sheets[new_name] = sd
        idx = self.sheet_order.index(old_name)
        self.sheet_order[idx] = new_name
        if old_name in self.models:
            self.models[new_name] = self.models.pop(old_name)
            self.models[new_name].sheet = sd
        if old_name in self.proxies:
            self.proxies[new_name] = self.proxies.pop(old_name)
        # migrate tab color key
        old_k = self._tab_color_key(old_name)
        new_k = self._tab_color_key(new_name)
        if old_k and new_k:
            for _dict_name in ("tab_colors", "tab_bg_colors", "tab_accents"):
                d = self.config.get(_dict_name, {})
                if old_k in d:
                    d[new_k] = d.pop(old_k)
            save_config(self.config)
        if self.workbook and old_name in self.workbook.sheetnames:
            try: self.workbook[old_name].title = new_name
            except Exception: log_exc()
        self.current_sheet = new_name
        self.mark_dirty()
        self._rebuild_tabbar()
        self._toast.show_msg("Renamed")

    def _duplicate_sheet(self, name):
        if name not in self.sheets or not self.workbook: return
        new_name, ok = QInputDialog.getText(
            self, "Duplicate sheet", "New sheet name:", text=f"{name} Copy")
        if not ok: return
        new_name = new_name.strip()
        if not new_name or new_name in self.sheets: return
        self.push_undo()
        src = self.sheets[name]
        nd = SheetData(new_name)
        nd.rows = [list(r) for r in src.rows]
        nd.styles = [_copy_style_row(r) for r in src.styles]
        nd.hidden_cols = set(src.hidden_cols)
        nd.tab_color = src.tab_color
        nd.tab_bg_color = src.tab_bg_color
        nd.tab_accent = src.tab_accent
        nd.overrides = dict(src.overrides)
        self.sheets[new_name] = nd
        self.sheet_order.append(new_name)
        try:
            new_ws = self.workbook.copy_worksheet(self.workbook[name])
            new_ws.title = new_name
        except Exception:
            new_ws = self.workbook.create_sheet(new_name)
            for r_i, row in enumerate(nd.rows, start=1):
                for c_i, v in enumerate(row, start=1):
                    new_ws.cell(row=r_i, column=c_i).value = v
        # copy tab color entry
        if src.tab_color:
            self._store_tab_color(new_name, src.tab_color)
        if src.tab_bg_color:
            self._store_tab_bg_color(new_name, src.tab_bg_color)
        if src.tab_accent:
            self._store_tab_accent_color(new_name, src.tab_accent)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._toast.show_msg("Duplicated")

    def _delete_sheet(self, name):
        if name not in self.sheet_order: return
        if len(self.sheet_order) <= 1:
            self._toast.show_msg("Cannot delete the only sheet", color="#a33"); return
        r = QMessageBox.question(
            self, "Delete sheet",
            f"Delete sheet '{name}'?",
            QMessageBox.Yes | QMessageBox.No)
        if r != QMessageBox.Yes: return
        self.push_undo()
        self.sheet_order.remove(name)
        self.sheets.pop(name, None)
        self.models.pop(name, None)
        self.proxies.pop(name, None)
        self._store_tab_color(name, None)
        self._store_tab_bg_color(name, None)
        self._store_tab_accent_color(name, None)
        if self.workbook and name in self.workbook.sheetnames:
            try: del self.workbook[name]
            except Exception: log_exc()
        self.mark_dirty()
        if self.current_sheet == name:
            self.current_sheet = self.sheet_order[0] if self.sheet_order else None
        self._rebuild_tabbar()
        self._toast.show_msg("Sheet deleted")

    def _new_sheet(self):
        if not self.workbook: return
        name, ok = QInputDialog.getText(self, "New sheet", "Sheet name:", text="Sheet")
        if not ok: return
        name = name.strip()
        if not name or name in self.sheets: return
        self.push_undo()
        try:
            ws = self.workbook.create_sheet(name)
            self.sheets[name] = SheetData(name)
            self.sheet_order.append(name)
            self.mark_dirty()
            self._rebuild_tabbar()
            idx = self.sheet_order.index(name)
            self.tabbar.setCurrentIndex(idx)
            self._toast.show_msg("Sheet created")
        except Exception as e:
            QMessageBox.critical(self, "Could not create sheet", str(e))

    def _set_tab_color(self, name):
        if name not in self.sheets: return
        sd = self.sheets[name]
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Tab colors — {name}")
        dlg.setMinimumWidth(440)
        form = QFormLayout(dlg)

        def make_row(label, initial):
            edit = QLineEdit(initial or "")
            edit.setPlaceholderText("(default)")
            pick_btn = QPushButton("Pick…")
            clear_btn = QPushButton("Clear")

            def pick():
                cur = QColor(edit.text().strip() or "#3b82f6")
                c = QColorDialog.getColor(cur, dlg, label)
                if c.isValid():
                    edit.setText(c.name())

            pick_btn.clicked.connect(pick)
            clear_btn.clicked.connect(lambda: edit.setText(""))
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.addWidget(edit, 1)
            row.addWidget(pick_btn)
            row.addWidget(clear_btn)
            form.addRow(label + ":", row)
            return edit

        text_edit = make_row("Text color", sd.tab_color)
        bg_edit = make_row("Background tint",
                           self._get_tab_bg_color(name))
        accent_edit = make_row("Accent bar",
                               self._get_tab_accent_color(name))

        hint = QLabel(
            "<i>Background tint is a semi-transparent overlay so text "
            "stays readable.<br>Accent bar is the underline on the "
            "currently-selected tab.</i>")
        hint.setWordWrap(True)
        form.addRow(hint)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        form.addRow(bb)

        if dlg.exec() != QDialog.Accepted:
            return
        self.push_undo()
        sd.tab_color = text_edit.text().strip() or None
        sd.tab_bg_color = bg_edit.text().strip() or None
        sd.tab_accent = accent_edit.text().strip() or None
        self._store_tab_color(name, sd.tab_color)
        self._store_tab_bg_color(name, sd.tab_bg_color)
        self._store_tab_accent_color(name, sd.tab_accent)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._toast.show_msg("Tab colors updated")

    def _clear_tab_color(self, name):
        if name not in self.sheets: return
        sd = self.sheets[name]
        if not (sd.tab_color or sd.tab_bg_color or sd.tab_accent):
            return
        self.push_undo()
        sd.tab_color = None
        sd.tab_bg_color = None
        sd.tab_accent = None
        self._store_tab_color(name, None)
        self._store_tab_bg_color(name, None)
        self._store_tab_accent_color(name, None)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._toast.show_msg("Tab colors cleared")

    # ---------------------------------------------------------- row actions
    def _current_source_row(self):
        idx = self.view.currentIndex()
        if not idx.isValid() and self.frozen_view.isVisible():
            idx = self.frozen_view.currentIndex()
        if not idx.isValid(): return None
        proxy = self.view.model()
        return proxy.mapToSource(idx).row() if proxy else None

    def _add_row_at_end(self):
        proxy = self.view.model()
        if not proxy: return
        m = proxy.sourceModel()
        m.add_row()
        idx = proxy.index(proxy.rowCount() - 1, 0)
        self.view.setCurrentIndex(idx)
        self.view.edit(idx)
        self._update_status()

    def _add_col_at_end(self):
        proxy = self.view.model()
        if not proxy: return
        m = proxy.sourceModel()
        m.add_column()
        idx = proxy.index(0, proxy.columnCount() - 1)
        self.view.setCurrentIndex(idx)
        self.view.edit(idx)
        self._update_status()

    def _insert_col(self, col):
        proxy = self.view.model()
        if not proxy: return
        src_idx = proxy.mapToSource(proxy.index(0, col))
        proxy.sourceModel().insert_column(src_idx.column())
        self._apply_column_visibility()
        self._update_status()

    def _duplicate_col(self, col):
        proxy = self.view.model()
        if not proxy: return
        src_idx = proxy.mapToSource(proxy.index(0, col))
        proxy.sourceModel().duplicate_column(src_idx.column())
        self._apply_column_visibility()
        self._update_status()

    def _delete_col(self, col):
        proxy = self.view.model()
        if not proxy: return
        src_idx = proxy.mapToSource(proxy.index(0, col))
        proxy.sourceModel().delete_column(src_idx.column())
        self._apply_column_visibility()
        self._update_status()

    def _autofit_col(self, col):
        proxy = self.view.model()
        if not proxy: return
        fm = self.view.fontMetrics()
        max_w = self.config.get("max_col_width", 420)
        min_w = self.config.get("min_col_width", 70)
        nrows = min(proxy.rowCount(), 100)
        widest = min_w
        for r in range(nrows):
            txt = proxy.data(proxy.index(r, col), Qt.DisplayRole) or ""
            w = fm.horizontalAdvance(str(txt)) + 28
            if w > widest: widest = w
        hdr = fm.horizontalAdvance(get_column_letter(col + 1)) + 34
        widest = min(max(widest, hdr), max_w)
        self.view.setColumnWidth(col, widest)
        self._save_col_width(col, widest)

    def _set_col_width(self, col):
        cur_w = self.view.columnWidth(col)
        w, ok = QInputDialog.getInt(self, "Column Width", "Enter width in pixels:", cur_w, 20, 2000)
        if ok:
            self.view.setColumnWidth(col, w)
            self._save_col_width(col, w)

    def _insert_row_at(self, row):
        proxy = self.view.model()
        if not proxy: return
        proxy.sourceModel().add_row(at=row)
        self._update_status()

    def _duplicate_row_at(self, row):
        proxy = self.view.model()
        if not proxy: return
        proxy.sourceModel().duplicate_row(row)
        self._update_status()

    def _delete_row_at(self, row):
        proxy = self.view.model()
        if not proxy: return
        proxy.sourceModel().delete_row(row)
        self._update_status()

    def _add_row(self):
        proxy = self.view.model()
        if proxy: proxy.sourceModel().add_row(); self._update_status()

    def _insert_row(self):
        proxy = self.view.model()
        if proxy:
            r = self._current_source_row()
            proxy.sourceModel().add_row(at=r if r is not None else None)
            self._update_status()

    def _duplicate_row(self):
        proxy = self.view.model()
        if not proxy: return
        r = self._current_source_row()
        if r is None:
            self._toast.show_msg("No row selected", color="#a33"); return
        m = proxy.sourceModel()
        if m._is_header_row(r):
            self._toast.show_msg("Cannot duplicate header row", color="#a33"); return
        m.duplicate_row(r); self._update_status()

    def _delete_row(self):
        proxy = self.view.model()
        if not proxy: return
        r = self._current_source_row()
        if r is None:
            self._toast.show_msg("No row selected", color="#a33"); return
        m = proxy.sourceModel()
        if m._is_header_row(r):
            self._toast.show_msg("Cannot delete header row", color="#a33"); return
        if self.config.get("confirm_delete", False):
            reply = QMessageBox.question(
                self, "Delete row",
                f"Delete row {r + 1}?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        m.delete_row(r)
        self._toast.show_msg("Row deleted  ·  Ctrl+Z to undo", color="#555")
        self._update_status()

    def _cycle_row_coloring(self):
        order = ["cycle", "alternate", "off"]
        cur = self.config.get("row_coloring", "cycle")
        nxt = order[(order.index(cur) + 1) % len(order)]
        self.config["row_coloring"] = nxt
        self.color_btn.setText(f"🎨 Colors: {nxt.title()}")
        for m in self.models.values(): m.refresh_style()
        save_config(self.config)
        self._toast.show_msg(f"Row coloring: {nxt}", color="#555")

    def _toggle_ribbon(self, checked):
        self.config["show_toolbar"] = bool(checked)
        save_config(self.config)
        self._apply_headers_visibility()
        # Keep the View-menu action in sync if toggled from context menu
        if hasattr(self, "_act_show_ribbon"):
            self._act_show_ribbon.blockSignals(True)
            self._act_show_ribbon.setChecked(bool(checked))
            self._act_show_ribbon.blockSignals(False)
        self._toast.show_msg(
            f"Ribbon {'shown' if checked else 'hidden'}", color="#555")

    def _toggle_statusbar(self, checked):
        self.config["show_statusbar"] = bool(checked)
        save_config(self.config)
        self._apply_headers_visibility()
        if hasattr(self, "_act_show_statusbar"):
            self._act_show_statusbar.blockSignals(True)
            self._act_show_statusbar.setChecked(bool(checked))
            self._act_show_statusbar.blockSignals(False)
        self._toast.show_msg(
            f"Status bar {'shown' if checked else 'hidden'}", color="#555")

    def _global_context_menu(self, pos):
        menu = QMenu(self)
        
        r_label = "Hide Ribbon" if self._toolbar.isVisible() else "Show Ribbon"
        menu.addAction(r_label, lambda: self._toggle_ribbon(not self._toolbar.isVisible()))
        
        m_label = "Hide Menu Bar" if self.menuBar().isVisible() else "Show Menu Bar"
        menu.addAction(m_label, lambda: self._toggle_menubar(not self.menuBar().isVisible()))
        
        menu.addSeparator()
        menu.addAction("⚙ Settings…", self._open_settings)
        
        if isinstance(self.sender(), QMenuBar):
            global_pos = self.menuBar().mapToGlobal(pos)
        elif isinstance(self.sender(), QToolBar):
            global_pos = self._toolbar.mapToGlobal(pos)
        else:
            global_pos = self.sender().mapToGlobal(pos)
            
        menu.exec(global_pos)

    def _toggle_menubar(self, checked):
        self.menuBar().setVisible(bool(checked))
        if hasattr(self, "_act_show_menubar"):
            self._act_show_menubar.blockSignals(True)
            self._act_show_menubar.setChecked(bool(checked))
            self._act_show_menubar.blockSignals(False)
        self._toast.show_msg(f"Menu Bar {'shown' if checked else 'hidden'}", color="#555")

    # ---------------------------------------------------------- clipboard
    def _on_copy(self):
        sel = self.view.selectionModel().selectedIndexes()
        if not sel:
            cur = self.view.currentIndex()
            if not cur.isValid(): return
            sel = [cur]
        rows = sorted({i.row() for i in sel})
        cols = sorted({i.column() for i in sel})
        proxy = self.view.model()
        lines = []
        for r in rows:
            line = []
            for c in cols:
                idx = proxy.index(r, c)
                line.append(str(proxy.data(idx, Qt.DisplayRole) or ""))
            lines.append("\t".join(line))
        QApplication.clipboard().setText("\n".join(lines))
        self._toast.show_msg(f"Copied {len(rows)}×{len(cols)}")

    def _on_cut(self):
        proxy = self.view.model()
        if not proxy: return
        sel = self.view.selectionModel().selectedIndexes()
        if not sel:
            cur = self.view.currentIndex()
            if not cur.isValid(): return
            sel = [cur]
        self._on_copy()
        self.push_undo()
        m = proxy.sourceModel()
        changed_any = False
        for idx in sel:
            src = proxy.mapToSource(idx)
            r, c = src.row(), src.column()
            if m._is_header_row(r): continue
            if r < len(m.sheet.rows) and c < len(m.sheet.rows[r]):
                if m.sheet.rows[r][c] is not None:
                    m.sheet.rows[r][c] = None
                    changed_any = True
        if changed_any:
            self.mark_dirty()
            for idx in sel:
                src = proxy.mapToSource(idx)
                m.dataChanged.emit(src, src)
            self._toast.show_msg("Cut")

    def _on_paste(self):
        proxy = self.view.model()
        if not proxy: return
        text = QApplication.clipboard().text()
        if not text: return
        cur = self.view.currentIndex()
        if not cur.isValid(): return
        m = proxy.sourceModel()
        start_src = proxy.mapToSource(cur)
        start_r, start_c = start_src.row(), start_src.column()
        lines = text.splitlines()
        if not lines: return
        self.push_undo()
        self._restoring = True
        m.beginResetModel()
        for dr, line in enumerate(lines):
            cells = line.split("\t")
            for dc, val in enumerate(cells):
                r = start_r + dr; c = start_c + dc
                while len(m.sheet.rows) <= r:
                    m.sheet.rows.append([]); m.sheet.styles.append([])
                row = m.sheet.rows[r]
                while len(row) <= c: row.append(None)
                row[c] = m._parse(val, row[c], c)
        m.endResetModel()
        self._restoring = False
        self.mark_dirty()
        self._apply_row_height()
        self._toast.show_msg(f"Pasted {len(lines)} row(s)")

    def _clear_selection(self):
        proxy = self.view.model()
        if not proxy: return
        sel = self.view.selectionModel().selectedIndexes()
        if not sel:
            cur = self.view.currentIndex()
            if cur.isValid(): sel = [cur]
        if not sel: return
        m = proxy.sourceModel()
        self.push_undo()
        self._restoring = True
        for idx in sel:
            src = proxy.mapToSource(idx)
            r, c = src.row(), src.column()
            if m._is_header_row(r): continue
            if r < len(m.sheet.rows) and c < len(m.sheet.rows[r]):
                m.sheet.rows[r][c] = None
        self._restoring = False
        self.mark_dirty()
        for idx in sel:
            src = proxy.mapToSource(idx)
            m.dataChanged.emit(src, src)
        self._toast.show_msg("Cleared")

    def _go_first_cell(self):
        proxy = self.view.model()
        if proxy and proxy.rowCount() > 0:
            self.view.setCurrentIndex(proxy.index(0, 0))
            self.view.scrollToTop()

    def _go_last_cell(self):
        proxy = self.view.model()
        if proxy and proxy.rowCount() > 0 and proxy.columnCount() > 0:
            self.view.setCurrentIndex(
                proxy.index(proxy.rowCount() - 1, proxy.columnCount() - 1))
            self.view.scrollToBottom()

    # ---------------------------------------------------------- zoom
    def _apply_zoom(self):
        for m in self.models.values(): m.refresh_style()
        self._apply_row_height()
        self._apply_wrap()
        self.view.viewport().update(); self.frozen_view.viewport().update()
        self.frozen_view.viewport().update()
        if hasattr(self, "zoom_slider"):
            self.zoom_slider.blockSignals(True)
            self.zoom_slider.setValue(int(round(self._zoom * 100)))
            self.zoom_slider.blockSignals(False)
        if hasattr(self, "zoom_label"):
            self.zoom_label.setText(f"{int(round(self._zoom * 100))}%")

    def _persist_zoom(self):
        if not self.config.get("save_zoom", True):
            return
        self.config["zoom"] = self._zoom
        save_config(self.config)
    def _on_zoom_slider(self, v):
        self._zoom = v / 100.0
        # Update the % label immediately (cheap) so the slider feels live
        if hasattr(self, "zoom_label"):
            self.zoom_label.setText(f"{int(round(self._zoom * 100))}%")
        # Debounce the expensive repaint + the disk write. Previously
        # every single pixel of the drag fired a full model refresh
        # AND a JSON disk write (~200 each per drag) which contributed
        # to Windows shutdown failures.
        self._zoom_apply_timer.start(30)   # repaint after 30 ms idle
        self._zoom_save_timer.start(400)   # persist after 400 ms idle
    def zoom_in(self):
        self._zoom = min(self._zoom * 1.15, 2.5)
        self._apply_zoom(); self._persist_zoom()
        self._toast.show_msg(f"Zoom {int(round(self._zoom*100))}%", color="#555")
    def zoom_out(self):
        self._zoom = max(self._zoom / 1.15, 0.5)
        self._apply_zoom(); self._persist_zoom()
        self._toast.show_msg(f"Zoom {int(round(self._zoom*100))}%", color="#555")
    def zoom_reset(self):
        self._zoom = 1.0
        self._apply_zoom(); self._persist_zoom()
        self._toast.show_msg("Zoom 100%", color="#555")
        
    # ---------------------------------------------------------- window events
    def resizeEvent(self, ev):
        super().resizeEvent(ev)

    def paintEvent(self, ev):
        super().paintEvent(ev)

    # ---------------------------------------------------------- search toggle
    def _toggle_search(self):
        if not hasattr(self, "_search_wrap"): return
        if self._search_wrap.isVisible():
            self.search.setFocus()
            self.search.selectAll()
        else:
            self._search_wrap.setVisible(True)
            self.search.setFocus()
            self.search.selectAll()

    def _hide_search(self):
        if not hasattr(self, "_search_wrap"): return
        if self.search.text():
            self.search.clear()   # triggers _on_search → clears filter
        self._search_wrap.setVisible(False)    


    
    # ---------------------------------------------------------- undo / redo
    def _snapshot(self):
        cur = self.current_sheet
        snap = {
            "current": cur,
            "sheet_order": list(self.sheet_order),
            "sheet_hidden": {n: sd.hidden for n, sd in self.sheets.items()},
            "sheet_tab_colors": {n: sd.tab_color for n, sd in self.sheets.items()},
            "sheet_tab_bg_colors": {n: sd.tab_bg_color for n, sd in self.sheets.items()},
            "sheet_tab_accents": {n: sd.tab_accent for n, sd in self.sheets.items()},
            "sheet_overrides": {n: dict(sd.overrides) for n, sd in self.sheets.items()},
            "current_sheet_data": None,
        }
        if cur and cur in self.sheets:
            sd = self.sheets[cur]
            snap["current_sheet_data"] = {
                "rows": [list(r) for r in sd.rows],
                "styles": [_copy_style_row(r) for r in sd.styles],
                "hidden_cols": set(sd.hidden_cols),
            }
        return snap

    def push_undo(self):
        if self._restoring: return
        if not self.view.model(): return
        self.undo_stack.append(self._snapshot())
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _restore(self, snap):
        self._restoring = True
        try:
            self.sheet_order = list(snap["sheet_order"])
            for n, sd in self.sheets.items():
                if n in snap["sheet_hidden"]:
                    sd.hidden = snap["sheet_hidden"][n]
                if n in snap["sheet_tab_colors"]:
                    sd.tab_color = snap["sheet_tab_colors"][n]
                if n in snap.get("sheet_tab_bg_colors", {}):
                    sd.tab_bg_color = snap["sheet_tab_bg_colors"][n]
                if n in snap.get("sheet_tab_accents", {}):
                    sd.tab_accent = snap["sheet_tab_accents"][n]
                if n in snap["sheet_overrides"]:
                    sd.overrides = dict(snap["sheet_overrides"][n])
            cur = snap.get("current")
            data = snap.get("current_sheet_data")
            if cur and data and cur in self.sheets:
                sd = self.sheets[cur]
                m = self.models.get(cur)
                if m:
                    m.beginResetModel()
                    sd.rows = [list(r) for r in data["rows"]]
                    sd.styles = [_copy_style_row(r) for r in data["styles"]]
                    sd.hidden_cols = set(data["hidden_cols"])
                    m.endResetModel()
                else:
                    sd.rows = [list(r) for r in data["rows"]]
                    sd.styles = [_copy_style_row(r) for r in data["styles"]]
                    sd.hidden_cols = set(data["hidden_cols"])
        finally:
            self._restoring = False

    def undo(self):
        if not self.undo_stack:
            self._toast.show_msg("Nothing to undo", color="#555"); return
        snap = self.undo_stack.pop()
        self.redo_stack.append(self._snapshot())
        self._restore(snap)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._apply_row_height(); self._apply_column_visibility()
        self._toast.show_msg("Undone", color="#555")

    def redo(self):
        if not self.redo_stack:
            self._toast.show_msg("Nothing to redo", color="#555"); return
        snap = self.redo_stack.pop()
        self.undo_stack.append(self._snapshot())
        self._restore(snap)
        self.mark_dirty()
        self._rebuild_tabbar()
        self._apply_row_height(); self._apply_column_visibility()
        self._toast.show_msg("Redone", color="#555")

    # ---------------------------------------------------------- settings
    def _open_settings(self):
        dlg = SettingsDialog(self, self.config)
        if dlg.exec() == QDialog.Accepted:
            if getattr(dlg, "_reset_requested", False):
                self._reset_all_settings()
                return
            vals = dlg.values()
            apply_all = vals.get("apply_to_all_sheets", True)
            self.config.update(vals)
            if not apply_all and self.current_sheet:
                sd = self.sheets.get(self.current_sheet)
                if sd:
                    sd.overrides["font_size"] = vals["font_size"]
                    sd.overrides["row_height"] = vals["row_height"]
                    sd.overrides["font_family"] = vals["font_family"]
            save_config(self.config)
            # Sync the View-menu checkable actions with the new config
            if hasattr(self, "_act_show_ribbon"):
                self._act_show_ribbon.blockSignals(True)
                self._act_show_ribbon.setChecked(
                    self.config.get("show_toolbar", True))
                self._act_show_ribbon.blockSignals(False)
            if hasattr(self, "_act_show_statusbar"):
                self._act_show_statusbar.blockSignals(True)
                self._act_show_statusbar.setChecked(
                    self.config.get("show_statusbar", True))
                self._act_show_statusbar.blockSignals(False)
            self._apply_theme()
            self._apply_row_height()
            self._apply_wrap()
            self._apply_column_visibility()
            self._apply_headers_visibility()
            self._apply_col_widths()
            for m in self.models.values(): m.refresh_style()
            if hasattr(self, "view"):
                self.view.viewport().update()
                self.frozen_view.viewport().update()
            self._toast.show_msg("Settings applied")
            new_path = self.config.get("file_path")
            if new_path and new_path != self.file_path and Path(new_path).exists():
                r = QMessageBox.question(
                    self, "Open new default file?",
                    f"Open the new default file now?\n{new_path}",
                    QMessageBox.Yes | QMessageBox.No)
                if r == QMessageBox.Yes: self._open_file(new_path)

    def _reset_all_settings(self):
        """Wipe the config file, reload pure defaults, and re-apply."""
        try:
            if CONFIG_FILE.exists():
                CONFIG_FILE.unlink()
        except Exception:
            log_exc()
        self.config.clear()
        self.config.update(copy.deepcopy(CONFIG))
        save_config(self.config)
        # Re-apply everything the config affects
        self._apply_theme()
        self._apply_row_height()
        self._apply_wrap()
        self._apply_column_visibility()
        self._apply_headers_visibility()
        self._apply_col_widths()
        for m in self.models.values():
            m.refresh_style()
        self._update_tab_text_colors()
        self._update_sort_indicator()
        if hasattr(self, "view"):
            self.view.viewport().update()
            self.frozen_view.viewport().update()
        self._toast.show_msg("All settings reset to defaults")

    def _reset_all_settings_and_apply(self):
        """Compatibility alias for direct call from the dialog."""
        self._reset_all_settings()


    # ---------------------------------------------------------- status / title
    def mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self._update_title(); self._update_status()

    def _update_title(self):
        name = Path(self.file_path).name if self.file_path else ""
        mark = " •" if self.dirty else ""
        self.setWindowTitle(f"FlickerView — {name}{mark}"
                            if name else "FlickerView")

    def _update_status(self):
        proxy = self.view.model()
        if not proxy:
            self.status_label.setText("No file loaded")
            if hasattr(self, "save_label"): self.save_label.setText("")
            return
        cur = self.view.currentIndex()
        r = cur.row() + 1 if cur.isValid() else 0
        total = proxy.rowCount()
        sn = self.tabbar.currentIndex() + 1
        st = self.tabbar.count()
        self.status_label.setText(f"Row {r} of {total}  ·  Sheet {sn}/{st}")
        if self.dirty:
            self.save_label.setText("● Unsaved changes")
        elif self.file_path:
            self.save_label.setText(f"✓ Saved {self._last_save}")


# ============================================================
def main():
    try:
        app = QApplication(sys.argv)
        app.setApplicationName("FlickerView")
        try:
            app.setWindowIcon(QIcon("FlickerView.ico"))
        except Exception:
            pass
        cfg = load_config()
        w = ExcelViewer(cfg)
        w.show()
        return app.exec()
    except Exception:
        log_exc()
        try:
            from PySide6.QtWidgets import QApplication as QA, QMessageBox as QMB
            if not QA.instance(): QA(sys.argv)
            QMB.critical(None, "Startup error",
                f"The app failed to start.\n\nSee log:\n{LOG}")
        except Exception: pass
        raise


if __name__ == "__main__":
    sys.exit(main())