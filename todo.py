# ==============================================================================
# APEX LEDGER PRO — Real Estate Financial Node  v2.0  [GLASSMORHPIC EDITION]
# Modern Glassmorphic Desktop Application · High-Fidelity Financial Suite
# ==============================================================================
import sys
import os
import math
import random
import sqlite3
import hashlib
import re
import shutil
import logging
import csv
import traceback
import subprocess
import json
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLineEdit, QPushButton, QLabel, QStackedWidget, QGraphicsOpacityEffect,
    QFrame, QTextEdit, QFileDialog, QScrollArea, QDialog, QComboBox,
    QCheckBox, QMessageBox, QDateEdit, QSpinBox, QInputDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QProgressBar, QTreeWidget, QTreeWidgetItem,
    QMenu, QSizePolicy, QAbstractButton, QGraphicsDropShadowEffect,
    QTableView, QCalendarWidget, QToolButton, QAbstractSpinBox,
)
from PyQt6.QtCore import (
    Qt, QTimer, QRectF, QPointF, QRegularExpression, QPropertyAnimation,
    QEasingCurve, QSequentialAnimationGroup, QPauseAnimation, QPoint, QRect,
    QDate, QThread, pyqtSignal, QSize, QObject, QEvent,
)
from PyQt6.QtGui import (
    QPainter, QColor, QPen, QRadialGradient, QLinearGradient,
    QRegularExpressionValidator, QFont, QDoubleValidator, QIntValidator, QBrush,
    QFontDatabase, QIcon, QPixmap, QPainterPath, QFontMetrics,
    QTextCharFormat, QPalette,
)

try:
    from PyQt6.QtSvg import QSvgRenderer
    HAS_QT_SVG = True
except ImportError:
    HAS_QT_SVG = False

import openpyxl
from openpyxl.styles import (
    Font as XlFont, PatternFill, Alignment as XlAlignment,
    Border as XlBorder, Side as XlSide,
)
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table as RlTable, TableStyle as RlTableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==============================================================================
# § 1  TOKENS, CONFIGURATION, LOGGING & PATHS
# ==============================================================================
APP_NAME = "Apex Ledger Pro"
APP_VERSION = "2.0"
APP_SUBTITLE = "Financial Management Node"
SCHEMA_VERSION = 8


def _config_dir():
    base = os.environ.get("XDG_CONFIG_HOME")
    if not base:
        base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(base, "ApexLedgerPro")
    os.makedirs(d, exist_ok=True)
    return d


def _config_file():
    return os.path.join(_config_dir(), "location.ini")


def _default_data_folder():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.isdir(docs):
        docs = os.path.expanduser("~")
    return os.path.join(docs, "ApexLedgerPro")


def save_data_folder(path):
    try:
        with open(_config_file(), "w", encoding="utf-8") as f:
            f.write(f"data_folder={path}\n")
    except Exception as e:
        log.error(f"Could not save data folder config: {e}")


def load_data_folder():
    p = _config_file()
    if not os.path.exists(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("data_folder="):
                    val = line.strip().split("=", 1)[1]
                    if val:
                        return val
    except Exception as e:
        log.error(f"Could not read data folder config: {e}")
    return None


def get_apex_root():
    configured = load_data_folder()
    if configured:
        try:
            os.makedirs(configured, exist_ok=True)
        except Exception:
            pass
        return configured
    legacy = os.path.join(os.path.expanduser("~"), "Desktop", "APEX")
    if os.path.isdir(legacy) and os.path.exists(os.path.join(legacy, "users.db")):
        return legacy
    default = _default_data_folder()
    try:
        os.makedirs(default, exist_ok=True)
    except Exception:
        pass
    return default


def get_db_path():
    os.makedirs(get_apex_root(), exist_ok=True)
    return os.path.join(get_apex_root(), "users.db")


def get_log_path():
    os.makedirs(get_apex_root(), exist_ok=True)
    return os.path.join(get_apex_root(), "apex.log")


def get_backup_dir():
    d = os.path.join(get_apex_root(), "backups")
    os.makedirs(d, exist_ok=True)
    return d


def get_settings_path():
    os.makedirs(get_apex_root(), exist_ok=True)
    return os.path.join(get_apex_root(), "settings.ini")


logging.basicConfig(
    filename=get_log_path(), level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("apex")


def auto_backup_database():
    src = get_db_path()
    if not os.path.exists(src): return
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(get_backup_dir(), f"users_{stamp}.db")
        shutil.copy2(src, dst)
        files = sorted(
            [os.path.join(get_backup_dir(), f) for f in os.listdir(get_backup_dir()) if f.startswith("users_")],
            key=os.path.getmtime, reverse=True,
        )
        for old in files[10:]:
            try: os.remove(old)
            except Exception: pass
        log.info(f"DB backed up to {dst}")
    except Exception as e:
        log.error(f"Backup failed: {e}")


def _pre_upgrade_backup_v2():
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='app_meta'")
        if not cur.fetchone():
            conn.close()
            return
        cur.execute("SELECT value FROM app_meta WHERE key='schema_version'")
        row = cur.fetchone()
        conn.close()
        stored = int(row[0]) if (row and row[0] and str(row[0]).isdigit()) else 0
        if stored >= SCHEMA_VERSION:
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(get_backup_dir(), f"pre_v2_{stamp}.db")
        shutil.copy2(db_path, dst)
        log.info(f"Pre-v2 upgrade backup written to {dst}")
    except Exception as e:
        log.error(f"Pre-upgrade backup failed: {e}")


def sanitize_fs_name(name):
    return re.sub(r'[\\/*?:\"<>|, ]', '_', str(name).strip()).strip('_')


def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def fmt_money(v):
    try: return f"₹{float(v):,.2f}"
    except (TypeError, ValueError): return "₹0.00"


def load_settings():
    p = get_settings_path()
    out = {}
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line:
                        k, v = line.strip().split("=", 1)
                        out[k] = v
        except Exception:
            pass
    return out


def save_settings(d):
    try:
        with open(get_settings_path(), "w", encoding="utf-8") as f:
            for k, v in d.items():
                f.write(f"{k}={v}\n")
    except Exception as e:
        log.error(f"Settings save failed: {e}")


# -----------------------------------------------------------------------------
# PDF font registration
# -----------------------------------------------------------------------------
_PDF_FONT_REGISTERED = {"regular": False, "bold": False, "semibold": False, "medium": False}


def register_pdf_fonts():
    candidates = {
        "PDF-Regular": [
            "C:/Windows/Fonts/SF-Pro-Display-Regular.otf",
            os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/SF-Pro-Display-Regular.otf"),
            "/Library/Fonts/SF-Pro-Display-Regular.otf",
            "/System/Library/Fonts/SFNSDisplay.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
        "PDF-Bold": [
            "C:/Windows/Fonts/SF-Pro-Display-Bold.otf",
            os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/SF-Pro-Display-Bold.otf"),
            "/Library/Fonts/SF-Pro-Display-Bold.otf",
            "/System/Library/Fonts/SFNSDisplay-Bold.ttf",
            "C:/Windows/Fonts/segoeuib.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/calibrib.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ],
        "PDF-Semibold": [
            "C:/Windows/Fonts/SF-Pro-Display-Semibold.otf",
            os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/SF-Pro-Display-Semibold.otf"),
            "/Library/Fonts/SF-Pro-Display-Semibold.otf",
            "C:/Windows/Fonts/segoeuisb.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/calibrib.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ],
        "PDF-Medium": [
            "C:/Windows/Fonts/SF-Pro-Display-Medium.otf",
            os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/SF-Pro-Display-Medium.otf"),
            "/Library/Fonts/SF-Pro-Display-Medium.otf",
            "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ],
    }
    for name, paths in candidates.items():
        for path in paths:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(name, path))
                    log.info(f"Registered PDF font: {name} ← {path}")
                    if name == "PDF-Regular":  _PDF_FONT_REGISTERED["regular"] = True
                    if name == "PDF-Bold":     _PDF_FONT_REGISTERED["bold"] = True
                    if name == "PDF-Semibold": _PDF_FONT_REGISTERED["semibold"] = True
                    if name == "PDF-Medium":   _PDF_FONT_REGISTERED["medium"] = True
                    break
                except Exception as e:
                    log.error(f"Font registration failed for {name} @ {path}: {e}")


def pdf_font(weight="regular"):
    if weight == "bold" and _PDF_FONT_REGISTERED["bold"]: return "PDF-Bold"
    if weight == "semibold" and _PDF_FONT_REGISTERED["semibold"]: return "PDF-Semibold"
    if weight == "medium" and _PDF_FONT_REGISTERED["medium"]: return "PDF-Medium"
    if weight == "regular" and _PDF_FONT_REGISTERED["regular"]: return "PDF-Regular"
    if weight == "bold": return "Helvetica-Bold"
    if weight == "semibold": return "Helvetica-Bold"
    if weight == "medium": return "Helvetica"
    return "Helvetica"


# ==============================================================================


def fix_stepper_margins(widget):
    if not widget:
        return
    try:
        if isinstance(widget, QAbstractSpinBox):
            widget.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        le = widget.findChild(QLineEdit)
        if le:
            le.setTextMargins(14, 0, 28, 0)
            le.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        if hasattr(widget, "lineEdit") and widget.lineEdit():
            widget.lineEdit().setTextMargins(14, 0, 28, 0)
            widget.lineEdit().setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
    except Exception:
        pass

class UniversalSpinBoxFilter(QObject):
    def eventFilter(self, watched, event):
        if event.type() in (QEvent.Type.Show, QEvent.Type.Resize):
            if isinstance(watched, QAbstractSpinBox):
                le = watched.findChild(QLineEdit)
                if le:
                    le.setTextMargins(14, 0, 26, 0)
        return super().eventFilter(watched, event)

def setup_spinbox(spin, width=None):
    if not spin:
        return
    if width:
        spin.setFixedWidth(width)
    fix_stepper_margins(spin)

def setup_combobox(combo):
    if not combo:
        return
    try:
        v = combo.view()
        if v:
            w = v.window()
            if w:
                w.setWindowFlags(Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint)
                w.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
    except Exception:
        pass

def setup_calendar_widget(cal):
    if not cal:
        return
    cal.setGridVisible(False)
    try:
        cal.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
    except Exception:
        pass

    # Shrink row heights so all 6 date rows fit comfortably inside popup window without clipping
    table = cal.findChild(QTableView)
    if table:
        table.verticalHeader().setDefaultSectionSize(22)
        table.verticalHeader().setMinimumSectionSize(18)
        table.horizontalHeader().setDefaultSectionSize(36)
        table.horizontalHeader().setMinimumSectionSize(24)
        table.setStyleSheet("""
            QHeaderView {
                background-color: #1A1A1E;
                border: none;
            }
            QHeaderView::section {
                background-color: #1A1A1E;
                color: #A6A6B0;
                border: none;
                padding: 2px 0px;
                font-size: 11px;
                font-weight: 600;
            }
            QTableView {
                background-color: #141416;
                border: none;
                outline: none;
                alternate-background-color: transparent;
            }
        """)

    prev_btn = cal.findChild(QToolButton, "qt_calendar_prevmonth")
    next_btn = cal.findChild(QToolButton, "qt_calendar_nextmonth")
    if prev_btn:
        prev_btn.setText("<")
        prev_btn.setStyleSheet("QToolButton { color: #FFFFFF; font-weight: 800; font-size: 13px; background: rgba(255,255,255,0.12); border-radius: 6px; padding: 2px 8px; } QToolButton:hover { background: rgba(255,255,255,0.25); color: #FFFFFF; }")
    if next_btn:
        next_btn.setText(">")
        next_btn.setStyleSheet("QToolButton { color: #FFFFFF; font-weight: 800; font-size: 13px; background: rgba(255,255,255,0.12); border-radius: 6px; padding: 2px 8px; } QToolButton:hover { background: rgba(255,255,255,0.25); color: #FFFFFF; }")

    fmt_workday = QTextCharFormat()
    fmt_workday.setForeground(QColor("#FFFFFF"))

    fmt_weekend = QTextCharFormat()
    fmt_weekend.setForeground(QColor("#A6A6B0"))

    for d in (Qt.DayOfWeek.Monday, Qt.DayOfWeek.Tuesday, Qt.DayOfWeek.Wednesday,
             Qt.DayOfWeek.Thursday, Qt.DayOfWeek.Friday):
        cal.setWeekdayTextFormat(d, fmt_workday)

    cal.setWeekdayTextFormat(Qt.DayOfWeek.Saturday, fmt_weekend)
    cal.setWeekdayTextFormat(Qt.DayOfWeek.Sunday, fmt_weekend)

# § 1b  GLASSMORHPIC DESIGN SYSTEM & VECTOR ICON ENGINE
# ==============================================================================
RADIUS_CARD   = 28
RADIUS_INNER  = 16
RADIUS_INPUT  = 12
RADIUS_MODAL  = 28
RADIUS_BTN    = 10
RADIUS_BADGE  = 9999

FS_MICRO      = "10px"
FS_LABEL      = "11px"
FS_HELPER     = "12px"
FS_BODY       = "13px"
FS_INPUT      = "13px"
FS_HEADING    = "14px"
FS_CARD_TITLE = "15px"
FS_SUB        = "16px"
FS_TITLE      = "20px"
FS_HERO       = "24px"
FS_MEGA       = "28px"
FS_MONEY      = "32px"

BTN_HERO      = 42
BTN_PRIMARY   = 40
BTN_SECONDARY = 36
BTN_INLINE    = 32
BTN_ICON      = 32
INPUT_H       = 42

FONT_STACK = "'Inter', 'Ubuntu', 'Noto Sans', 'Cantarell', 'DejaVu Sans', 'SF Pro Display', 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif"

ACCENT_PURPLE     = "#7B2FF7"
ACCENT_MAGENTA    = "#C2185B"
ACCENT_RED_ORANGE = "#FF3B1F"
ACCENT_ORANGE     = "#FF6A00"
ACCENT_LIME       = "#D4E600"

ACCENT_GREEN  = "#10B981"
ACCENT_AMBER  = "#FF6A00"
ACCENT_RED    = "#EF4444"
ACCENT_BLUE   = "#3B82F6"
ACCENT_SKY    = "#0EA5E9"
ACCENT_VIOLET = "#7B2FF7"
ACCENT_INDIGO = "#FFFFFF"

SVG_ICONS = {
    "dashboard": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/></svg>""",
    "owner": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>""",
    "properties": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="16" height="20" x="4" y="2" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01"/><path d="M16 6h.01"/><path d="M8 10h.01"/><path d="M16 10h.01"/><path d="M8 14h.01"/><path d="M16 14h.01"/></svg>""",
    "tenants": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>""",
    "water": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"/></svg>""",
    "billing": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>""",
    "records": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1-2.5-2.5Z"/><path d="M6 6h10"/><path d="M6 10h10"/></svg>""",
    "reports": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg>""",
    "add": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>""",
    "edit": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/></svg>""",
    "trash": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>""",
    "power": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18.36 6.64a9 9 0 1 1-12.73 0"/><line x1="12" y1="2" x2="12" y2="12"/></svg>""",
    "eye": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>""",
    "eye_off": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9.88 9.88a3 3 0 1 0 4.24 4.24"/><path d="M10.73 5.08A10.43 10.43 0 0 1 12 5c7 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"/><path d="M6.61 6.61A13.526 13.526 0 0 0 2 12s3 7 10 7a9.74 9.74 0 0 0 5.39-1.61"/><line x1="2" y1="2" x2="22" y2="22"/></svg>""",
    "copy": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/></svg>""",
    "check": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>""",
    "close": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>""",
    "folder": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/></svg>""",
    "file": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></svg>""",
    "file_excel": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h8"/><path d="M8 17h8"/><path d="M10 9h1"/></svg>""",
    "download": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>""",
    "refresh": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16"/><path d="M16 21h5v-5"/></svg>""",
    "logout": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>""",
    "dollar": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="2" x2="12" y2="22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>""",
    "calendar": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>""",
    "search": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>""",
    "filter": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>""",
    "link": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>""",
    "settings": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0-.33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>""",
    "alert": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>""",
    "info": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>""",
    "chevron_down": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>""",
}


class SVGIconManager:
    _cache = {}

    @classmethod
    def pixmap(cls, name, color="#94A3B8", size=18):
        key = (name, color, size)
        if key in cls._cache:
            return cls._cache[key]
        svg_xml = SVG_ICONS.get(name)
        pix = QPixmap(size, size)
        pix.fill(Qt.GlobalColor.transparent)

        if svg_xml and HAS_QT_SVG:
            styled_svg = svg_xml.replace('stroke="currentColor"', f'stroke="{color}"').encode("utf-8")
            renderer = QSvgRenderer(styled_svg)
            p = QPainter(pix)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            renderer.render(p)
            p.end()
        else:
            # Fallback path renderer when QtSvg is not bundled
            p = QPainter(pix)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            p.setPen(QPen(QColor(color), 1.6, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            p.setBrush(Qt.BrushStyle.NoBrush)
            s = float(size)
            if name in ("check", "ok"):
                p.drawLine(QPointF(s*0.2, s*0.5), QPointF(s*0.45, s*0.75))
                p.drawLine(QPointF(s*0.45, s*0.75), QPointF(s*0.8, s*0.25))
            elif name in ("close", "x"):
                p.drawLine(QPointF(s*0.25, s*0.25), QPointF(s*0.75, s*0.75))
                p.drawLine(QPointF(s*0.75, s*0.25), QPointF(s*0.25, s*0.75))
            elif name in ("add", "plus"):
                p.drawLine(QPointF(s*0.5, s*0.2), QPointF(s*0.5, s*0.8))
                p.drawLine(QPointF(s*0.2, s*0.5), QPointF(s*0.8, s*0.5))
            else:
                p.drawRoundedRect(QRectF(s*0.18, s*0.18, s*0.64, s*0.64), 3, 3)
            p.end()

        cls._cache[key] = pix
        return pix

    @classmethod
    def icon(cls, name, color="#94A3B8", size=18):
        return QIcon(cls.pixmap(name, color=color, size=size))


class ThemeManager(QObject):
    changed = pyqtSignal(str)
    _instance = None

    DARK = {
        "mode": "dark",
        "bg_deep":            "#080808",
        "surface":            "rgba(40, 40, 40, 0.60)",
        "card":               "rgba(40, 40, 40, 0.60)",
        "card_alt":           "rgba(30, 30, 30, 0.65)",
        "card_hover":         "rgba(50, 50, 50, 0.72)",
        "border":             "rgba(255, 255, 255, 0.18)",
        "border_specular":    "rgba(255, 255, 255, 0.28)",
        "border_hover":       "rgba(255, 255, 255, 0.36)",
        "divider":            "rgba(255, 255, 255, 0.10)",

        "accent":             "#FFFFFF",
        "accent_hover":       "#E5E5E5",
        "accent_press":       "#CCCCCC",
        "accent_soft":        "rgba(255, 255, 255, 0.12)",
        "accent_soft_2":      "rgba(255, 255, 255, 0.22)",
        "accent_on":          "#000000",

        "input_bg":           "rgba(20, 20, 20, 0.55)",
        "input_bg_hover":     "rgba(30, 30, 30, 0.65)",
        "input_bg_focus":     "rgba(25, 25, 25, 0.85)",
        "input_border":       "rgba(255, 255, 255, 0.20)",
        "input_border_hover": "rgba(255, 255, 255, 0.45)",
        "input_border_focus": "rgba(255, 255, 255, 0.95)",
        "input_text":         "#FFFFFF",
        "input_placeholder":  "#8E8E93",
        "input_disabled_bg":  "rgba(15, 15, 15, 0.40)",
        "input_disabled_tx":  "#555555",

        "text_primary":       "#FFFFFF",
        "text_secondary":     "#A6A6B0",
        "text_muted":         "#8A8A93",
        "text_faint":         "#71717A",
        "text_dim":           "#A6A6B0",
        "text_soft":          "#B8B8C2",
        "text_dim_soft":      "#9898A2",

        "nav_bg":             "rgba(30, 30, 30, 0.55)",
        "nav_btn_text":       "#A0A0A0",
        "nav_btn_text_hover": "#FFFFFF",
        "nav_btn_text_active":"#000000",
        "nav_btn_bg_hover":   "rgba(255, 255, 255, 0.08)",
        "nav_btn_bg_active":  "#FFFFFF",

        "combo_bg":           "#141414",
        "combo_border":       "rgba(255, 255, 255, 0.20)",
        "combo_sel_bg":       "rgba(255, 255, 255, 0.20)",
        "combo_sel_bg_soft":  "rgba(255, 255, 255, 0.12)",

        "calendar_bg":        "#141414",
        "calendar_alt":       "#1E1E1E",
        "calendar_sel_bg":    "#FFFFFF",
        "calendar_sel_tx":    "#000000",

        "table_bg":           "rgba(30, 30, 30, 0.50)",
        "table_border":       "rgba(255, 255, 255, 0.12)",
        "table_grid":         "rgba(255, 255, 255, 0.06)",
        "table_hdr_bg":       "rgba(35, 35, 35, 0.85)",
        "table_sel_bg":       "#2D3142",

        "scroll_handle":      "rgba(255, 255, 255, 0.20)",
        "menu_bg":            "#141414",
        "menu_border":        "rgba(255, 255, 255, 0.20)",
        "menu_sel_bg":        "rgba(255, 255, 255, 0.18)",

        "toggle_track_off":   "rgba(255, 255, 255, 0.18)",
        "toggle_track_on":    "#FFFFFF",
        "toggle_knob":        "#FFFFFF",
        "toggle_knob_on":     "#000000",

        "chip_bg":            "rgba(255, 255, 255, 0.08)",
        "chip_border":        "rgba(255, 255, 255, 0.18)",
        "chip_text":          "#D0D0D0",

        "shadow_color":       (0, 0, 0, 160),
        "shadow_color_2":     (0, 0, 0, 220),

        "info_bg":            "rgba(14, 165, 233, 0.15)",
        "info_border":        "rgba(14, 165, 233, 0.35)",
        "warn_bg":            "rgba(245, 158, 11, 0.15)",
        "warn_border":        "rgba(245, 158, 11, 0.35)",
        "danger_bg":          "rgba(239, 68, 68, 0.15)",
        "danger_border":      "rgba(239, 68, 68, 0.35)",

        # Backward-compatible aliases
        "card_bg":            "rgba(40, 40, 40, 0.60)",
        "card_bg_hover":      "rgba(50, 50, 50, 0.72)",
        "card_border":        "rgba(255, 255, 255, 0.18)",
        "card_border_hover":  "rgba(255, 255, 255, 0.36)",
        "dialog_bg":          "rgba(28, 28, 28, 0.88)",
        "dialog_border":      "rgba(255, 255, 255, 0.22)",
    }

    def __init__(self):
        super().__init__()

    @classmethod
    def instance(cls):
        if cls._instance is None:
            cls._instance = ThemeManager()
        return cls._instance

    def mode(self):
        return "dark"

    def colors(self):
        return self.DARK

    def c(self, key):
        if key not in self.DARK:
            return "#FF00FF"
        return self.DARK[key]

    def set_mode(self, mode): pass
    def toggle(self):         pass


T = ThemeManager.instance()

# § 2  DATABASE
# ==============================================================================
def init_database():
    _pre_upgrade_backup_v2()
    print(f"👉 Database: {get_db_path()}")
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()

    cur.execute("CREATE TABLE IF NOT EXISTS app_meta (key TEXT PRIMARY KEY, value TEXT)")
    cur.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, fullname TEXT NOT NULL,
        username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS profiles (
        user_id INTEGER PRIMARY KEY, owner_name TEXT NOT NULL, phone TEXT NOT NULL, email TEXT,
        address TEXT NOT NULL, city TEXT NOT NULL, state TEXT NOT NULL,
        country TEXT NOT NULL, pin TEXT NOT NULL, excel_pdf_dir TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS properties (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
        name TEXT NOT NULL, address TEXT NOT NULL, property_type TEXT NOT NULL,
        bhk TEXT, area_sqft REAL, rate_per_sqft REAL, rent REAL NOT NULL,
        water_available INTEGER DEFAULT 0, water_meter_type TEXT DEFAULT 'Common',
        electricity_available INTEGER DEFAULT 0, card_color TEXT DEFAULT '#10B981',
        owner_members INTEGER DEFAULT 1,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS tenants (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
        property_id INTEGER NOT NULL, tenant_uid TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL, phone TEXT NOT NULL, permanent_address TEXT NOT NULL,
        security_deposit REAL NOT NULL, monthly_rent REAL NOT NULL,
        entry_date TEXT NOT NULL, exit_date TEXT,
        occupants_count INTEGER NOT NULL DEFAULT 1, is_active INTEGER NOT NULL DEFAULT 1,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(property_id) REFERENCES properties(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS tenant_bills (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
        tenant_id INTEGER NOT NULL, billing_month TEXT NOT NULL, bill_date TEXT NOT NULL,
        base_rent REAL NOT NULL,
        elec_prev REAL DEFAULT 0, elec_pres REAL DEFAULT 0, elec_units REAL DEFAULT 0,
        elec_rate REAL DEFAULT 0, elec_total REAL DEFAULT 0,
        water_prev REAL DEFAULT 0, water_pres REAL DEFAULT 0,
        water_units REAL DEFAULT 0, water_rate REAL DEFAULT 0, water_total REAL DEFAULT 0,
        discount_total REAL DEFAULT 0, prev_due REAL DEFAULT 0,
        total_payable REAL NOT NULL, amount_paid REAL DEFAULT 0,
        balance_remaining REAL NOT NULL,
        payment_status TEXT DEFAULT 'Unpaid', settled_status TEXT DEFAULT 'unpaid',
        guests_count INTEGER DEFAULT 0, owner_members_snapshot INTEGER DEFAULT 1,
        campus_persons_snapshot INTEGER DEFAULT 0, remarks TEXT,
        common_note TEXT, water_note TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(tenant_id) REFERENCES tenants(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS tenant_transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL,
        tenant_id INTEGER NOT NULL, property_id INTEGER NOT NULL,
        txn_date TEXT NOT NULL, txn_type TEXT NOT NULL, description TEXT,
        amount REAL NOT NULL, payment_method TEXT, reference_no TEXT,
        linked_bill_id INTEGER, affects_deposit INTEGER DEFAULT 0, notes TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
        FOREIGN KEY(tenant_id) REFERENCES tenants(id) ON DELETE CASCADE,
        FOREIGN KEY(property_id) REFERENCES properties(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS campus_water_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        billing_month TEXT NOT NULL,
        reading_date TEXT NOT NULL,
        prev_reading REAL NOT NULL,
        pres_reading REAL NOT NULL,
        rate_elec REAL NOT NULL,
        total_bill REAL NOT NULL,
        total_persons INTEGER NOT NULL,
        rate_per_person REAL NOT NULL,
        breakdown_json TEXT,
        remarks TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(user_id, billing_month),
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS edit_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        tenant_id INTEGER,
        txn_id INTEGER,
        action TEXT NOT NULL,
        old_values TEXT,
        new_values TEXT,
        event_time TEXT DEFAULT CURRENT_TIMESTAMP
    )""")

    migrations = [
        "ALTER TABLE properties ADD COLUMN water_meter_type TEXT DEFAULT 'Common'",
        "ALTER TABLE properties ADD COLUMN owner_members INTEGER NOT NULL DEFAULT 1",
        "ALTER TABLE profiles ADD COLUMN owner_name TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE profiles ADD COLUMN excel_pdf_dir TEXT",
        "ALTER TABLE tenants ADD COLUMN exit_date TEXT",
        "ALTER TABLE tenant_bills ADD COLUMN water_prev REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN water_pres REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN water_units REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN water_rate REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN water_total REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN discount_total REAL DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN guests_count INTEGER DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN settled_status TEXT DEFAULT 'unpaid'",
        "ALTER TABLE tenant_bills ADD COLUMN owner_members_snapshot INTEGER DEFAULT 1",
        "ALTER TABLE tenant_bills ADD COLUMN campus_persons_snapshot INTEGER DEFAULT 0",
        "ALTER TABLE tenant_bills ADD COLUMN common_note TEXT",
        "ALTER TABLE tenant_bills ADD COLUMN water_note TEXT",
    ]
    for sql in migrations:
        try: cur.execute(sql)
        except sqlite3.OperationalError: pass

    cur.execute("""UPDATE tenant_bills
                   SET total_payable = ROUND(total_payable, 2),
                       amount_paid = ROUND(amount_paid, 2),
                       balance_remaining = ROUND(balance_remaining, 2)""")
    cur.execute("""UPDATE tenant_bills
                   SET settled_status='closed', payment_status='Paid', balance_remaining=0
                   WHERE (ROUND(balance_remaining, 2) <= 0.009 OR (ROUND(amount_paid, 2) >= ROUND(total_payable, 2) AND total_payable > 0))
                     AND settled_status NOT IN ('closed', 'closed_manual')""")
    cur.execute("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('schema_version', ?)", (str(SCHEMA_VERSION),))
    conn.commit(); conn.close()


DEBIT_TYPES  = {"bill", "damage_charge", "other_debit"}
CREDIT_TYPES = {"payment", "advance", "other_credit", "deposit_adjust"}


def log_edit(user_id, tenant_id, txn_id, action, old_values=None, new_values=None):
    try:
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""INSERT INTO edit_history (user_id, tenant_id, txn_id, action, old_values, new_values)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (user_id, tenant_id, txn_id, action,
                     json.dumps(old_values) if old_values else None,
                     json.dumps(new_values) if new_values else None))
        conn.commit(); conn.close()
    except Exception as e:
        log.error(f"edit_history write failed: {e}")


def get_tenant_balance(tenant_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("SELECT txn_type, SUM(amount) FROM tenant_transactions WHERE tenant_id=? GROUP BY txn_type", (tenant_id,))
    rows = cur.fetchall(); conn.close()
    debits = credits = 0.0
    for t, s in rows:
        s = s or 0.0
        if t in DEBIT_TYPES: debits += s
        elif t in CREDIT_TYPES: credits += s
    return debits - credits


def fetch_campus_water_record(user_id, billing_month):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT id, reading_date, prev_reading, pres_reading, rate_elec,
                          total_bill, total_persons, rate_per_person, breakdown_json, remarks
                   FROM campus_water_records
                   WHERE user_id=? AND billing_month=?""", (user_id, billing_month))
    row = cur.fetchone(); conn.close()
    if not row: return None
    return {
        "id": row[0], "reading_date": row[1], "prev_reading": row[2], "pres_reading": row[3],
        "rate_elec": row[4], "total_bill": row[5], "total_persons": row[6],
        "rate_per_person": row[7], "breakdown_json": row[8], "remarks": row[9],
    }


# ==============================================================================
# § 3  FILESYSTEM
# ==============================================================================
def get_output_root(user_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("SELECT excel_pdf_dir FROM profiles WHERE user_id=?", (user_id,))
    row = cur.fetchone(); conn.close()
    if row and row[0] and os.path.isdir(row[0]):
        return row[0]
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if os.path.isdir(desktop):
        return desktop
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if os.path.isdir(docs):
        return docs
    return os.path.expanduser("~")


def get_location_dir(user_id, address):
    return os.path.join(get_output_root(user_id), f"Location_{sanitize_fs_name(address)}")


def get_property_dir(user_id, address, property_name):
    return os.path.join(get_location_dir(user_id, address), f"Property_{sanitize_fs_name(property_name)}")


def get_tenant_dir(user_id, address, property_name, tenant_name, tenant_uid):
    return os.path.join(get_property_dir(user_id, address, property_name),
                        f"Tenant_{sanitize_fs_name(tenant_name + '_' + tenant_uid)}")


def get_bill_dir(user_id, address, property_name, tenant_name, tenant_uid, year, month_str, kind):
    base = get_tenant_dir(user_id, address, property_name, tenant_name, tenant_uid)
    m_clean = str(month_str).replace('_', ' ').strip()
    if not m_clean.endswith("- Bill"):
        month_folder = f"{m_clean} - Bill"
    else:
        month_folder = m_clean

    sub_map = {
        "Generated": "Generated Bill",
        "Paid": "Paid Receipts",
        "Receipts": "Paid Receipts",
        "Unpaid": "Unpaid Bills",
        "Final": "Final Settlement",
    }
    sub_folder = sub_map.get(kind, kind)
    return os.path.join(base, str(year), month_folder, sub_folder)


def get_master_water_path(user_id):
    return os.path.join(get_output_root(user_id), "Master_Water_Ledger.xlsx")


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


# ==============================================================================
# § 4  EXCEL WRITERS
# ==============================================================================
def _xl_border(color="94A3B8"):
    s = XlSide(style="thin", color=color)
    return XlBorder(left=s, right=s, top=s, bottom=s)


def _safe_save_workbook(wb, target_path):
    try:
        wb.save(target_path)
        return
    except (PermissionError, OSError) as e:
        if os.path.exists(target_path):
            try:
                backup = target_path + ".locked_" + datetime.now().strftime("%H%M%S")
                os.replace(target_path, backup)
                wb.save(target_path)
                return
            except Exception:
                raise RuntimeError(
                    f"Cannot write to '{os.path.basename(target_path)}'.\n\n"
                    f"The file appears to be open in another program (likely Excel).\n"
                    f"Please close it and try again.\n\nOriginal error: {e}"
                )
        raise


def _xl_font(size=11, bold=False, color="0F172A"):
    try:
        return XlFont(name="SF Pro Display", size=size, bold=bold, color=color)
    except Exception:
        return XlFont(name="Calibri", size=size, bold=bold, color=color)


def rebuild_tenant_ledger(user_id, tenant_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT t.name, t.tenant_uid, p.name, p.address,
                          t.security_deposit, t.entry_date, t.exit_date, t.is_active
                   FROM tenants t JOIN properties p ON t.property_id = p.id
                   WHERE t.id=?""", (tenant_id,))
    trow = cur.fetchone()
    if not trow:
        conn.close(); return None
    (tenant_name, tenant_uid, prop_name, prop_address,
     sec_deposit, entry_date, exit_date, is_active) = trow

    cur.execute("""SELECT id, billing_month, bill_date, base_rent,
                          elec_prev, elec_pres, elec_rate, elec_total,
                          water_total, discount_total, prev_due, total_payable,
                          amount_paid, balance_remaining, settled_status, remarks,
                          common_note, water_note
                   FROM tenant_bills
                   WHERE tenant_id=?
                   ORDER BY bill_date ASC, id ASC""", (tenant_id,))
    bills = cur.fetchall()

    payments_by_bill = {}
    for b in bills:
        bill_id = b[0]
        cur.execute("""SELECT txn_date, amount, payment_method, reference_no, notes
                       FROM tenant_transactions
                       WHERE linked_bill_id=? AND txn_type IN ('payment','advance')
                       ORDER BY txn_date ASC, id ASC""", (bill_id,))
        payments_by_bill[bill_id] = cur.fetchall()

    cur.execute("""SELECT txn_date, amount, payment_method, reference_no, notes
                   FROM tenant_transactions
                   WHERE tenant_id=? AND txn_type IN ('payment','advance')
                     AND linked_bill_id IS NULL
                   ORDER BY txn_date ASC, id ASC""", (tenant_id,))
    unlinked_payments = cur.fetchall()

    cur.execute("""SELECT txn_date, txn_type, description, amount, payment_method, reference_no, notes
                   FROM tenant_transactions
                   WHERE tenant_id=? AND (affects_deposit=1 OR txn_type IN ('damage_charge', 'deposit_adjust', 'deposit_refund'))
                   ORDER BY txn_date ASC, id ASC""", (tenant_id,))
    settlement_txns = cur.fetchall()

    conn.close()

    file_path = os.path.join(
        get_tenant_dir(user_id, prop_address, prop_name, tenant_name, tenant_uid),
        "Tenant_Master_Ledger.xlsx"
    )
    ensure_dir(os.path.dirname(file_path))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tenant Ledger"
    ws.views.sheetView[0].showGridLines = True

    thin = _xl_border()
    f_title = _xl_font(size=12, bold=True, color="FFFFFF")
    f_hdr   = _xl_font(size=10, bold=True, color="0F172A")
    f_data  = _xl_font(size=10, bold=False, color="0F172A")
    f_calc  = _xl_font(size=10, bold=True, color="0F766E")
    f_due   = _xl_font(size=10, bold=True, color="B91C1C")
    f_pay   = _xl_font(size=10, bold=False, color="1E40AF")
    f_paid  = _xl_font(size=10, bold=True, color="15803D")
    f_part  = _xl_font(size=10, bold=True, color="B45309")
    f_unp   = _xl_font(size=10, bold=True, color="B91C1C")
    f_note  = _xl_font(size=9, bold=False, color="334155")

    ws.merge_cells("A1:S1")
    ws["A1"] = f"APEX LEDGER PRO — {tenant_name.upper()} | {tenant_uid}"
    ws["A1"].font = f_title
    ws["A1"].fill = PatternFill("solid", fgColor="0F172A")
    ws["A1"].alignment = XlAlignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    headers = ["Month","Bill Date","Base Rent","Elec Prev","Elec Pres","Units","Rate","Elec Bill",
               "Water Bill","Discount","Prev Due","Total","Paid","Pay Date","Balance","Remarks","Status",
               "Common Note","Water Note"]
    ws.row_dimensions[3].height = 40
    for ci, t in enumerate(headers, start=1):
        c = ws.cell(row=3, column=ci, value=t)
        c.font = f_hdr
        c.fill = PatternFill("solid", fgColor="E2E8F0")
        c.border = thin
        c.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)

    row = 4
    sec_deposit = float(sec_deposit or 0.0)
    entry_date = str(entry_date or "").strip()
    exit_date = str(exit_date or "").strip()
    is_active = bool(is_active)

    # 1. Opening Security Deposit Row
    if sec_deposit > 0 or entry_date:
        ws.cell(row=row, column=1,  value="Opening")
        ws.cell(row=row, column=2,  value=entry_date)
        for c in range(3, 12):
            ws.cell(row=row, column=c, value="—")
        ws.cell(row=row, column=12, value=round(sec_deposit, 2)).number_format = '"₹"#,##0.00'
        ws.cell(row=row, column=13, value=round(sec_deposit, 2)).number_format = '"₹"#,##0.00'
        ws.cell(row=row, column=14, value=entry_date)
        ws.cell(row=row, column=15, value=0.0).number_format = '"₹"#,##0.00'
        ws.cell(row=row, column=16, value=f"Security Deposit Received at Entry ({fmt_money(sec_deposit)})")
        ws.cell(row=row, column=17, value="Deposit Held")
        ws.cell(row=row, column=18, value=f"Initial security deposit received at lease start ({entry_date})")
        ws.cell(row=row, column=19, value="—")

        ws.row_dimensions[row].height = 24
        for col in range(1, 20):
            cell = ws.cell(row=row, column=col)
            cell.border = thin
            cell.fill = PatternFill("solid", fgColor="F0FDF4")
            if col in [12, 13, 15]:
                cell.font = f_paid
            elif col == 17:
                cell.font = _xl_font(size=10, bold=True, color="0F766E")
            elif col in [16, 18, 19]:
                cell.font = f_note
            else:
                cell.font = f_data

            if col in [1, 2, 14, 17]:
                cell.alignment = XlAlignment(horizontal="center", vertical="center")
            elif col in [16, 18, 19]:
                cell.alignment = XlAlignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = XlAlignment(horizontal="right", vertical="center")

        row += 1

    # 2. Monthly Bills & Linked Payments
    for b in bills:
        (bid, b_month, b_date, base_rent,
         eprev, epres, erate, e_total,
         w_total, disc, prev_due, total_payable,
         amt_paid, bal_remaining, settled, remarks,
         common_note, water_note) = b

        date_totals = {}
        for p in payments_by_bill.get(bid, []):
            pdate, pamt, pmethod, pref, pnotes = p
            key = pdate
            if key not in date_totals:
                date_totals[key] = {"amount": 0.0, "method": pmethod or "", "ref": pref or "", "notes": pnotes or ""}
            date_totals[key]["amount"] += float(pamt or 0.0)
            if pmethod: date_totals[key]["method"] = pmethod
            if pref: date_totals[key]["ref"] = pref
            if pnotes: date_totals[key]["notes"] = pnotes

        payment_dates = sorted(date_totals.keys())
        block_rows = max(1, len(payment_dates))
        start_row = row
        end_row = row + block_rows - 1

        ws.cell(row=start_row, column=1,  value=b_month)
        ws.cell(row=start_row, column=2,  value=b_date)
        ws.cell(row=start_row, column=3,  value=round(base_rent or 0, 2))
        ws.cell(row=start_row, column=4,  value=round(eprev or 0, 2))
        ws.cell(row=start_row, column=5,  value=round(epres or 0, 2))
        ws.cell(row=start_row, column=6,  value=f"=ROUND(E{start_row}-D{start_row}, 2)")
        ws.cell(row=start_row, column=7,  value=round(erate or 0, 2))
        ws.cell(row=start_row, column=8,  value=f"=ROUND(F{start_row}*G{start_row}, 2)")
        ws.cell(row=start_row, column=9,  value=round(w_total or 0, 2))
        ws.cell(row=start_row, column=10, value=round(disc or 0, 2))
        ws.cell(row=start_row, column=11, value=round(prev_due or 0, 2))
        ws.cell(row=start_row, column=12, value=f"=ROUND(C{start_row}+H{start_row}+I{start_row}+K{start_row}-J{start_row}, 2)")

        running_balance = round(float(total_payable or 0.0), 2)
        if payment_dates:
            for i, d in enumerate(payment_dates):
                r = start_row + i
                p = date_totals[d]
                running_balance = round(running_balance - p["amount"], 2)
                ws.cell(row=r, column=13, value=round(p["amount"], 2)).number_format = '"₹"#,##0.00'
                ws.cell(row=r, column=14, value=d)
                ws.cell(row=r, column=15, value=round(max(0.0, running_balance), 2)).number_format = '"₹"#,##0.00'
                rem_bits = []
                if p["method"]: rem_bits.append(p["method"])
                if p["ref"]:    rem_bits.append(f"Ref {p['ref']}")
                if p["notes"]:  rem_bits.append(p["notes"])
                if rem_bits:
                    ws.cell(row=r, column=16, value=" · ".join(rem_bits))
                ws.cell(row=r, column=16).font = f_pay
        else:
            ws.cell(row=start_row, column=13, value=round(0.0, 2)).number_format = '"₹"#,##0.00'
            ws.cell(row=start_row, column=14, value="")
            ws.cell(row=start_row, column=15, value=round(total_payable, 2)).number_format = '"₹"#,##0.00'

        eff_bal = round(float(bal_remaining or 0.0), 2)
        if payment_dates:
            eff_bal = min(eff_bal, round(max(0.0, running_balance), 2))

        if eff_bal <= 0.009 or (settled in ("closed", "closed_manual", "settled")):
            status_str = "Paid"; status_font = f_paid
        elif (amt_paid or 0) > 0.009 or (float(total_payable or 0.0) - eff_bal) > 0.009:
            status_str = "Partial"; status_font = f_part
        else:
            status_str = "Unpaid"; status_font = f_unp
        status_cell = ws.cell(row=end_row, column=17, value=status_str)
        status_cell.font = status_font
        status_cell.alignment = XlAlignment(horizontal="center", vertical="center")

        cn_cell = ws.cell(row=start_row, column=18, value=(common_note or ""))
        cn_cell.font = f_note
        cn_cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)
        wn_cell = ws.cell(row=start_row, column=19, value=(water_note or ""))
        wn_cell.font = f_note
        wn_cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)

        if block_rows > 1:
            for col in range(1, 13):
                ws.merge_cells(start_row=start_row, start_column=col,
                               end_row=end_row,     end_column=col)
            ws.merge_cells(start_row=start_row, start_column=18,
                           end_row=end_row,     end_column=18)
            ws.merge_cells(start_row=start_row, start_column=19,
                           end_row=end_row,     end_column=19)

        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].height = 24
            for col in range(1, 20):
                cell = ws.cell(row=r, column=col)
                cell.border = thin
                if col in [6, 8, 12] and not cell.font.bold:
                    cell.font = f_calc
                elif col == 15 and not cell.font.bold:
                    cell.font = f_due
                elif col == 13 and not cell.font.bold:
                    cell.font = f_pay
                elif col in [18, 19]:
                    cell.font = f_note
                elif not cell.font.bold:
                    cell.font = f_data
                if col in [1, 2, 14, 17]:
                    cell.alignment = XlAlignment(horizontal="center", vertical="center")
                elif col in [16, 18, 19]:
                    cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)
                else:
                    cell.alignment = XlAlignment(horizontal="right", vertical="center")

        row = end_row + 1

    # 3. Unlinked Payments
    if unlinked_payments:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=19)
        banner = ws.cell(row=row, column=1, value="UNLINKED PAYMENTS (not yet attached to a bill)")
        banner.font = _xl_font(size=10, bold=True, color="FFFFFF")
        banner.fill = PatternFill("solid", fgColor="B45309")
        banner.alignment = XlAlignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 26
        row += 1
        for pdate, pamt, pmethod, pref, pnotes in unlinked_payments:
            ws.cell(row=row, column=1, value="—")
            ws.cell(row=row, column=2, value=pdate)
            ws.cell(row=row, column=13, value=round(float(pamt or 0), 2)).number_format = '"₹"#,##0.00'
            notes_bits = [b for b in [pmethod, f"Ref {pref}" if pref else "", pnotes] if b]
            ws.cell(row=row, column=16, value=" · ".join(notes_bits))
            ws.cell(row=row, column=17, value="Unlinked")
            for col in range(1, 20):
                cell = ws.cell(row=row, column=col)
                cell.border = thin
                if col in [16, 18, 19]:
                    cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)
                elif col in [1, 2, 17]:
                    cell.alignment = XlAlignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = XlAlignment(horizontal="right", vertical="center")
            ws.row_dimensions[row].height = 24
            row += 1

    # 4. Exit / Closing Settlement Section
    if not is_active or settlement_txns or exit_date:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=19)
        banner_exit = ws.cell(row=row, column=1,
                              value=f"TENANT VACATED & LEASE TERMINATED — EXIT DATE: {exit_date or 'CLOSED'}")
        banner_exit.font = _xl_font(size=10, bold=True, color="FFFFFF")
        banner_exit.fill = PatternFill("solid", fgColor="0F172A")
        banner_exit.alignment = XlAlignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 28
        row += 1

        if settlement_txns:
            for st_date, st_type, st_desc, st_amt, st_method, st_ref, st_notes in settlement_txns:
                ws.cell(row=row, column=1,  value="Settlement")
                ws.cell(row=row, column=2,  value=st_date)
                for col in range(3, 12):
                    ws.cell(row=row, column=col, value="—")
                ws.cell(row=row, column=12, value=round(float(st_amt or 0), 2)).number_format = '"₹"#,##0.00'
                if st_type in ('deposit_adjust', 'deposit_refund'):
                    ws.cell(row=row, column=13, value=round(float(st_amt or 0), 2)).number_format = '"₹"#,##0.00'
                else:
                    ws.cell(row=row, column=13, value=0.0).number_format = '"₹"#,##0.00'
                ws.cell(row=row, column=14, value=st_date)
                ws.cell(row=row, column=15, value=0.0).number_format = '"₹"#,##0.00'
                desc_txt = st_desc or st_notes or "Deposit settlement"
                if st_ref: desc_txt += f" (Ref {st_ref})"
                ws.cell(row=row, column=16, value=desc_txt)

                if st_type == 'damage_charge':
                    st_label = "Damages Deducted"
                    st_font = _xl_font(size=10, bold=True, color="B91C1C")
                elif st_type == 'deposit_adjust':
                    st_label = "Dues Adjusted"
                    st_font = _xl_font(size=10, bold=True, color="1E40AF")
                elif st_type == 'deposit_refund':
                    st_label = "Deposit Refunded"
                    st_font = _xl_font(size=10, bold=True, color="15803D")
                else:
                    st_label = "Settled"
                    st_font = _xl_font(size=10, bold=True, color="0F766E")

                stat_c = ws.cell(row=row, column=17, value=st_label)
                stat_c.font = st_font
                ws.cell(row=row, column=18, value=f"Exit reconciliation: {st_type.replace('_', ' ').title()}")
                ws.cell(row=row, column=19, value="—")

                ws.row_dimensions[row].height = 24
                for col in range(1, 20):
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin
                    if col not in [17]:
                        cell.font = f_data
                    if col in [1, 2, 14, 17]:
                        cell.alignment = XlAlignment(horizontal="center", vertical="center")
                    elif col in [16, 18, 19]:
                        cell.alignment = XlAlignment(horizontal="left", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = XlAlignment(horizontal="right", vertical="center")
                row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        summary_cell = ws.cell(row=row, column=1,
                               value="FINAL ACCOUNT STATUS: TENANT LEFT — ALL DUES & SECURITY DEPOSIT SQUARED OFF")
        summary_cell.font = _xl_font(size=10, bold=True, color="15803D")
        summary_cell.fill = PatternFill("solid", fgColor="DCFCE7")
        summary_cell.alignment = XlAlignment(horizontal="left", vertical="center")

        bal_cell = ws.cell(row=row, column=15, value=0.0)
        bal_cell.number_format = '"₹"#,##0.00'
        bal_cell.font = _xl_font(size=10, bold=True, color="15803D")
        bal_cell.fill = PatternFill("solid", fgColor="DCFCE7")
        bal_cell.alignment = XlAlignment(horizontal="right", vertical="center")

        rem_cell = ws.cell(row=row, column=16, value="Account closed with zero outstanding balance")
        rem_cell.font = _xl_font(size=9, bold=True, color="15803D")
        rem_cell.fill = PatternFill("solid", fgColor="DCFCE7")
        rem_cell.alignment = XlAlignment(horizontal="left", vertical="center")

        stat_cell = ws.cell(row=row, column=17, value="Squared Off")
        stat_cell.font = _xl_font(size=10, bold=True, color="15803D")
        stat_cell.fill = PatternFill("solid", fgColor="DCFCE7")
        stat_cell.alignment = XlAlignment(horizontal="center", vertical="center")

        for col in [18, 19]:
            c = ws.cell(row=row, column=col, value="—")
            c.font = _xl_font(size=9, bold=False, color="15803D")
            c.fill = PatternFill("solid", fgColor="DCFCE7")
            c.alignment = XlAlignment(horizontal="center", vertical="center")

        for col in range(1, 20):
            ws.cell(row=row, column=col).border = thin

        ws.row_dimensions[row].height = 26
        row += 1

    widths = {"A":14,"B":14,"C":12,"D":12,"E":12,"F":12,"G":12,"H":12,
              "I":12,"J":12,"K":12,"L":14,"M":12,"N":14,"O":12,"P":30,"Q":18,
              "R":34,"S":34}
    for cl, w in widths.items():
        ws.column_dimensions[cl].width = w

    _safe_save_workbook(wb, file_path)
    return file_path


WATER_SHEET_HEADERS = [
    "Month", "Reading Date", "Previous Reading", "Present Reading",
    "Units Consumed", "Rate per Unit", "Campus Water Bill",
    "Campus Persons", "Per-Head Cost",
    "This Party's Family", "Guests This Month", "This Party's Water Bill",
    "Remarks", "Bill Note",
]


def _sanitize_sheet_name(name, fallback="Sheet"):
    if not name:
        name = fallback
    clean = re.sub(r'[\[\]:*?/\\]', '_', str(name).strip())
    clean = clean[:31] if len(clean) > 31 else clean
    return clean or fallback


def _ensure_water_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True
    fh = _xl_font(size=11, bold=True, color="FFFFFF")
    thin = _xl_border("000000")
    for ci, header in enumerate(WATER_SHEET_HEADERS, start=1):
        c = ws.cell(row=1, column=ci, value=header)
        c.font = fh
        c.fill = PatternFill("solid", fgColor="1E293B")
        c.border = thin
        c.alignment = XlAlignment(horizontal="center", vertical="center", wrap_text=True)
    widths = {"A": 14, "B": 14, "C": 12, "D": 12, "E": 12, "F": 12,
              "G": 14, "H": 12, "I": 12, "J": 14, "K": 12, "L": 16, "M": 32, "N": 40}
    for cl, w in widths.items():
        ws.column_dimensions[cl].width = w
    ws.row_dimensions[1].height = 52
    ws.freeze_panes = "A2"
    return ws


def _find_or_create_month_row(ws, billing_month):
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).strip() == billing_month:
            return r
    last_v = ws.cell(row=ws.max_row, column=1).value
    if last_v and "Exit" in str(last_v):
        target = ws.max_row
        ws.insert_rows(target)
    else:
        target = max(ws.max_row + 1, 2)
    ws.row_dimensions[target].height = 45
    return target


def _write_water_entry_row(ws, row, entry_date, family_count):
    thin = _xl_border("000000")
    f_entry = _xl_font(size=10, bold=True, color="15803D")
    f_note = _xl_font(size=9, bold=False, color="334155")
    fill_entry = PatternFill("solid", fgColor="F0FDF4")

    ws.row_dimensions[row].height = 32
    ws.cell(row=row, column=1, value="Tenant Entry")
    ws.cell(row=row, column=2, value=str(entry_date or "—"))
    for c in range(3, 10):
        ws.cell(row=row, column=c, value="—")
    ws.cell(row=row, column=10, value=int(family_count or 1))
    ws.cell(row=row, column=11, value=0)
    ws.cell(row=row, column=12, value=0.0).number_format = '"₹"#,##0.00'
    ws.cell(row=row, column=13, value=f"Tenant entered on {entry_date}. Initial family: {family_count}.")
    ws.cell(row=row, column=14, value="Lease commencement · Water metering starts")

    for c in range(1, len(WATER_SHEET_HEADERS) + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = thin
        cell.fill = fill_entry
        if c in [1, 2]:
            cell.font = f_entry
            cell.alignment = XlAlignment(horizontal="center", vertical="center")
        elif c in [10, 11, 12]:
            cell.font = f_entry
            cell.alignment = XlAlignment(horizontal="right", vertical="center")
        elif c in [13, 14]:
            cell.font = f_note
            cell.alignment = XlAlignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.font = f_entry
            cell.alignment = XlAlignment(horizontal="center", vertical="center")


def _write_water_row(ws, row, billing_month, reading_date,
                     prev_reading, pres_reading, rate_elec,
                     total_persons, this_party_family, guests_this_month,
                     remarks="", bill_note=""):
    thin = _xl_border("000000")
    fd = _xl_font(size=11, bold=False, color="0F172A")
    fb = _xl_font(size=11, bold=True, color="0F172A")

    ws.cell(row=row, column=1, value=billing_month)
    ws.cell(row=row, column=2, value=reading_date)
    ws.cell(row=row, column=3, value=prev_reading)
    ws.cell(row=row, column=4, value=pres_reading)
    ws.cell(row=row, column=5, value=f"=D{row}-C{row}")
    ws.cell(row=row, column=6, value=rate_elec)
    ws.cell(row=row, column=7, value=f"=E{row}*F{row}")
    ws.cell(row=row, column=8, value=total_persons)
    ws.cell(row=row, column=9, value=f"=IF(H{row}>0, G{row}/H{row}, 0)")
    ws.cell(row=row, column=10, value=this_party_family)
    ws.cell(row=row, column=11, value=guests_this_month)
    ws.cell(row=row, column=12, value=f"=IF(I{row}>0, (J{row}+K{row})*I{row}, 0)")
    ws.cell(row=row, column=13, value=remarks)
    ws.cell(row=row, column=14, value=bill_note)

    for ci in [6, 7, 9, 12]:
        ws.cell(row=row, column=ci).number_format = '"₹"#,##0.00'

    for c in range(1, len(WATER_SHEET_HEADERS) + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = thin
        if not cell.font.bold:
            cell.font = fb if c in [5, 7, 9, 12] else fd
        if c in [1, 2]:
            cell.alignment = XlAlignment(horizontal="center", vertical="center")
        elif c in [13, 14]:
            cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)
        else:
            cell.alignment = XlAlignment(horizontal="right", vertical="center")


def ensure_all_water_sheets(user_id, month_str, reading_date,
                            prev_reading, pres_reading, rate_elec,
                            owner_name, owner_family,
                            tenants_data, remarks=""):
    file_path = get_master_water_path(user_id)
    ensure_dir(os.path.dirname(file_path))

    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception:
            try: os.remove(file_path)
            except Exception: pass
            wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)
    else:
        wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)

    total_persons = owner_family + sum(int(t.get("family", 0)) for t in tenants_data)

    owner_sheet = _sanitize_sheet_name(f"Owner_{owner_name}", "Owner")
    ws_o = _ensure_water_sheet(wb, owner_sheet)
    r_o = _find_or_create_month_row(ws_o, month_str)
    _write_water_row(ws_o, r_o, month_str, reading_date,
                     prev_reading, pres_reading, rate_elec,
                     total_persons, this_party_family=owner_family,
                     guests_this_month=0, remarks=remarks, bill_note="")

    for t in tenants_data:
        t_sheet = _sanitize_sheet_name(f"Tenant_{t['name']}", f"Tenant_{t['name'][:20]}")
        ws_t = _ensure_water_sheet(wb, t_sheet)
        r_t = _find_or_create_month_row(ws_t, month_str)
        existing_note = ""
        if r_t <= ws_t.max_row:
            v = ws_t.cell(row=r_t, column=14).value
            if v: existing_note = v
        _write_water_row(ws_t, r_t, month_str, reading_date,
                         prev_reading, pres_reading, rate_elec,
                         total_persons, this_party_family=int(t.get("family", 0)),
                         guests_this_month=int(t.get("guests", 0)),
                         remarks=remarks, bill_note=existing_note)

    for sn in list(wb.sheetnames):
        ws_check = wb[sn]
        if ws_check.max_row <= 1 and ws_check.max_column <= 1 and ws_check["A1"].value in (None, ""):
            if sn not in (owner_sheet,) and not sn.startswith("Tenant_"):
                wb.remove(ws_check)

    _safe_save_workbook(wb, file_path)
    return file_path


def update_single_water_sheet_guests(user_id, month_str, party_kind, party_name,
                                      new_guests, reading_date="",
                                      prev_reading=0, pres_reading=0, rate_elec=0,
                                      owner_name="", owner_family=0, tenants_data=None,
                                      remarks="", bill_note=""):
    file_path = get_master_water_path(user_id)
    ensure_dir(os.path.dirname(file_path))

    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception:
            wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)
    else:
        wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)

    if party_kind == "owner":
        sheet_name = _sanitize_sheet_name(f"Owner_{owner_name}", "Owner")
    else:
        sheet_name = _sanitize_sheet_name(f"Tenant_{party_name}", f"Tenant_{party_name[:20]}")

    ws = _ensure_water_sheet(wb, sheet_name)

    target_row = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).strip() == month_str:
            target_row = r
            break

    if target_row is None:
        total_persons = owner_family + sum(int(t.get("family", 0)) for t in (tenants_data or []))
        this_family = owner_family if party_kind == "owner" else \
                      next((int(t["family"]) for t in (tenants_data or []) if t["name"] == party_name), 0)
        target_row = max(ws.max_row + 1, 2)
        ws.row_dimensions[target_row].height = 45
        _write_water_row(ws, target_row, month_str, reading_date,
                         prev_reading, pres_reading, rate_elec,
                         total_persons, this_party_family=this_family,
                         guests_this_month=int(new_guests),
                         remarks=remarks, bill_note=bill_note)
    else:
        cell = ws.cell(row=target_row, column=11, value=int(new_guests))
        cell.number_format = '0'
        ws.cell(row=target_row, column=5).value = f"=D{target_row}-C{target_row}"
        ws.cell(row=target_row, column=7).value = f"=E{target_row}*F{target_row}"
        ws.cell(row=target_row, column=9).value = f"=IF(H{target_row}>0, G{target_row}/H{target_row}, 0)"
        ws.cell(row=target_row, column=12).value = f"=IF(I{target_row}>0, (J{target_row}+K{target_row})*I{target_row}, 0)"
        if bill_note is not None:
            note_cell = ws.cell(row=target_row, column=14, value=bill_note)
            note_cell.alignment = XlAlignment(horizontal="left", vertical="top", wrap_text=True)

    _safe_save_workbook(wb, file_path)
    return file_path

def record_tenant_entry_in_water_ledger(user_id, tenant_name, entry_date, family_count):
    file_path = get_master_water_path(user_id)
    ensure_dir(os.path.dirname(file_path))

    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception:
            wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)
    else:
        wb = openpyxl.Workbook(); d = wb.active; wb.remove(d)

    sheet_name = _sanitize_sheet_name(f"Tenant_{tenant_name}", f"Tenant_{tenant_name[:20]}")
    ws = _ensure_water_sheet(wb, sheet_name)

    entry_row = None
    for r in range(2, ws.max_row + 1):
        v1 = ws.cell(row=r, column=1).value
        v13 = ws.cell(row=r, column=13).value
        if (v1 and "Entry" in str(v1)) or (v13 and "entered" in str(v13).lower()):
            entry_row = r
            break

    if entry_row is None:
        if ws.max_row >= 2 and ws.cell(row=2, column=1).value not in (None, ""):
            ws.insert_rows(2)
            entry_row = 2
        else:
            entry_row = 2

    _write_water_entry_row(ws, entry_row, entry_date, family_count)
    _safe_save_workbook(wb, file_path)
    return file_path


def record_tenant_exit_in_water_ledger(user_id, tenant_name, exit_date):
    file_path = get_master_water_path(user_id)
    if not os.path.exists(file_path):
        return None
    try:
        wb = openpyxl.load_workbook(file_path)
    except Exception as e:
        log.error(f"Cannot load water ledger for exit record: {e}")
        return None

    sheet_name = _sanitize_sheet_name(f"Tenant_{tenant_name}", f"Tenant_{tenant_name[:20]}")
    if sheet_name not in wb.sheetnames:
        return None

    ws = wb[sheet_name]

    for r in range(2, ws.max_row + 1):
        v1 = ws.cell(row=r, column=1).value
        v13 = ws.cell(row=r, column=13).value
        if v1 and "Exit" in str(v1):
            return file_path
        if v13 and "exited" in str(v13).lower():
            return file_path

    target_row = max(ws.max_row + 1, 2)
    ws.row_dimensions[target_row].height = 34

    thin = _xl_border("000000")
    f_exit = _xl_font(size=10, bold=True, color="B91C1C")
    f_note = _xl_font(size=9, bold=False, color="475569")
    fill_exit = PatternFill("solid", fgColor="FEE2E2")

    ws.cell(row=target_row, column=1, value="Tenant Exited")
    ws.cell(row=target_row, column=2, value=exit_date)
    for c in range(3, 12):
        ws.cell(row=target_row, column=c, value="—")
    ws.cell(row=target_row, column=12, value=0.0).number_format = '"₹"#,##0.00'
    ws.cell(row=target_row, column=13, value=f"Tenant exited on {exit_date}. Water account squared off & closed.")
    ws.cell(row=target_row, column=14, value=f"Lease terminated on {exit_date}. Final settlement complete.")

    for c in range(1, len(WATER_SHEET_HEADERS) + 1):
        cell = ws.cell(row=target_row, column=c)
        cell.border = thin
        cell.fill = fill_exit
        if c in [1, 2]:
            cell.font = f_exit
            cell.alignment = XlAlignment(horizontal="center", vertical="center")
        elif c in [13, 14]:
            cell.font = f_note
            cell.alignment = XlAlignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.font = f_exit
            cell.alignment = XlAlignment(horizontal="center", vertical="center")

    _safe_save_workbook(wb, file_path)
    return file_path

def export_transactions_csv(user_id, output_path):
    ensure_dir(os.path.dirname(output_path))
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT tt.txn_date, t.name, p.name, tt.txn_type, tt.description,
                          tt.amount, tt.payment_method, tt.reference_no, tt.notes
                   FROM tenant_transactions tt
                   JOIN tenants t ON tt.tenant_id = t.id
                   JOIN properties p ON tt.property_id = p.id
                   WHERE tt.user_id=? ORDER BY tt.txn_date ASC, tt.id ASC""", (user_id,))
    rows = cur.fetchall(); conn.close()
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Date","Tenant","Property","Type","Description","Amount","Method","Reference","Notes"])
        for r in rows:
            w.writerow(r)
    return len(rows)


# ==============================================================================
# § 5  PDF GENERATORS
# ==============================================================================
PDF_INK        = "#09090B"
PDF_INK_SOFT   = "#27272A"
PDF_MUTED      = "#71717A"
PDF_LINE       = "#E4E4E7"
PDF_LINE_SOFT  = "#F4F4F5"
PDF_AL_DARK    = "#18181B"
PDF_AL_MID     = "#71717A"
PDF_ORANGE     = "#FF6A00"
PDF_GREEN      = "#16A34A"
PDF_AMBER      = "#F59E0B"
PDF_RED        = "#DC2626"


def _pdf_money(v):
    try:
        amount = float(v)
    except (TypeError, ValueError):
        amount = 0.0
    return f"Rs.{amount:,.2f}"


def _flatten_story(story):
    flat = []
    for item in story:
        if isinstance(item, (list, tuple)):
            flat.extend(_flatten_story(item))
        else:
            flat.append(item)
    return flat


def _safe_build(doc, story):
    doc.build(_flatten_story(story))


def _pdf_styles():
    styles = getSampleStyleSheet()
    return {
        "title":  ParagraphStyle("title", parent=styles["Normal"],
                                  fontName=pdf_font("bold"), fontSize=22,
                                  textColor=colors.HexColor(PDF_INK),
                                  alignment=2, leading=26),
        "subtitle": ParagraphStyle("subtitle", parent=styles["Normal"],
                                   fontName=pdf_font("regular"), fontSize=9,
                                   textColor=colors.HexColor(PDF_MUTED),
                                   alignment=2, leading=11),
        "meta":   ParagraphStyle("meta", parent=styles["Normal"],
                                 fontName=pdf_font("regular"), fontSize=9,
                                 textColor=colors.HexColor(PDF_INK_SOFT), leading=12),
        "kv_label": ParagraphStyle("kvl", parent=styles["Normal"],
                                    fontName=pdf_font("semibold"), fontSize=9,
                                    textColor=colors.HexColor(PDF_MUTED), leading=12),
        "kv_value": ParagraphStyle("kvv", parent=styles["Normal"],
                                    fontName=pdf_font("regular"), fontSize=9.5,
                                    textColor=colors.HexColor(PDF_INK), leading=12),
        "kv_value_big": ParagraphStyle("kvbig", parent=styles["Normal"],
                                       fontName=pdf_font("bold"), fontSize=11,
                                       textColor=colors.HexColor(PDF_INK), leading=14),
        "small":  ParagraphStyle("small", parent=styles["Normal"],
                                 fontName=pdf_font("regular"), fontSize=8,
                                 textColor=colors.HexColor(PDF_MUTED), leading=10),
        "table_hdr": ParagraphStyle("th", parent=styles["Normal"],
                                    fontName=pdf_font("bold"), fontSize=9.5,
                                    textColor=colors.white, leading=12),
        "table_val": ParagraphStyle("tv", parent=styles["Normal"],
                                    fontName=pdf_font("regular"), fontSize=9.5,
                                    textColor=colors.HexColor(PDF_INK), leading=12),
        "table_sub": ParagraphStyle("ts", parent=styles["Normal"],
                                    fontName=pdf_font("regular"), fontSize=8,
                                    textColor=colors.HexColor(PDF_MUTED), leading=10),
        "note_lbl": ParagraphStyle("nl", parent=styles["Normal"],
                                   fontName=pdf_font("semibold"), fontSize=8.5,
                                   textColor=colors.HexColor(PDF_MUTED), leading=11),
        "note_txt": ParagraphStyle("nt", parent=styles["Normal"],
                                   fontName=pdf_font("regular"), fontSize=9,
                                   textColor=colors.HexColor(PDF_INK), leading=12),
        "note_hdr": ParagraphStyle("nh", parent=styles["Normal"],
                                   fontName=pdf_font("bold"), fontSize=8.5,
                                   textColor=colors.HexColor(PDF_ORANGE), leading=11),
    }


def _apex_wordmark():
    S = _pdf_styles()
    line1 = ParagraphStyle("wm1", parent=S["title"],
                           fontName=pdf_font("bold"), fontSize=18,
                           textColor=colors.HexColor(PDF_INK),
                           alignment=1, leading=22)
    line2 = ParagraphStyle("wm2", parent=S["title"],
                           fontName=pdf_font("regular"), fontSize=10,
                           textColor=colors.HexColor(PDF_MUTED),
                           alignment=1, leading=13)
    tbl = RlTable(
        [[Paragraph("<b>APEX</b> &nbsp;<font color='#FF6A00'>LEDGER</font>", line1)],
         [Paragraph("P R O", line2)]],
        colWidths=[540])
    tbl.setStyle(RlTableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (0, 0), 0),
        ("BOTTOMPADDING", (0, 1), (0, 1), 8),
        ("LINEBELOW", (0, 1), (-1, 1), 1.5, colors.HexColor(PDF_ORANGE)),
    ]))
    return tbl


def _header_band(S, title_text, subtitle_text=""):
    title_para = Paragraph(f"<b>{title_text}</b>", S["title"])
    sub_para = Paragraph(subtitle_text, S["subtitle"]) if subtitle_text else Paragraph("", S["subtitle"])
    right_stack = RlTable([[title_para], [sub_para]], colWidths=[540])
    right_stack.setStyle(RlTableStyle([
        ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return [_apex_wordmark(), Spacer(1, 12), right_stack]


def _footer_line(S, page_label=""):
    ts = datetime.now().strftime("%d %b %Y · %H:%M")
    line = f"Apex Ledger Pro v{APP_VERSION} · Generated {ts}"
    if page_label:
        line = f"{page_label}  ·  {line}"
    return RlTable(
        [[Paragraph(f"<font size=7.5 color='#94A3B8'>{line}</font>", S["small"])]],
        colWidths=[540],
        style=RlTableStyle([
            ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("LINEABOVE", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE_SOFT)),
        ]))


def _info_block(rows, label_width=90, value_width=460):
    data = []
    for k, v in rows:
        data.append([
            Paragraph(f"<b>{k}</b>", _pdf_styles()["kv_label"]),
            Paragraph(str(v), _pdf_styles()["kv_value"]),
        ])
    t = RlTable(data, colWidths=[label_width, value_width])
    t.setStyle(RlTableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _items_table_customer(items):
    S = _pdf_styles()
    rows = [[
        Paragraph("<b>Description</b>", S["table_hdr"]),
        Paragraph("<b>Total (Rs.)</b>", S["table_hdr"]),
    ]]
    for label, amount, sub in items:
        desc_cell = Paragraph(f"<b>{label}</b>", S["table_val"])
        if sub:
            desc_cell = Paragraph(
                f"<b>{label}</b><br/><font size=7.5 color='#64748B'>{sub}</font>",
                S["table_val"])
        rows.append([desc_cell, Paragraph(str(amount), S["table_val"])])
    t = RlTable(rows, colWidths=[400, 140])
    t.setStyle(RlTableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PDF_AL_DARK)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(PDF_LINE)),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
    ]))
    return t


def _items_table_owner(items):
    S = _pdf_styles()
    rows = [[
        Paragraph("<b>Description</b>", S["table_hdr"]),
        Paragraph("<b>Rate</b>", S["table_hdr"]),
        Paragraph("<b>Total (Rs.)</b>", S["table_hdr"]),
    ]]
    for label, rate, amount, sub in items:
        desc_cell = Paragraph(f"<b>{label}</b>", S["table_val"])
        if sub:
            desc_cell = Paragraph(
                f"<b>{label}</b><br/><font size=7.5 color='#64748B'>{sub}</font>",
                S["table_val"])
        rows.append([
            desc_cell,
            Paragraph(str(rate), S["table_val"]),
            Paragraph(str(amount), S["table_val"]),
        ])
    t = RlTable(rows, colWidths=[320, 90, 130])
    t.setStyle(RlTableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PDF_AL_DARK)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(PDF_LINE)),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
    ]))
    return t


def _totals_block_customer(total_str):
    S = _pdf_styles()
    t = RlTable([
        [Paragraph("", S["kv_value"]),
         Paragraph(f"<b>{total_str}</b>", S["kv_value_big"])],
    ], colWidths=[380, 160])
    t.setStyle(RlTableStyle([
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LINEABOVE", (0, 0), (-1, 0), 1.5, colors.HexColor(PDF_ORANGE)),
    ]))
    return t


def _totals_block_owner(rows):
    S = _pdf_styles()
    data = []
    for i, (lbl, val) in enumerate(rows):
        is_last = (i == len(rows) - 1)
        lbl_p = Paragraph(f"<b>{lbl}</b>" if is_last else lbl, S["kv_value_big"] if is_last else S["kv_value"])
        val_p = Paragraph(f"<b>{val}</b>" if is_last else val, S["kv_value_big"] if is_last else S["kv_value"])
        data.append([lbl_p, val_p])
    t = RlTable(data, colWidths=[380, 160])
    style = [
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LINEABOVE", (0, 0), (-1, 0), 1.5, colors.HexColor(PDF_ORANGE)),
    ]
    if len(rows) > 1:
        style.append(("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor(PDF_LINE_SOFT)))
    t.setStyle(RlTableStyle(style))
    return t


def _notes_box(common_note, water_note, discount_reason=""):
    common_note = (common_note or "").strip()
    water_note = (water_note or "").strip()
    discount_reason = (discount_reason or "").strip()
    if not common_note and not water_note and not discount_reason:
        return None

    S = _pdf_styles()

    def prep(text):
        t = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return t.replace("\n", "<br/>")

    rows = []
    if common_note:
        rows.append([
            Paragraph("<b>Rent &amp; Electricity:</b>", S["note_lbl"]),
            Paragraph(prep(common_note), S["note_txt"]),
        ])
    if water_note:
        rows.append([
            Paragraph("<b>Water:</b>", S["note_lbl"]),
            Paragraph(prep(water_note), S["note_txt"]),
        ])
    if discount_reason:
        rows.append([
            Paragraph("<b>Discount Reason:</b>", S["note_lbl"]),
            Paragraph(prep(discount_reason), S["note_txt"]),
        ])

    inner = RlTable(rows, colWidths=[110, 430])
    inner_style = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if len(rows) > 1:
        inner_style.append(("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor(PDF_LINE_SOFT)))
    inner.setStyle(RlTableStyle(inner_style))

    header = Paragraph("<b>NOTES (OWNER ONLY)</b>", S["note_hdr"])
    outer = RlTable([[header], [inner]], colWidths=[540])
    outer.setStyle(RlTableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FAFAFA")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, 0), (0, 0), 4),
        ("TOPPADDING", (0, 1), (0, 1), 2),
        ("BOTTOMPADDING", (0, 1), (0, 1), 10),
    ]))
    return outer


def generate_bill_pdfs(target_dir, data, is_paid=False):
    ensure_dir(target_dir)
    prefix = "Paid_Bill" if is_paid else "Bill"
    cust_path = os.path.join(target_dir, f"{prefix}_Customer_Copy.pdf")
    own_path  = os.path.join(target_dir, f"{prefix}_Owner_Audit.pdf")
    S = _pdf_styles()

    render_discount_customer = bool(data.get("render_discount_customer", True))
    discount_amt = float(data.get("discount_total", 0) or 0)
    discount_reason = (data.get("discount_reason") or "").strip()

    elec_units = max(0.0, data["elec_pres"] - data["elec_prev"])
    elec_sub = f"Meter: {data['elec_prev']:g} -> {data['elec_pres']:g}  ({elec_units:g} u x Rs.{data['elec_rate']:.2f})"

    doc = SimpleDocTemplate(cust_path, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "RENT INVOICE", "PAID" if is_paid else ""))
    story.append(Spacer(1, 14))

    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Property", data['property_name']),
    ]))
    story.append(Spacer(1, 10))
    cust_inv_info = [
        ("Invoice No.", f"INV-{data['bill_id']:05d}"),
        ("Invoice Date", data['bill_date']),
        ("Billing Cycle", data['billing_month']),
    ]
    if is_paid:
        paid_dt = str(data.get("payment_date") or data.get("paid_date") or "").strip()
        if paid_dt:
            cust_inv_info.append(("Paid Date", paid_dt))
    story.append(_info_block(cust_inv_info))
    story.append(Spacer(1, 16))

    cust_items = [
        ["Base Rent", _pdf_money(data['base_rent']), None],
        ["Electricity", _pdf_money(data['elec_total']), elec_sub],
        ["Water", _pdf_money(data['water_total']), None],
    ]
    if data.get("prev_due", 0) != 0:
        label = "Previous Balance" if data['prev_due'] > 0 else "Advance Credit"
        cust_items.append([label, _pdf_money(data['prev_due']), None])

    if discount_amt > 0 and render_discount_customer:
        cust_items.append([
            "Discount",
            f"-{_pdf_money(discount_amt)}",
            None,
        ])

    story.append(_items_table_customer(cust_items))
    story.append(Spacer(1, 4))
    story.append(_totals_block_customer(_pdf_money(data['total_payable'])))
    if discount_amt > 0 and render_discount_customer and discount_reason:
        story.append(Spacer(1, 14))
        reason_para = Paragraph(
            f"<b>Discount Note:</b> {discount_reason}",
            ParagraphStyle("creason", parent=S["meta"],
                           fontSize=9, leading=12,
                           textColor=colors.HexColor(PDF_INK_SOFT))
        )
        reason_box = RlTable([[reason_para]], colWidths=[540])
        reason_box.setStyle(RlTableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story.append(reason_box)

    story.append(Spacer(1, 20))
    story.append(Paragraph(
        f"<font size=9 color='#334155'>Thank you for your business. Please settle by "
        f"{data['bill_date']}.</font>", S["meta"]))
    story.append(Spacer(1, 20))
    story.append(_footer_line(S))
    _safe_build(doc, story)

    doc = SimpleDocTemplate(own_path, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "OWNER COPY / INVOICE", "PAID" if is_paid else ""))
    story.append(Spacer(1, 14))

    owner_phone = data.get('owner_phone', '—') or '—'
    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Phone", owner_phone),
        ("Property", f"{data['property_name']} — {data.get('property_address', '')}"),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Occupants", f"{data['occupants_count']} + guests {data.get('guests_count', 0)}"),
    ]))
    story.append(Spacer(1, 10))
    own_inv_info = [
        ("Invoice No.", f"INV-{data['bill_id']:05d}"),
        ("Invoice Date", data['bill_date']),
        ("Billing Cycle", data['billing_month']),
    ]
    if is_paid:
        paid_dt = str(data.get("payment_date") or data.get("paid_date") or "").strip()
        if paid_dt:
            own_inv_info.append(("Paid Date", paid_dt))
    story.append(_info_block(own_inv_info))
    story.append(Spacer(1, 16))

    if data.get("is_common_water"):
        flat_persons = int(data.get('occupants_count', 0) or 0)
        guests = int(data.get('guests_count', 0) or 0)
        if guests > 0:
            water_basis = (f"Shared pump · {flat_persons} persons in flat · "
                           f"{guests} guest{'s' if guests != 1 else ''}")
        else:
            water_basis = f"Shared pump · {flat_persons} persons in flat"
        water_rate_str = _pdf_money(data['water_rate'])
    else:
        water_basis = f"Meter: {data['water_prev']:g} -> {data['water_pres']:g}"
        water_rate_str = _pdf_money(data['water_rate'])

    elec_basis = f"Meter: {data['elec_prev']:g} -> {data['elec_pres']:g}  ({elec_units:g} u x Rs.{data['elec_rate']:.2f})"

    own_items = [
        ["Base Rent", _pdf_money(data['base_rent']), _pdf_money(data['base_rent']), "Contract rate"],
        ["Electricity", f"Rs.{data['elec_rate']:.2f}", _pdf_money(data['elec_total']), elec_basis],
        ["Water", water_rate_str, _pdf_money(data['water_total']), water_basis],
    ]
    if data.get("prev_due", 0) != 0:
        label = "Previous Balance" if data['prev_due'] > 0 else "Advance Credit"
        own_items.append([label, "—", _pdf_money(data['prev_due']), "Carried forward"])
    if discount_amt > 0:
        own_items.append([
            "Discount", "—", f"-{_pdf_money(discount_amt)}",
            ""   # sub-line removed — reason is shown in the NOTES box below
        ])

    story.append(_items_table_owner(own_items))
    story.append(Spacer(1, 4))

    notes_box = _notes_box(
        data.get("common_note", ""),
        data.get("water_note", ""),
        data.get("discount_reason", ""),
    )
    if notes_box is not None:
        story.append(Spacer(1, 10))
        story.append(notes_box)
        story.append(Spacer(1, 4))

    totals_rows = [
        ("Subtotal", _pdf_money(data['total_payable'])),
        ("Other", _pdf_money(0.0)),
        ("Total", _pdf_money(data['total_payable'])),
    ]
    story.append(_totals_block_owner(totals_rows))

    if is_paid:
        story.append(Spacer(1, 10))
        paid_so_far = float(data.get('amount_paid', 0) or 0)
        bal_due = float(data.get('balance_remaining', 0) or 0)
        pay_rows = [
            ("Amount Received", _pdf_money(paid_so_far)),
            ("Balance Due", _pdf_money(bal_due)),
        ]
        story.append(_totals_block_owner(pay_rows))

    story.append(Spacer(1, 20))
    story.append(_footer_line(S, "Page 1 of 1"))
    _safe_build(doc, story)
    if is_paid:
        try:
            consumer_bill_path = os.path.join(target_dir, f"{prefix}_Consumer_Copy.pdf")
            shutil.copy2(cust_path, consumer_bill_path)
        except Exception:
            pass
    return cust_path, own_path


def generate_receipt_pdfs(target_dir, data):
    ensure_dir(target_dir)
    tag = f"R{data['txn_id']:05d}"
    paid_date_val = str(data.get("payment_date") or data.get("paid_date") or "").strip()
    pdate_clean = paid_date_val.replace("-", "").replace("/", "")
    pamt_int = int(round(float(data.get("amount", 0.0))))
    amt_tag = f"_Rs{pamt_int}" if pamt_int > 0 else ""
    date_tag = f"_{pdate_clean}" if pdate_clean else ""
    cust = os.path.join(target_dir, f"Receipt_{tag}{date_tag}{amt_tag}_Customer.pdf")
    own = os.path.join(target_dir, f"Receipt_{tag}{date_tag}{amt_tag}_Owner.pdf")
    consumer = os.path.join(target_dir, f"Receipt_{tag}{date_tag}{amt_tag}_Consumer.pdf")
    S = _pdf_styles()

    bill_total = data.get("linked_bill_amount") or 0.0
    bill_remaining = data.get("linked_bill_remaining") or 0.0
    amount_paid = float(data.get('amount', 0.0))
    balance_after = float(data.get('balance_after', 0.0))

    # Customer Copy (and Consumer Copy)
    doc = SimpleDocTemplate(cust, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "PAYMENT RECEIPT", ""))
    story.append(Spacer(1, 14))
    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Property", data['property_name']),
        ("Receipt No.", f"RCPT-{data['txn_id']:05d}"),
        ("Paid Date", paid_date_val),
    ]))
    story.append(Spacer(1, 16))

    sub = None
    if data.get("linked_bill_month"):
        sub = f"Against bill: {data['linked_bill_month']}"
    cust_items = [
        ["Amount Received", _pdf_money(amount_paid), sub],
    ]
    story.append(_items_table_customer(cust_items))
    story.append(Spacer(1, 4))
    if data.get("linked_bill_month") or bill_total > 0:
        story.append(_totals_block_owner([
            ("Paid", _pdf_money(amount_paid)),
            ("Remaining", _pdf_money(bill_remaining)),
        ]))
    else:
        story.append(_totals_block_customer(_pdf_money(amount_paid)))
    story.append(Spacer(1, 12))

    method_line = data.get('payment_method') or "—"
    if data.get('reference_no'):
        method_line += f"  ·  Ref: {data['reference_no']}"
    story.append(Paragraph(f"<b>Method:</b> {method_line}", S["meta"]))
    story.append(Spacer(1, 8))
    if bill_remaining > 0.001:
        story.append(Paragraph(
            f"<font size=9 color='#B91C1C'>Remaining balance: {_pdf_money(bill_remaining)}</font>",
            S["meta"]))
    else:
        story.append(Paragraph(
            "<font size=9 color='#15803D'>Paid in full.</font>", S["meta"]))
    story.append(Spacer(1, 20))
    story.append(_footer_line(S))
    _safe_build(doc, story)

    # Owner Copy
    doc = SimpleDocTemplate(own, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "OWNER COPY / RECEIPT", ""))
    story.append(Spacer(1, 14))
    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Phone", data.get('owner_phone', '—') or '—'),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Property", data['property_name']),
        ("Receipt No.", f"RCPT-{data['txn_id']:05d}"),
        ("Paid Date", paid_date_val),
    ]))
    story.append(Spacer(1, 16))

    own_items = [
        ["Amount Received", "—", _pdf_money(amount_paid), method_line],
    ]
    if data.get("linked_bill_month"):
        own_items.append(["Against Bill", "—", data['linked_bill_month'], ""])
        own_items.append(["Bill Total", "—", _pdf_money(bill_total), ""])
        own_items.append(["This Payment", "—", _pdf_money(amount_paid), ""])
        own_items.append(["Remaining Due", "—", _pdf_money(bill_remaining), ""])
    story.append(_items_table_owner(own_items))
    story.append(Spacer(1, 4))
    story.append(_totals_block_owner([
        ("Paid", _pdf_money(amount_paid)),
        ("Remaining", _pdf_money(bill_remaining)),
        ("Balance After", _pdf_money(balance_after)),
    ]))
    story.append(Spacer(1, 12))
    if data.get("description"):
        story.append(Paragraph(f"<b>Notes:</b> {data['description']}", S["meta"]))
        story.append(Spacer(1, 8))
    story.append(_footer_line(S, "Page 1 of 1"))
    _safe_build(doc, story)
    try:
        shutil.copy2(cust, consumer)
    except Exception:
        pass
    return cust, own


def generate_final_settlement_pdfs(target_dir, data):
    ensure_dir(target_dir)
    cust = os.path.join(target_dir, "Final_Settlement_Customer.pdf")
    own = os.path.join(target_dir, "Final_Settlement_Owner.pdf")
    S = _pdf_styles()

    total_owed = data.get("prev_due", 0) + data.get("damages", 0)
    refund = max(0.0, data.get("deposit", 0) - total_owed)
    shortfall = max(0.0, total_owed - data.get("deposit", 0))

    doc = SimpleDocTemplate(cust, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "FINAL SETTLEMENT", ""))
    story.append(Spacer(1, 14))
    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Property", data['property_name']),
        ("Exit Date", data['bill_date']),
    ]))
    story.append(Spacer(1, 16))

    cust_items = []
    if data.get("prev_due", 0) > 0:
        cust_items.append(["Outstanding Dues", _pdf_money(data['prev_due']), None])
    if data.get("damages", 0) > 0:
        cust_items.append(["Damage Charges", _pdf_money(data['damages']), None])
    if data.get("deposit", 0) > 0:
        cust_items.append(["Security Deposit Held", f"-{_pdf_money(data['deposit'])}", None])
    story.append(_items_table_customer(cust_items))
    story.append(Spacer(1, 4))
    if refund > 0:
        story.append(_totals_block_customer(f"Refund: {_pdf_money(refund)}"))
    else:
        story.append(_totals_block_customer(f"Balance Owed: {_pdf_money(shortfall)}"))
    story.append(Spacer(1, 20))
    story.append(_footer_line(S))
    _safe_build(doc, story)

    doc = SimpleDocTemplate(own, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = []
    story.extend(_header_band(S, "OWNER COPY / SETTLEMENT", ""))
    story.append(Spacer(1, 14))
    story.append(_info_block([
        ("Owner", data['owner_name']),
        ("Phone", data.get('owner_phone', '—') or '—'),
        ("Tenant", f"{data['tenant_name']} ({data['tenant_uid']})"),
        ("Property", data['property_name']),
        ("Exit Date", data['bill_date']),
    ]))
    story.append(Spacer(1, 16))

    own_items = [
        ["Outstanding Dues", "—", _pdf_money(data.get('prev_due', 0)), "Carried forward"],
        ["Damage Charges", "—", _pdf_money(data.get('damages', 0)), data.get('remarks', '') or ""],
        ["Deposit Held", "—", _pdf_money(data.get('deposit', 0)), "Contract amount"],
    ]
    story.append(_items_table_owner(own_items))
    story.append(Spacer(1, 4))
    story.append(_totals_block_owner([
        ("Total Owed", _pdf_money(total_owed)),
        ("Refund Paid", _pdf_money(refund)),
        ("Shortfall", _pdf_money(shortfall)),
    ]))
    story.append(Spacer(1, 12))
    if data.get("remarks"):
        story.append(Paragraph(f"<b>Notes:</b> {data['remarks']}", S["meta"]))
        story.append(Spacer(1, 8))
    story.append(_footer_line(S, "Page 1 of 1"))
    _safe_build(doc, story)
    return cust, own


def generate_tenant_statement_pdf(target_path, info, txns):
    ensure_dir(os.path.dirname(target_path))
    S = _pdf_styles()
    doc = SimpleDocTemplate(target_path, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = _header_band(S, "ACCOUNT STATEMENT", "")
    story.append(Spacer(1, 14))
    story.append(_info_block([
        ("Tenant", f"{info['name']} ({info['tenant_uid']})"),
        ("Property", info['property_name']),
        ("Generated", datetime.now().strftime("%d %b %Y · %H:%M")),
    ]))
    story.append(Spacer(1, 16))

    rows = [[
        Paragraph("<b>Date</b>", S["table_hdr"]),
        Paragraph("<b>Type</b>", S["table_hdr"]),
        Paragraph("<b>Description</b>", S["table_hdr"]),
        Paragraph("<b>Debit</b>", S["table_hdr"]),
        Paragraph("<b>Credit</b>", S["table_hdr"]),
        Paragraph("<b>Balance</b>", S["table_hdr"]),
    ]]
    running = 0.0
    for t in txns:
        tt = t["txn_type"]; amt = t["amount"] or 0.0
        debit = credit = 0.0
        if tt in DEBIT_TYPES:
            debit = amt; running += amt
        elif tt in CREDIT_TYPES:
            credit = amt; running -= amt
        rows.append([
            Paragraph(t["txn_date"], S["table_val"]),
            Paragraph(tt.replace("_", " ").title(), S["table_val"]),
            Paragraph((t.get("description") or "")[:45], S["table_val"]),
            Paragraph(_pdf_money(debit) if debit else "—", S["table_val"]),
            Paragraph(_pdf_money(credit) if credit else "—", S["table_val"]),
            Paragraph(_pdf_money(running), S["table_val"]),
        ])
    tbl = RlTable(rows, colWidths=[70, 80, 160, 80, 80, 90])
    tbl.setStyle(RlTableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PDF_AL_DARK)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (3, 1), (5, -1), "RIGHT"),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(PDF_LINE)),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 6))
    story.append(_totals_block_customer(f"Closing: {_pdf_money(running)}"))
    story.append(Spacer(1, 20))
    story.append(_footer_line(S, "Page 1 of 1"))
    _safe_build(doc, story)


def generate_monthly_summary_pdf(target_path, user_id, month_str):
    ensure_dir(os.path.dirname(target_path))
    S = _pdf_styles()
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT t.name, p.name, b.total_payable, b.amount_paid, b.balance_remaining, b.settled_status
                   FROM tenant_bills b JOIN tenants t ON b.tenant_id = t.id
                   JOIN properties p ON t.property_id = p.id
                   WHERE b.user_id=? AND b.billing_month=? ORDER BY t.name""", (user_id, month_str))
    rows = cur.fetchall(); conn.close()
    total_billed = sum(r[2] or 0 for r in rows)
    total_paid = sum(r[3] or 0 for r in rows)
    total_out = sum(r[4] or 0 for r in rows)

    doc = SimpleDocTemplate(target_path, pagesize=letter,
                            rightMargin=44, leftMargin=44, topMargin=44, bottomMargin=44)
    story = _header_band(S, "MONTHLY SUMMARY", month_str.upper())
    story.append(Spacer(1, 14))

    data = [[
        Paragraph("<b>Tenant</b>", S["table_hdr"]),
        Paragraph("<b>Property</b>", S["table_hdr"]),
        Paragraph("<b>Billed</b>", S["table_hdr"]),
        Paragraph("<b>Paid</b>", S["table_hdr"]),
        Paragraph("<b>Outstanding</b>", S["table_hdr"]),
        Paragraph("<b>Status</b>", S["table_hdr"]),
    ]]
    for r in rows:
        total = float(r[2] or 0)
        paid = float(r[3] or 0)
        remaining = float(r[4] or 0)
        if remaining <= 0.001:
            status_str = "Paid"
        elif paid > 0.001:
            status_str = "Partial"
        else:
            status_str = "Unpaid"
        data.append([
            Paragraph(r[0], S["table_val"]),
            Paragraph(r[1], S["table_val"]),
            Paragraph(_pdf_money(total), S["table_val"]),
            Paragraph(_pdf_money(paid), S["table_val"]),
            Paragraph(_pdf_money(remaining), S["table_val"]),
            Paragraph(status_str, S["table_val"]),
        ])
    tbl = RlTable(data, colWidths=[110, 110, 90, 90, 100, 60])
    tbl.setStyle(RlTableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PDF_AL_DARK)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 1), (-2, -1), "RIGHT"),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor(PDF_LINE)),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(PDF_LINE)),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 6))
    story.append(_totals_block_owner([
        ("Total Billed", _pdf_money(total_billed)),
        ("Total Collected", _pdf_money(total_paid)),
        ("Total Outstanding", _pdf_money(total_out)),
    ]))
    story.append(Spacer(1, 20))
    story.append(_footer_line(S, "Page 1 of 1"))
    _safe_build(doc, story)


# ==============================================================================


# ==============================================================================
# § 6  GLASSMORHPIC COMPONENT SUITE & SHARED WIDGETS
# ==============================================================================
class GlassCard(QFrame):
    """Refined glassmorphic surface with layered translucency, specular edge
    glint, and clean hardware-accelerated rendering."""
    def __init__(self, parent=None, radius=RADIUS_CARD, level=1):
        super().__init__(parent)
        self._radius = radius
        self._level = level
        self.setObjectName("GlassCard")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(28)
        shadow.setColor(QColor(0, 0, 0, 140))
        shadow.setOffset(0, 8)
        self.setGraphicsEffect(shadow)

        self.setStyleSheet(f"""
            QFrame#GlassCard {{
                background-color: {T.c('card')};
                border: 1px solid {T.c('border')};
                border-top: 1px solid {T.c('border_specular')};
                border-radius: {self._radius}px;
            }}
            QFrame#GlassCard:hover {{
                background-color: {T.c('card_hover')};
                border: 1px solid {T.c('border_hover')};
                border-top: 1px solid rgba(255, 255, 255, 0.40);
            }}
            QLabel {{
                background: transparent;
                border: none;
                color: {T.c('text_primary')};
                font-family: {FONT_STACK};
            }}
        """)


class GlassModal(QFrame):
    """Elevated translucent modal frame with specular rim."""
    def __init__(self, parent=None, radius=RADIUS_MODAL):
        super().__init__(parent)
        self._radius = radius
        self.setObjectName("GlassModal")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(36)
        shadow.setColor(QColor(0, 0, 0, 180))
        shadow.setOffset(0, 12)
        self.setGraphicsEffect(shadow)

        self.setStyleSheet(f"""
            QFrame#GlassModal {{
                background-color: {T.c('dialog_bg')};
                border: 1px solid {T.c('dialog_border')};
                border-top: 1px solid rgba(255, 255, 255, 0.35);
                border-radius: {self._radius}px;
            }}
            QLabel {{
                background: transparent;
                border: none;
                color: {T.c('text_primary')};
                font-family: {FONT_STACK};
            }}
        """)


# High-fidelity backward-compatible aliases
class SolidCard(GlassCard):
    pass

class SolidModal(GlassModal):
    pass

class FrostedGlassFrame(GlassModal):
    pass


class GlassDialog(QDialog):
    """Standard frameless glassmorphic dialog with smooth dragging support."""
    def __init__(self, parent=None, width=500, height=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Dialog)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setFixedWidth(width)
        if height:
            self.setFixedHeight(height)
        self.drag_position = None

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            wh = self.windowHandle()
            if wh and hasattr(wh, "startSystemMove") and wh.startSystemMove():
                e.accept()
                return
            self.drag_position = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            e.accept()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.MouseButton.LeftButton and self.drag_position:
            self.move(e.globalPosition().toPoint() - self.drag_position)
            e.accept()

    def mouseReleaseEvent(self, e):
        self.drag_position = None
        super().mouseReleaseEvent(e)


class GlassBadge(QLabel):
    """Polished status pill with translucent background and accent border."""
    def __init__(self, text="", status="neutral", parent=None):
        super().__init__(parent)
        self.setText(f"  {text}  ")
        self.setFixedHeight(22)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.set_status(status)

    def set_status(self, status):
        styles = {
            "success": (ACCENT_GREEN, "rgba(16, 185, 129, 0.16)", "rgba(16, 185, 129, 0.40)"),
            "danger":  (ACCENT_RED,   "rgba(239, 68, 68, 0.16)",  "rgba(239, 68, 68, 0.40)"),
            "warning": (ACCENT_AMBER, "rgba(245, 158, 11, 0.16)", "rgba(245, 158, 11, 0.40)"),
            "info":    (ACCENT_SKY,   "rgba(14, 165, 233, 0.16)", "rgba(14, 165, 233, 0.40)"),
            "neutral": ("#94A3B8",    "rgba(255, 255, 255, 0.06)","rgba(255, 255, 255, 0.12)"),
        }
        color, bg, border = styles.get(status, styles["neutral"])
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg};
                color: {color};
                border: 1px solid {border};
                border-radius: 9999px;
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 0.5px;
                padding: 0 8px;
                font-family: {FONT_STACK};
            }}
        """)


class ToggleCheckBox(QAbstractButton):
    def __init__(self, label="", parent=None, width_pill=44, height_pill=24):
        super().__init__(parent)
        self.setCheckable(True)
        self._label = label
        self._h = height_pill
        self._w = width_pill
        fm = self.fontMetrics()
        total_w = self._w + (12 + fm.horizontalAdvance(label) if label else 0)
        self.setFixedSize(total_w, self._h + 4)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        x = 2
        y = (self.height() - self._h) // 2
        track = QRectF(x, y, self._w, self._h)
        on = self.isChecked()

        p.setPen(QPen(QColor(255, 255, 255, 60) if on else QColor(255, 255, 255, 30), 1))
        p.setBrush(QColor("#FFFFFF" if on else "rgba(255, 255, 255, 0.16)"))
        p.drawRoundedRect(track, self._h / 2, self._h / 2)

        knob_r = self._h - 6
        knob_x = x + self._w - knob_r - 3 if on else x + 3
        p.setBrush(QColor("#000000" if on else "#FFFFFF"))
        p.setPen(QPen(QColor(0, 0, 0, 40), 1))
        p.drawEllipse(QRectF(knob_x, y + 3, knob_r, knob_r))

        if self._label:
            f = QApplication.font()
            f.setFamily("SF Pro Display")
            f.setPixelSize(12)
            f.setWeight(QFont.Weight.Medium)
            p.setFont(f)
            p.setPen(QColor("#F8FAFC"))
            p.drawText(QRectF(self._w + 8, 0, self.width() - self._w - 8, self.height()),
                       Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, self._label)
        p.end()

    def sizeHint(self):
        return QSize(self.width(), self.height())


class PremiumToggle(ToggleCheckBox):
    def __init__(self, label="", parent=None):
        super().__init__(label, parent)



# ==============================================================================
# macOS Interactive Animation & UI Engine
# ==============================================================================

class AnimatedNumberLabel(QLabel):
    """macOS-style smooth odometer number interpolation (60fps cubic ease-out)."""
    def __init__(self, text="—", prefix="₹", is_currency=True, parent=None):
        super().__init__(text, parent)
        self._target_value = 0.0
        self._current_value = 0.0
        self._start_value = 0.0
        self._prefix = prefix
        self._is_currency = is_currency
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._step)
        self._steps = 22
        self._step_idx = 0

    def animate_to_value(self, target):
        try:
            target = float(target)
        except (ValueError, TypeError):
            self.setText(str(target))
            return
        if abs(self._current_value - target) < 0.001 and self._current_value != 0:
            self._render_val(target)
            return
        self._start_value = self._current_value
        self._target_value = target
        self._step_idx = 0
        self._timer.start(16)

    def _step(self):
        self._step_idx += 1
        t = self._step_idx / self._steps
        # Cubic ease-out: 1 - (1 - t)^3
        ease = 1.0 - (1.0 - t) ** 3
        val = self._start_value + (self._target_value - self._start_value) * ease
        self._current_value = val
        self._render_val(val)
        if self._step_idx >= self._steps:
            self._timer.stop()
            self._current_value = self._target_value
            self._render_val(self._target_value)

    def _render_val(self, val):
        if self._is_currency:
            self.setText(f"{self._prefix}{val:,.2f}")
        else:
            self.setText(str(int(round(val))))


class MacOSToast(QFrame):
    """macOS Dynamic Island floating glass notification banner with spring motion."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MacOSToast")
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setFixedWidth(380)
        self.setFixedHeight(50)
        self.setStyleSheet("""
            QFrame#MacOSToast {
                background-color: rgba(22, 22, 26, 0.94);
                border: 1px solid rgba(255, 255, 255, 0.22);
                border-top: 1px solid rgba(255, 255, 255, 0.45);
                border-radius: 25px;
            }
            QLabel {
                background: transparent;
                font-family: SF Pro Display, Inter, system-ui;
            }
        """)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setColor(QColor(0, 0, 0, 160))
        shadow.setOffset(0, 8)
        self.setGraphicsEffect(shadow)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 8, 20, 8)
        lay.setSpacing(12)

        self.icon_lbl = QLabel()
        self.icon_lbl.setFixedSize(22, 22)
        lay.addWidget(self.icon_lbl)

        self.msg_lbl = QLabel("")
        self.msg_lbl.setStyleSheet("color: #FFFFFF; font-size: 13px; font-weight: 600;")
        lay.addWidget(self.msg_lbl, 1)

        self._anim = None
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self._slide_out)
        self.hide()

    def show_message(self, message, icon_name="check", color="#10B981", duration_ms=2800):
        self.msg_lbl.setText(message)
        self.icon_lbl.setPixmap(SVGIconManager.pixmap(icon_name, color=color, size=20))

        if self.parent():
            w = self.parent().width()
            target_x = (w - self.width()) // 2
            self.move(target_x, -60)
            self.show()
            self.raise_()

            self._anim = QPropertyAnimation(self, b"pos")
            self._anim.setDuration(340)
            self._anim.setStartValue(QPoint(target_x, -60))
            self._anim.setEndValue(QPoint(target_x, 22))
            self._anim.setEasingCurve(QEasingCurve.Type.OutBack)
            self._anim.start()

            self._hide_timer.start(duration_ms)

    def _slide_out(self):
        target_x = self.x()
        self._anim = QPropertyAnimation(self, b"pos")
        self._anim.setDuration(260)
        self._anim.setStartValue(self.pos())
        self._anim.setEndValue(QPoint(target_x, -70))
        self._anim.setEasingCurve(QEasingCurve.Type.InCubic)
        self._anim.finished.connect(self.hide)
        self._anim.start()

# ==============================================================================
# QSS Stylesheet Builders
# ==============================================================================

_ICON_PNG_CACHE = {
    "chevron_down": "iVBORw0KGgoAAAANSUhEUgAAAAwAAAAMCAYAAABWdVznAAAAJklEQVR4nGNgGEZg2bIN/0mSAwlik8AljlUSr2J0RUQpJsnkwQMAMFovsVRamN0AAAAASUVORK5CYII=",
    "calendar": "iVBORw0KGgoAAAANSUhEUgAAAA4AAAAOCAYAAAAfSC3RAAAALklEQVR4nGNgoAZYtmzDfxAmVpx8jTAJYjGKRlK8Qx2NA+dUQuyhbCNZoUoOAAB4Gr7B0BDLjQAAAABJRU5ErkJggg==",
    "spin_up": "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAGCAYAAAD+Bd/7AAAAIElEQVR4nGNgIAUsW7bhPwjjlcSqCF0QhY/LWLzWIQMA9eIpuyZIfcEAAAAASUVORK5CYII=",
    "spin_down": "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAGCAYAAAD+Bd/7AAAAIElEQVR4nGNgIASWLdvwH4TxiqMrwqoJJojLRLzW4QUARDEpu7mtsXgAAAAASUVORK5CYII=",
}

def _get_qss_icon_path(name, color='#A6A6B0', size=14):
    try:
        base = os.path.join(_config_dir(), "icons")
        os.makedirs(base, exist_ok=True)
        p = os.path.join(base, f"{name}_{size}.png")
        if not os.path.exists(p):
            b64_data = _ICON_PNG_CACHE.get(name)
            if b64_data:
                import base64
                with open(p, "wb") as f:
                    f.write(base64.b64decode(b64_data))
        return p.replace("\\", "/")
    except Exception:
        return ""

def input_qss():
    return f"""
        QLineEdit, QTextEdit, QComboBox, QDateEdit, QSpinBox {{
            background-color: {T.c('input_bg')};
            border: 1px solid {T.c('input_border')};
            border-radius: {RADIUS_INPUT}px;
            padding: 10px 16px;
            color: {T.c('input_text')};
            font-size: {FS_INPUT};
            font-family: {FONT_STACK};
            font-weight: 500;
            selection-background-color: rgba(255, 255, 255, 0.25);
            selection-color: #FFFFFF;
            outline: none;
        }}
        QLineEdit:hover, QTextEdit:hover, QComboBox:hover, QDateEdit:hover, QSpinBox:hover {{
            border: 1px solid {T.c('input_border_hover')};
            background-color: {T.c('input_bg_hover')};
        }}
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus {{
            border: 1.5px solid {T.c('input_border_focus')};
            background-color: {T.c('input_bg_focus')};
            outline: none;
        }}
        QLineEdit:disabled, QSpinBox:disabled {{
            background-color: {T.c('input_disabled_bg')};
            color: {T.c('input_disabled_tx')};
            border: 1px solid {T.c('input_border')};
        }}
        QComboBox {{
            combobox-popup: 0;
            padding: 10px 32px 10px 14px;
        }}
        QDateEdit {{
            padding: 10px 32px 10px 14px;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            border: none;
            margin-right: 8px;
            width: 22px;
        }}
        QComboBox::down-arrow {{
            image: url({_get_qss_icon_path('chevron_down', color='#A6A6B0', size=12)});
            width: 12px;
            height: 12px;
        }}
        QComboBox QFrame {{
            background-color: #141416;
            border: none;
        }}
        QComboBox QAbstractItemView {{
            background-color: #141416;
            color: #FFFFFF;
            border: 1px solid rgba(255, 255, 255, 0.20);
            border-radius: 8px;
            padding: 4px;
            selection-background-color: rgba(255, 255, 255, 0.16);
            selection-color: #FFFFFF;
            outline: none;
            font-family: {FONT_STACK};
        }}
        QComboBox QAbstractItemView::item {{
            min-height: 28px;
            padding: 4px 10px;
            border-radius: 6px;
            color: #D1D5DB;
        }}
        QComboBox QAbstractItemView::item:hover, QComboBox QAbstractItemView::item:selected {{
            background-color: rgba(255, 255, 255, 0.18);
            color: #FFFFFF;
        }}
        QDateEdit::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 24px;
            border: none;
            margin-right: 8px;
        }}
        QDateEdit::down-arrow {{
            image: url({_get_qss_icon_path('calendar', color='#A6A6B0', size=14)});
            width: 14px;
            height: 14px;
        }}
        QSpinBox, QDoubleSpinBox {{
            padding: 8px 30px 8px 14px;
        }}
        QSpinBox QLineEdit, QDoubleSpinBox QLineEdit {{
            background: transparent;
            border: none;
            padding-left: 12px;
            color: #FFFFFF;
            font-size: {FS_INPUT};
            font-family: {FONT_STACK};
            font-weight: 600;
        }}
        QSpinBox::up-button, QSpinBox::down-button {{
            subcontrol-origin: border;
            width: 20px;
            border: none;
            background: rgba(255, 255, 255, 0.05);
            margin-right: 4px;
        }}
        QSpinBox::up-button {{
            subcontrol-position: top right;
            margin-top: 4px;
            border-top-right-radius: 8px;
        }}
        QSpinBox::down-button {{
            subcontrol-position: bottom right;
            margin-bottom: 4px;
            border-bottom-right-radius: 8px;
        }}
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
            background: rgba(255, 255, 255, 0.18);
        }}
        QSpinBox::up-arrow {{
            image: url({_get_qss_icon_path('spin_up', size=8)});
            width: 8px;
            height: 6px;
        }}
        QSpinBox::down-arrow {{
            image: url({_get_qss_icon_path('spin_down', size=8)});
            width: 8px;
            height: 6px;
        }}
        QCalendarWidget {{
            background-color: #141416;
            border: 1px solid rgba(255, 255, 255, 0.22);
            border-radius: 14px;
            font-family: {FONT_STACK};
        }}
        QCalendarWidget QWidget#qt_calendar_navigationbar {{
            background-color: #1A1A1E;
            border-bottom: 1px solid rgba(255, 255, 255, 0.12);
            border-top-left-radius: 12px;
            border-top-right-radius: 12px;
            min-height: 32px;
            max-height: 34px;
            padding: 2px 6px;
        }}
        QCalendarWidget QToolButton {{
            color: #FFFFFF;
            background-color: transparent;
            font-weight: 600;
            font-size: 12px;
            border-radius: 6px;
            padding: 3px 8px;
            margin: 1px;
            font-family: {FONT_STACK};
        }}
        QCalendarWidget QToolButton:hover {{
            background-color: rgba(255, 255, 255, 0.14);
            color: #FFFFFF;
        }}
        QCalendarWidget QToolButton::menu-indicator {{
            image: none;
            width: 0px;
        }}
        QCalendarWidget QSpinBox {{
            background-color: transparent;
            color: #FFFFFF;
            font-weight: 600;
            font-size: 12px;
            border: none;
            padding: 0 4px;
        }}
        QCalendarWidget QSpinBox::up-button, QCalendarWidget QSpinBox::down-button {{
            width: 0px;
            border: none;
        }}
        QCalendarWidget QMenu {{
            background-color: #1C1C20;
            color: #FFFFFF;
            border: 1px solid rgba(255, 255, 255, 0.20);
            border-radius: 8px;
            padding: 4px;
            font-family: {FONT_STACK};
        }}
        QCalendarWidget QMenu::item {{
            padding: 4px 12px;
            border-radius: 4px;
        }}
        QCalendarWidget QMenu::item:selected {{
            background-color: rgba(255, 255, 255, 0.18);
        }}
        QCalendarWidget QHeaderView {{
            background-color: #1A1A1E;
            border: none;
        }}
        QCalendarWidget QHeaderView::section {{
            background-color: #1A1A1E;
            color: #A6A6B0;
            border: none;
            font-size: 11px;
            font-weight: 600;
            padding: 2px 0px;
        }}
        QCalendarWidget QTableView {{
            background-color: #141416;
            border: none;
            outline: none;
            selection-background-color: #FFFFFF;
            selection-color: #000000;
        }}
        QCalendarWidget QAbstractItemView:enabled {{
            color: #FFFFFF;
            background-color: #141416;
            selection-background-color: #FFFFFF;
            selection-color: #000000;
            font-size: 11px;
            font-family: {FONT_STACK};
            outline: none;
            border: none;
            padding: 0px;
            margin: 0px;
        }}
        QCalendarWidget QAbstractItemView:disabled {{
            color: #55555C;
        }}
    """

INPUT_STYLE = input_qss()


def btn_qss(accent_hex, solid=True, height=40):
    is_destructive = accent_hex in (ACCENT_RED, "#EF4444", "#DC2626")
    h = accent_hex.lstrip("#")
    if len(h) == 6:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    else:
        r, g, b = 255, 255, 255

    if solid:
        bg_col = "#EF4444" if is_destructive else "#FFFFFF"
        tx_col = "#FFFFFF" if is_destructive else "#000000"
        hover_col = "#DC2626" if is_destructive else "#E5E5E5"
        press_col = "#B91C1C" if is_destructive else "#D4D4D4"
        return f"""
            QPushButton {{
                background-color: {bg_col};
                color: {tx_col};
                border: 1px solid rgba(255, 255, 255, 0.20);
                border-radius: {RADIUS_BTN}px;
                padding: 0 20px;
                font-weight: 700;
                font-size: {FS_INPUT};
                letter-spacing: 0.3px;
                font-family: {FONT_STACK};
                outline: none;
            }}
            QPushButton:hover {{
                background-color: {hover_col};
                border: 1px solid rgba(255, 255, 255, 0.40);
            }}
            QPushButton:pressed {{
                background-color: {press_col};
            }}
            QPushButton:focus {{ outline: none; }}
            QPushButton:disabled {{
                background-color: rgba(255, 255, 255, 0.20);
                color: rgba(255, 255, 255, 0.35);
                border: none;
            }}
        """
    # Secondary outline button: transparent with light outline
    return f"""
        QPushButton {{
            background-color: transparent;
            color: #FFFFFF;
            border: 1px solid rgba(255, 255, 255, 0.65);
            border-radius: {RADIUS_BTN}px;
            padding: 0 18px;
            font-weight: 600;
            font-size: {FS_INPUT};
            font-family: {FONT_STACK};
            outline: none;
        }}
        QPushButton:hover {{
            background-color: rgba(255, 255, 255, 0.12);
            border: 1px solid #FFFFFF;
            color: #FFFFFF;
        }}
        QPushButton:pressed {{
            background-color: rgba(255, 255, 255, 0.20);
        }}
        QPushButton:focus {{ outline: none; }}
        QPushButton:disabled {{
            color: rgba(255, 255, 255, 0.25);
            border: 1px solid rgba(255, 255, 255, 0.12);
            background-color: transparent;
        }}
    """


def primary_button_qss(accent_hex=None, height=40):
    if accent_hex is None:
        accent_hex = "#FFFFFF"
    return btn_qss(accent_hex, solid=True, height=height)


def pill_button_qss(accent_hex, height=34):
    return btn_qss(accent_hex, solid=False, height=height)


def round_button_qss(accent_hex, size=28):
    is_destructive = accent_hex in (ACCENT_RED, "#EF4444", "#DC2626")
    bg = "rgba(239, 68, 68, 0.20)" if is_destructive else "rgba(255, 255, 255, 0.10)"
    col = "#EF4444" if is_destructive else "#FFFFFF"
    border = "rgba(239, 68, 68, 0.40)" if is_destructive else "rgba(255, 255, 255, 0.25)"
    return f"""
        QPushButton {{
            background-color: {bg};
            color: {col};
            border: 1px solid {border};
            border-radius: {size // 2}px;
            font-weight: 700;
            font-size: {max(11, size // 3 + 1)}px;
            padding: 0;
            outline: none;
        }}
        QPushButton:hover {{
            background-color: {"rgba(239, 68, 68, 0.35)" if is_destructive else "rgba(255, 255, 255, 0.22)"};
            border: 1px solid {"#EF4444" if is_destructive else "rgba(255, 255, 255, 0.50)"};
            color: {col};
        }}
        QPushButton:pressed {{
            background-color: {"rgba(239, 68, 68, 0.45)" if is_destructive else "rgba(255, 255, 255, 0.32)"};
        }}
        QPushButton:focus {{ outline: none; }}
    """


def card_qss():
    return f"""
        QFrame#GlassCard, QFrame#MetricCard {{
            background-color: {T.c('card')};
            border: 1px solid {T.c('border')};
            border-top: 1px solid {T.c('border_specular')};
            border-radius: {RADIUS_CARD}px;
        }}
        QFrame#GlassCard:hover, QFrame#MetricCard:hover {{
            background-color: {T.c('card_hover')};
            border: 1px solid {T.c('border_hover')};
            border-top: 1px solid rgba(255, 255, 255, 0.40);
        }}
        QFrame#GlassCard QLabel, QFrame#MetricCard QLabel {{
            font-family: {FONT_STACK};
            color: {T.c('text_primary')};
            background: transparent;
        }}
    """


CARD_QSS_STATIC = card_qss()


def billing_card_qss():
    return f"""
        QFrame#BillingCard {{
            background-color: {T.c('card')};
            border: 1px solid {T.c('border')};
            border-top: 1px solid {T.c('border_specular')};
            border-radius: {RADIUS_CARD}px;
        }}
        QFrame#BillingCard QLabel {{
            color: {T.c('text_primary')};
            font-family: {FONT_STACK};
            background: transparent;
        }}
    """ + input_qss()


LIGHT_CARD_STYLE = billing_card_qss()


def table_qss():
    return f"""
        QTableWidget {{
            background-color: {T.c('table_bg')};
            border: 1px solid {T.c('table_border')};
            border-radius: 16px;
            color: {T.c('text_primary')};
            gridline-color: {T.c('table_grid')};
            font-size: 12px;
            font-family: {FONT_STACK};
            outline: none;
            alternate-background-color: rgba(255, 255, 255, 0.02);
            selection-background-color: #2D3142;
            selection-color: {T.c('text_primary')};
        }}
        QHeaderView::section {{
            background-color: {T.c('table_hdr_bg')};
            color: {T.c('text_primary')};
            border: none;
            border-bottom: 1px solid {T.c('table_border')};
            padding: 8px 6px;
            font-weight: 700;
            font-size: 11px;
            font-family: {FONT_STACK};
            letter-spacing: 0.4px;
        }}
        QTableWidget::item {{
            padding: 6px 8px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.04);
        }}
        QTableWidget::item:hover {{
            background-color: rgba(255, 255, 255, 0.08);
        }}
        QTableWidget::item:selected {{
            background-color: #2D3142;
        }}
        QTableWidget::item:focus {{
            background-color: #2D3142;
        }}
        QScrollBar:vertical {{
            background: transparent;
            width: 6px;
            margin: 4px 2px 4px 0;
        }}
        QScrollBar::handle:vertical {{
            background: {T.c('scroll_handle')};
            border-radius: 3px;
            min-height: 20px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: rgba(255, 255, 255, 0.35);
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
        }}
    """


CHECKBOX_QSS = f"""
    QCheckBox {{
        color: {T.c('text_primary')};
        font-size: 13px;
        font-family: {FONT_STACK};
        font-weight: 500;
        background: transparent;
        spacing: 8px;
        outline: none;
        padding: 3px 0;
    }}
    QCheckBox::indicator {{
        width: 18px; height: 18px;
        background-color: {T.c('input_bg')};
        border: 1.5px solid {T.c('input_border')};
        border-radius: 5px;
    }}
    QCheckBox::indicator:hover {{
        border: 1.5px solid {T.c('input_border_focus')};
        background-color: {T.c('input_bg_hover')};
    }}
    QCheckBox::indicator:checked {{
        background-color: {T.c('accent')};
        border: 1.5px solid {T.c('accent')};
    }}
    QCheckBox::indicator:disabled {{
        background-color: {T.c('input_disabled_bg')};
        border: 1.5px solid {T.c('input_border')};
    }}
"""


def section_label(text):
    w = QWidget()
    w.setStyleSheet("background: transparent;")
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 0, 0, 0)
    h.setSpacing(8)
    bar = QFrame()
    bar.setFixedSize(3, 13)
    bar.setStyleSheet("background-color: #FF6A00; border-radius: 2px;")
    lbl = QLabel(text.upper())
    lbl.setStyleSheet(
        f"color: #A0A0A0; font-size: {FS_LABEL}; "
        f"font-weight: 700; letter-spacing: 1.4px; background: transparent; "
        f"font-family: {FONT_STACK};"
    )
    h.addWidget(bar)
    h.addWidget(lbl)
    h.addStretch(1)
    return w


# ==============================================================================
# UtilityMetricsBarChart (Glassmorphic Financial Bar Chart)
# ==============================================================================
class UtilityMetricsBarChart(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(160)
        self.months = ["—"] * 6
        self.billed = [0.0] * 6
        self.collected = [0.0] * 6

    def set_data(self, months, billed, collected):
        self.months = months
        self.billed = billed
        self.collected = collected
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect()).adjusted(1, 1, -1, -1)

        # Frosted glass container frame
        p.setBrush(QColor(40, 40, 40, 150))
        border_grad = QLinearGradient(0, 0, 0, rect.height())
        border_grad.setColorAt(0.0, QColor(255, 255, 255, 45))
        border_grad.setColorAt(1.0, QColor(255, 255, 255, 12))
        p.setPen(QPen(QBrush(border_grad), 1))
        p.drawRoundedRect(rect, RADIUS_CARD, RADIUS_CARD)

        total_b = sum(self.billed)
        total_c = sum(self.collected)
        pct = (total_c / total_b * 100.0) if total_b > 0 else 0.0

        # Donut Chart Geometry (Left Side of Card)
        ring_size = min(rect.height() - 36, 136.0)
        ring_rect = QRectF(rect.left() + 28, rect.top() + (rect.height() - ring_size) / 2, ring_size, ring_size)
        stroke_w = 13.0

        # Background track for ring
        track_pen = QPen(QColor(255, 255, 255, 20), stroke_w)
        track_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(track_pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawArc(ring_rect.adjusted(stroke_w / 2, stroke_w / 2, -stroke_w / 2, -stroke_w / 2), 0, 360 * 16)

        # Multi-stop Gradient Stroke: Purple -> Magenta -> Red -> Orange -> Yellow-Lime
        ring_grad = QLinearGradient(ring_rect.topLeft(), ring_rect.bottomRight())
        ring_grad.setColorAt(0.00, QColor("#7B2FF7"))   # Purple
        ring_grad.setColorAt(0.25, QColor("#C2185B"))   # Magenta
        ring_grad.setColorAt(0.50, QColor("#FF3B1F"))   # Red-Orange
        ring_grad.setColorAt(0.75, QColor("#FF6A00"))   # Orange
        ring_grad.setColorAt(1.00, QColor("#D4E600"))   # Yellow-Lime

        span_angle = max(int((pct / 100.0) * 360 * 16), 16 if pct > 0 else 0)
        if span_angle > 360 * 16:
            span_angle = 360 * 16

        if span_angle > 0:
            arc_pen = QPen(QBrush(ring_grad), stroke_w)
            arc_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            p.setPen(arc_pen)
            p.drawArc(ring_rect.adjusted(stroke_w / 2, stroke_w / 2, -stroke_w / 2, -stroke_w / 2), 90 * 16, -span_angle)

        # Dark inner circle
        inner_margin = stroke_w + 3
        inner_rect = ring_rect.adjusted(inner_margin, inner_margin, -inner_margin, -inner_margin)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(18, 18, 18, 220))
        p.drawEllipse(inner_rect)

        # Center KPI Number & Caption inside donut dynamically
        f_num = QFont(QApplication.font())
        f_num.setPixelSize(18)
        f_num.setBold(True)
        p.setFont(f_num)
        p.setPen(QColor("#FFFFFF"))
        num_str = f"{int(round(pct))}%" if total_b > 0 else "0%"

        f_cap = QFont(QApplication.font())
        f_cap.setPixelSize(9)
        f_cap.setWeight(QFont.Weight.DemiBold)

        fm_num = QFontMetrics(f_num)
        fm_cap = QFontMetrics(f_cap)
        h_num = fm_num.height()
        h_cap = fm_cap.height()
        total_text_h = h_num + 2 + h_cap
        start_y = inner_rect.center().y() - total_text_h / 2.0

        p.drawText(QRectF(inner_rect.left(), start_y, inner_rect.width(), h_num),
                   Qt.AlignmentFlag.AlignCenter, num_str)

        p.setFont(f_cap)
        p.setPen(QColor("#A0A0A0"))
        p.drawText(QRectF(inner_rect.left(), start_y + h_num + 2, inner_rect.width(), h_cap),
                   Qt.AlignmentFlag.AlignCenter, "COLLECTED")

        # Monthly Comparison Bars (Right Side of Card)
        pad_x_left = ring_rect.right() + 44
        pad_x_right = rect.right() - 28
        cw = max(pad_x_right - pad_x_left, 120.0)
        pad_y = 26
        ch = rect.height() - pad_y * 2 - 12
        max_val = max(max(self.billed), max(self.collected), 100.0) * 1.15

        # Subtle reference lines
        for i in range(3):
            y = pad_y + (ch / 2) * i
            p.setPen(QPen(QColor(255, 255, 255, 14), 1, Qt.PenStyle.DashLine))
            p.drawLine(int(pad_x_left), int(y), int(pad_x_right), int(y))

        n = len(self.months)
        seg = cw / max(n, 1)
        bw = 10

        f2 = QFont(QApplication.font())
        f2.setPixelSize(10)
        f2.setWeight(QFont.Weight.DemiBold)
        p.setFont(f2)

        for i in range(n):
            cx = pad_x_left + seg * i + seg / 2
            bh = (self.billed[i] / max_val) * ch
            ch_ = (self.collected[i] / max_val) * ch
            r1 = QRectF(cx - bw - 2, pad_y + ch - bh, bw, max(bh, 2.0))
            r2 = QRectF(cx + 2, pad_y + ch - ch_, bw, max(ch_, 2.0))

            # Billed bar: Purple to Magenta vertical gradient
            g1 = QLinearGradient(r1.left(), r1.top(), r1.left(), r1.bottom())
            g1.setColorAt(0.0, QColor("#C2185B"))
            g1.setColorAt(1.0, QColor("#7B2FF7"))
            p.setBrush(g1)
            p.setPen(Qt.PenStyle.NoPen)
            p.drawRoundedRect(r1, 3, 3)

            # Collected bar: Red-Orange to Yellow-Lime vertical gradient
            g2 = QLinearGradient(r2.left(), r2.top(), r2.left(), r2.bottom())
            g2.setColorAt(0.0, QColor("#D4E600"))
            g2.setColorAt(1.0, QColor("#FF3B1F"))
            p.setBrush(g2)
            p.drawRoundedRect(r2, 3, 3)

            # Month label
            p.setPen(QColor("#A0A0A0"))
            p.drawText(QRectF(cx - seg / 2, pad_y + ch + 6, seg, 18),
                       Qt.AlignmentFlag.AlignCenter, self.months[i])


# ==============================================================================
# FileGenerator (Background Asynchronous Generator)
# ==============================================================================
class FileGenerator(QThread):
    finished_ok = pyqtSignal(list)
    failed = pyqtSignal(str)

    def __init__(self, tasks, parent=None):
        super().__init__(parent)
        self.tasks = tasks

    def run(self):
        try:
            results = []
            for t in self.tasks:
                results.append(t())
            self.finished_ok.emit(results)
        except Exception as e:
            tb = traceback.format_exc()
            log.error(f"FileGenerator failed: {e}\n{tb}")
            self.failed.emit(f"{e}\n\n{tb}")


# ==============================================================================
# § 7  AUTHENTICATION & ONBOARDING
# ==============================================================================
class PasswordStrengthChecker:
    @staticmethod
    def evaluate(pw):
        rules = {
            "length": len(pw) >= 8,
            "upper": bool(re.search(r"[A-Z]", pw)),
            "lower": bool(re.search(r"[a-z]", pw)),
            "digit": bool(re.search(r"\d", pw)),
            "special": bool(re.search(r"[!@#$%^&*()_+\-=\[\]{};':\"\\|,.<>/?`~]", pw)),
        }
        score = sum(rules.values())
        if score <= 2: label, color = "Weak", "#EF4444"
        elif score == 3: label, color = "Fair", "#F59E0B"
        elif score == 4: label, color = "Good", "#3B82F6"
        else: label, color = "Strong", "#10B981"
        return rules, score, label, color


class PasswordRuleWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 6, 2, 0)
        lay.setSpacing(6)
        self.bar = QProgressBar()
        self.bar.setRange(0, 5)
        self.bar.setValue(0)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(4)
        self.bar.setStyleSheet("""
            QProgressBar { background-color: rgba(255,255,255,0.08); border: none; border-radius: 2px; }
            QProgressBar::chunk { background-color: #EF4444; border-radius: 2px; }
        """)
        lay.addWidget(self.bar)
        self.label = QLabel("Password strength: —")
        self.label.setStyleSheet(f"font-size: 11px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        lay.addWidget(self.label)
        self.rules_lbl = QLabel()
        self.rules_lbl.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; background: transparent; line-height: 150%; font-family: {FONT_STACK};")
        self.rules_lbl.setWordWrap(True)
        lay.addWidget(self.rules_lbl)

    def update_for(self, pw):
        rules, score, label, color = PasswordStrengthChecker.evaluate(pw)
        self.bar.setValue(score)
        self.bar.setStyleSheet(f"""
            QProgressBar {{ background-color: rgba(255,255,255,0.08); border: none; border-radius: 2px; }}
            QProgressBar::chunk {{ background-color: {color}; border-radius: 2px; }}
        """)
        self.label.setText(f"Password strength: <b style='color:{color};'>{label}</b>")
        def mark(ok): return "<span style='color:#10B981;'>✓</span>" if ok else "<span style='color:#EF4444;'>✗</span>"
        self.rules_lbl.setText(
            f"{mark(rules['length'])}  Minimum 8 characters<br>"
            f"{mark(rules['upper'])}  At least 1 uppercase (A–Z)<br>"
            f"{mark(rules['lower'])}  At least 1 lowercase (a–z)<br>"
            f"{mark(rules['digit'])}  At least 1 digit (0–9)<br>"
            f"{mark(rules['special'])}  At least 1 special character (!@#$%^&*)"
        )

    def is_strong(self): return self.bar.value() == 5


def _storage_row_qss_cyan():
    return f"""
        QPushButton {{ background: transparent; border: none;
            color: {T.c('text_secondary')}; font-size: 11px;
            font-weight: 600; text-decoration: underline;
            outline: none; padding: 2px 4px;
            font-family: {FONT_STACK}; }}
        QPushButton:hover {{ color: {ACCENT_ORANGE}; }}
    """


def _storage_row_qss_red():
    return f"""
        QPushButton {{ background: transparent; border: none;
            color: {T.c('text_faint')}; font-size: 11px;
            font-weight: 600; text-decoration: underline;
            outline: none; padding: 2px 4px;
            font-family: {FONT_STACK}; }}
        QPushButton:hover {{ color: {ACCENT_RED}; }}
    """


def _storage_actions(controller, parent_widget):
    data_folder_btn = QPushButton("Data Folder")
    data_folder_btn.setIcon(SVGIconManager.icon("folder", color="#FFFFFF", size=14))
    data_folder_btn.setIconSize(QSize(14, 14))
    data_folder_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    data_folder_btn.setStyleSheet(_storage_row_qss_cyan())
    data_folder_btn.clicked.connect(lambda: _open_data_folder_picker(controller, parent_widget))

    reset_btn = QPushButton("Reset App Data")
    reset_btn.setIcon(SVGIconManager.icon("settings", color="#888888", size=14))
    reset_btn.setIconSize(QSize(14, 14))
    reset_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    reset_btn.setStyleSheet(_storage_row_qss_red())
    reset_btn.clicked.connect(lambda: _reset_app_data(controller, parent_widget))

    dot = QLabel("·")
    dot.setStyleSheet(f"color: {T.c('text_muted')}; background: transparent; font-family: {FONT_STACK}; font-size: 11px;")

    row = QHBoxLayout()
    row.setSpacing(8)
    row.addStretch(1)
    row.addWidget(data_folder_btn)
    row.addWidget(dot)
    row.addWidget(reset_btn)
    row.addStretch(1)
    return row


def _open_data_folder_picker(controller, parent_widget):
    dlg = DataFolderPickerDialog(parent=parent_widget, is_change=True, current_path=get_apex_root())
    if not dlg.exec() or not dlg.chosen:
        return
    new_path = dlg.chosen
    old_path = get_apex_root()
    if os.path.abspath(new_path) == os.path.abspath(old_path):
        QMessageBox.information(parent_widget, "Same Folder", "That's already your data folder.")
        return
    reply = QMessageBox.question(
        parent_widget, "Copy Existing Data?",
        f"Copy existing database and backups from:\n\n"
        f"    {old_path}\n\n"
        f"to:\n\n"
        f"    {new_path}\n\n"
        f"Yes  →  copy files (recommended)\n"
        f"No   →  start fresh in the new folder\n\n"
        f"Note: this only takes effect after the app restarts.",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
    )
    if reply == QMessageBox.StandardButton.Cancel:
        return
    try:
        os.makedirs(new_path, exist_ok=True)
        if reply == QMessageBox.StandardButton.Yes:
            for name in ["users.db", "settings.ini", "apex.log"]:
                src = os.path.join(old_path, name)
                dst = os.path.join(new_path, name)
                if os.path.exists(src):
                    shutil.copy2(src, dst)
            src_bk = os.path.join(old_path, "backups")
            dst_bk = os.path.join(new_path, "backups")
            if os.path.isdir(src_bk):
                if os.path.isdir(dst_bk):
                    shutil.rmtree(dst_bk, ignore_errors=True)
                shutil.copytree(src_bk, dst_bk)
        save_data_folder(new_path)
        QMessageBox.information(
            parent_widget, "Saved",
            "Data folder updated.\n\n"
            "Please close and reopen the app for the change to take effect."
        )
    except Exception as e:
        log.error(f"Data folder change failed: {e}")
        QMessageBox.warning(parent_widget, "Error", f"Could not change folder:\n\n{e}")


def _reset_app_data(controller, parent_widget):
    if QMessageBox.warning(
        parent_widget, "Reset App Data",
        "This will permanently delete:\n\n"
        "  • The login database (all accounts and passwords)\n"
        "  • All tenants, properties, bills, and payments\n"
        "  • The pointer to your data folder\n\n"
        "Generated PDFs and Excel files in the data folder are NOT deleted.\n\n"
        "You will be asked to choose a folder the next time you launch.\n\n"
        "This cannot be undone. Continue?",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No
    ) != QMessageBox.StandardButton.Yes:
        return
    text, ok = QInputDialog.getText(parent_widget, "Confirm Reset", "Type RESET (all caps) to confirm:")
    if not ok or text.strip() != "RESET":
        QMessageBox.information(parent_widget, "Cancelled", "Reset cancelled.")
        return
    data_folder = get_apex_root()
    try:
        for name in ["users.db", "settings.ini"]:
            p = os.path.join(data_folder, name)
            if os.path.exists(p):
                try: os.remove(p)
                except Exception as e: log.warning(f"Reset: could not remove {name}: {e}")
        bk = os.path.join(data_folder, "backups")
        if os.path.isdir(bk):
            shutil.rmtree(bk, ignore_errors=True)
        pointer = _config_file()
        if os.path.exists(pointer):
            os.remove(pointer)
    except Exception as e:
        QMessageBox.warning(parent_widget, "Reset Failed", f"Could not reset:\n\n{e}")
        return
    QMessageBox.information(
        parent_widget, "Reset Complete",
        "App data has been reset.\n\n"
        "Please close and reopen the application — you will be asked "
        "to choose a new data folder on next launch."
    )


class DataFolderPickerDialog(GlassDialog):
    def __init__(self, parent=None, is_change=False, current_path=None):
        super().__init__(parent, width=520)
        self.is_change = is_change
        self.current_path = current_path
        self.chosen = None
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(30, 26, 30, 26)
        lay.setSpacing(14)

        header_row = QHBoxLayout()
        icon_lbl = QLabel()
        icon_lbl.setPixmap(SVGIconManager.pixmap("folder", color="#FFFFFF", size=24))
        title_text = (
            "<span style='color: #FFFFFF;'>CHOOSE</span> <span style='color: #FF6A00;'>FOLDER</span>"
            if self.is_change else
            "<span style='color: #FFFFFF;'>DATA</span> <span style='color: #FF6A00;'>STORAGE</span>"
        )
        title = QLabel(title_text)
        title.setStyleSheet(f"font-size: {FS_HERO}; font-weight: 800; background: transparent; font-family: {FONT_STACK}; letter-spacing: 1px;")
        header_row.addWidget(icon_lbl)
        header_row.addSpacing(8)
        header_row.addWidget(title)
        header_row.addStretch(1)
        lay.addLayout(header_row)

        sub = QLabel(
            "Your database, backups, logs, and generated Excel/PDF files will be stored "
            "in this folder. You can change this later from the login screen."
        )
        sub.setStyleSheet(f"font-size: 12px; color: {T.c('text_dim')}; line-height: 155%; background: transparent; font-family: {FONT_STACK};")
        sub.setWordWrap(True)
        lay.addWidget(sub)
        lay.addSpacing(6)

        rec_lbl = QLabel("RECOMMENDED REPOSITORY")
        rec_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.8px; background: transparent; font-family: {FONT_STACK};")
        lay.addWidget(rec_lbl)

        rec_path = QLabel(_default_data_folder())
        rec_path.setStyleSheet(f"""
            background: rgba(30, 30, 30, 0.65);
            border: 1px solid rgba(255, 255, 255, 0.18);
            border-radius: {RADIUS_INPUT}px;
            padding: 11px 15px; color: #FFFFFF;
            font-size: 12px; font-weight: 600;
            font-family: {FONT_STACK};
        """)
        rec_path.setWordWrap(True)
        lay.addWidget(rec_path)

        rec_btn = QPushButton("USE RECOMMENDED")
        rec_btn.setIcon(SVGIconManager.icon("check", color="#000000", size=16))
        rec_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        rec_btn.setFixedHeight(BTN_PRIMARY)
        rec_btn.setStyleSheet(primary_button_qss("#FFFFFF", height=BTN_PRIMARY))
        rec_btn.clicked.connect(self._use_recommended)
        lay.addWidget(rec_btn)

        lay.addSpacing(6)
        or_lbl = QLabel("— OR CHOOSE CUSTOM LOCATION —")
        or_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        or_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_muted')}; background: transparent; font-family: {FONT_STACK}; letter-spacing: 0.8px;")
        lay.addWidget(or_lbl)

        pick_btn = QPushButton("CHOOSE A DIFFERENT FOLDER…")
        pick_btn.setIcon(SVGIconManager.icon("folder", color="#FFFFFF", size=16))
        pick_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        pick_btn.setFixedHeight(BTN_PRIMARY)
        pick_btn.setStyleSheet(pill_button_qss("#FFFFFF", height=BTN_PRIMARY))
        pick_btn.clicked.connect(self._browse)
        lay.addWidget(pick_btn)

        if self.is_change and self.current_path:
            lay.addSpacing(6)
            cur_lbl = QLabel(f"Current: {self.current_path}")
            cur_lbl.setStyleSheet(f"font-size: 11px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            cur_lbl.setWordWrap(True)
            lay.addWidget(cur_lbl)

        if self.is_change:
            lay.addSpacing(4)
            cancel_btn = QPushButton("CANCEL")
            cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            cancel_btn.setFixedHeight(BTN_SECONDARY)
            cancel_btn.setStyleSheet(pill_button_qss("#FFFFFF", height=BTN_SECONDARY))
            cancel_btn.clicked.connect(self.reject)
            lay.addWidget(cancel_btn)

        base.addWidget(env)

    def _use_recommended(self):
        self.chosen = _default_data_folder()
        self.accept()

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "Choose data folder", os.path.expanduser("~"))
        if d:
            self.chosen = os.path.join(d, "ApexLedgerPro") if not d.endswith("ApexLedgerPro") else d
            self.accept()


class LoginCard(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 40, 36, 36)
        layout.setSpacing(12)

        self.opacity_effect = QGraphicsOpacityEffect(self)
        self.opacity_effect.setOpacity(1.0)
        self.setGraphicsEffect(self.opacity_effect)

        title = QLabel("<span style='color: #FFFFFF;'>WELCOME</span> <span style='color: #FF6A00;'>BACK</span>")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"font-size: {FS_HERO}; font-weight: 800; background: transparent; font-family: {FONT_STACK}; letter-spacing: 1.5px;")
        subtitle = QLabel(f"Sign in to {APP_NAME} v{APP_VERSION}")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")

        self.username_in = QLineEdit()
        self.username_in.setPlaceholderText("Username")
        self.username_in.setStyleSheet(INPUT_STYLE)
        self.username_in.setFixedHeight(INPUT_H)

        pw_row = QHBoxLayout()
        pw_row.setSpacing(8)
        self.password_in = QLineEdit()
        self.password_in.setPlaceholderText("Password")
        self.password_in.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_in.setStyleSheet(INPUT_STYLE)
        self.password_in.setFixedHeight(INPUT_H)
        self.password_in.returnPressed.connect(self.attempt_login)

        self.toggle_pw_btn = QPushButton()
        self.toggle_pw_btn.setIcon(SVGIconManager.icon("eye", color="#FFFFFF", size=16))
        self.toggle_pw_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_pw_btn.setFixedSize(INPUT_H, INPUT_H)
        self.toggle_pw_btn.setCheckable(True)
        self.toggle_pw_btn.setStyleSheet(round_button_qss("#FFFFFF", size=INPUT_H))
        self.toggle_pw_btn.toggled.connect(self._toggle_login_pw)
        pw_row.addWidget(self.password_in, 1)
        pw_row.addWidget(self.toggle_pw_btn)

        self.remember_toggle = PremiumToggle("Remember Me")

        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet(f"font-size: 11px; color: {ACCENT_RED}; background: transparent; font-family: {FONT_STACK};")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setMinimumHeight(18)

        self.login_btn = QPushButton("SIGN IN")
        self.login_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.login_btn.setFixedHeight(BTN_HERO)
        self.login_btn.setStyleSheet(primary_button_qss("#FFFFFF", height=BTN_HERO))
        self.login_btn.clicked.connect(self.attempt_login)

        switch_row = QHBoxLayout()
        switch_row.setSpacing(6)
        no_acc = QLabel("Don't have an account?")
        no_acc.setStyleSheet(f"color: {T.c('text_secondary')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        switch_btn = QPushButton("Create one")
        switch_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        switch_btn.setStyleSheet(f"""
            QPushButton {{ background: transparent; border: none; color: #FFFFFF;
                font-size: 12px; font-weight: 600; text-decoration: underline; outline: none;
                font-family: {FONT_STACK}; }}
            QPushButton:hover {{ color: {ACCENT_ORANGE}; }}
        """)
        switch_btn.clicked.connect(lambda: self.controller.fade_to_page(1))
        switch_row.addStretch(1)
        switch_row.addWidget(no_acc)
        switch_row.addWidget(switch_btn)
        switch_row.addStretch(1)

        storage_row = _storage_actions(self.controller, self)

        layout.addStretch(1)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(12)
        layout.addWidget(self.username_in)
        layout.addSpacing(2)
        layout.addLayout(pw_row)
        layout.addSpacing(8)
        layout.addWidget(self.remember_toggle)
        layout.addSpacing(4)
        layout.addWidget(self.status_lbl)
        layout.addSpacing(4)
        layout.addWidget(self.login_btn)
        layout.addSpacing(12)
        layout.addLayout(switch_row)
        layout.addSpacing(6)
        layout.addLayout(storage_row)
        layout.addStretch(1)

        settings = load_settings()
        if settings.get("remember_username"):
            self.username_in.setText(settings["remember_username"])
            self.remember_toggle.setChecked(True)

    def attempt_login(self):
        username = self.username_in.text().strip()
        password = self.password_in.text().strip()
        if not username or not password:
            self.status_lbl.setText("Enter both username and password.")
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT id, fullname FROM users WHERE username=? AND password_hash=?",
                    (username, hash_password(password)))
        row = cur.fetchone()
        conn.close()
        if not row:
            self.status_lbl.setText("Invalid credentials. Try again.")
            log.warning(f"Failed login for {username}")
            return
        settings = load_settings()
        if self.remember_toggle.isChecked(): settings["remember_username"] = username
        else: settings.pop("remember_username", None)
        save_settings(settings)
        self.status_lbl.setText("")
        self.controller.current_session_id = row[0]
        self.controller.dashboard_page.populate_verified_session(row[1])
        self.controller.dashboard_page.refresh_tab_states()
        self.controller.fade_to_page(2)
        log.info(f"Login ok: uid={row[0]}")

    def prefill_username(self, username):
        self.username_in.setText(username)
        self.password_in.clear()
        self.status_lbl.setText("")
        self.password_in.setFocus()

    def _toggle_login_pw(self, on):
        self.password_in.setEchoMode(QLineEdit.EchoMode.Normal if on else QLineEdit.EchoMode.Password)
        self.toggle_pw_btn.setIcon(SVGIconManager.icon("eye_off" if on else "eye", color="#FFFFFF", size=16))


class CredentialReminderDialog(GlassDialog):
    def __init__(self, username, password, parent=None):
        super().__init__(parent, width=440)
        self.username = username
        self.password = password
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(30, 26, 30, 26)
        lay.setSpacing(12)

        head_row = QHBoxLayout()
        ok_icon = QLabel()
        ok_icon.setPixmap(SVGIconManager.pixmap("check", color="#FFFFFF", size=22))
        title = QLabel("<span style='color: #FFFFFF;'>ACCOUNT</span> <span style='color: #FF6A00;'>CREATED</span>")
        title.setStyleSheet(f"font-size: {FS_HERO}; font-weight: 800; background: transparent; font-family: {FONT_STACK}; letter-spacing: 1px;")
        head_row.addWidget(ok_icon)
        head_row.addSpacing(8)
        head_row.addWidget(title)
        head_row.addStretch(1)
        lay.addLayout(head_row)

        sub = QLabel("Please save these credentials.\nThis is the last time we will show your password.")
        sub.setStyleSheet(f"font-size: 12px; color: {T.c('text_dim')}; background: transparent; line-height: 155%; font-family: {FONT_STACK};")
        sub.setWordWrap(True)
        lay.addWidget(sub)
        lay.addSpacing(4)

        u_lbl = QLabel("USERNAME")
        u_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.8px; background: transparent; font-family: {FONT_STACK};")
        lay.addWidget(u_lbl)
        u_row = QHBoxLayout()
        u_row.setSpacing(8)
        self.u_val = QLineEdit(self.username)
        self.u_val.setReadOnly(True)
        self.u_val.setStyleSheet(INPUT_STYLE)
        self.u_val.setFixedHeight(INPUT_H)
        u_row.addWidget(self.u_val, 1)

        u_copy = QPushButton()
        u_copy.setIcon(SVGIconManager.icon("copy", color="#FFFFFF", size=16))
        u_copy.setCursor(Qt.CursorShape.PointingHandCursor)
        u_copy.setFixedSize(BTN_SECONDARY, BTN_SECONDARY)
        u_copy.setStyleSheet(round_button_qss("#FFFFFF", size=BTN_SECONDARY))
        u_copy.clicked.connect(lambda: self._copy(self.username, "Username copied"))
        u_row.addWidget(u_copy)
        lay.addLayout(u_row)

        p_lbl = QLabel("PASSWORD")
        p_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.8px; background: transparent; font-family: {FONT_STACK};")
        lay.addWidget(p_lbl)
        p_row = QHBoxLayout()
        p_row.setSpacing(8)
        self.p_val = QLineEdit(self.password)
        self.p_val.setReadOnly(True)
        self.p_val.setEchoMode(QLineEdit.EchoMode.Password)
        self.p_val.setStyleSheet(INPUT_STYLE)
        self.p_val.setFixedHeight(INPUT_H)
        p_row.addWidget(self.p_val, 1)

        self.show_btn = QPushButton()
        self.show_btn.setIcon(SVGIconManager.icon("eye", color="#FFFFFF", size=16))
        self.show_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.show_btn.setFixedSize(BTN_SECONDARY, BTN_SECONDARY)
        self.show_btn.setCheckable(True)
        self.show_btn.setStyleSheet(round_button_qss("#FFFFFF", size=BTN_SECONDARY))
        self.show_btn.toggled.connect(self._toggle_pw)
        p_row.addWidget(self.show_btn)

        p_copy = QPushButton()
        p_copy.setIcon(SVGIconManager.icon("copy", color="#FFFFFF", size=16))
        p_copy.setCursor(Qt.CursorShape.PointingHandCursor)
        p_copy.setFixedSize(BTN_SECONDARY, BTN_SECONDARY)
        p_copy.setStyleSheet(round_button_qss("#FFFFFF", size=BTN_SECONDARY))
        p_copy.clicked.connect(lambda: self._copy(self.password, "Password copied"))
        p_row.addWidget(p_copy)
        lay.addLayout(p_row)
        lay.addSpacing(6)

        copy_both = QPushButton("COPY BOTH")
        copy_both.setIcon(SVGIconManager.icon("copy", color="#000000", size=16))
        copy_both.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_both.setFixedHeight(BTN_PRIMARY)
        copy_both.setStyleSheet(primary_button_qss("#FFFFFF", height=BTN_PRIMARY))
        copy_both.clicked.connect(self._copy_both)
        lay.addWidget(copy_both)
        lay.addSpacing(4)

        self.confirm_cb = ToggleCheckBox("I have securely saved these credentials")
        self.confirm_cb.toggled.connect(self._toggle_proceed)
        lay.addWidget(self.confirm_cb)
        lay.addSpacing(2)

        self.proceed_btn = QPushButton("PROCEED TO LOGIN")
        self.proceed_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.proceed_btn.setFixedHeight(BTN_HERO)
        self.proceed_btn.setEnabled(False)
        self.proceed_btn.setStyleSheet(primary_button_qss("#FFFFFF", height=BTN_HERO))
        self.proceed_btn.clicked.connect(self.accept)
        lay.addWidget(self.proceed_btn)

        base.addWidget(env)

    def _toggle_pw(self, on):
        self.p_val.setEchoMode(QLineEdit.EchoMode.Normal if on else QLineEdit.EchoMode.Password)
        self.show_btn.setIcon(SVGIconManager.icon("eye_off" if on else "eye", color="#FFFFFF", size=16))

    def _copy(self, text, msg):
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copied", msg)

    def _copy_both(self):
        QApplication.clipboard().setText(f"Username: {self.username}\nPassword: {self.password}")
        QMessageBox.information(self, "Copied", "Username and password copied.")

    def _toggle_proceed(self, checked):
        self.proceed_btn.setEnabled(checked)


class SignUpCard(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(36, 36, 36, 36)
        layout.setSpacing(10)

        title = QLabel("<span style='color: #FFFFFF;'>CREATE</span> <span style='color: #FF6A00;'>ACCOUNT</span>")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 800; background: transparent; font-family: {FONT_STACK}; letter-spacing: 1.5px;")
        subtitle = QLabel("Register to begin managing your portfolio")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")

        self.fullname_in = QLineEdit()
        self.fullname_in.setPlaceholderText("Full Name")
        self.fullname_in.setStyleSheet(INPUT_STYLE)
        self.fullname_in.setFixedHeight(INPUT_H)

        self.username_in = QLineEdit()
        self.username_in.setPlaceholderText("Username")
        self.username_in.setStyleSheet(INPUT_STYLE)
        self.username_in.setFixedHeight(INPUT_H)

        self.password_in = QLineEdit()
        self.password_in.setPlaceholderText("Password")
        self.password_in.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_in.setStyleSheet(INPUT_STYLE)
        self.password_in.setFixedHeight(INPUT_H)

        self.rules_widget = PasswordRuleWidget()
        self.password_in.textChanged.connect(self.rules_widget.update_for)

        self.confirm_in = QLineEdit()
        self.confirm_in.setPlaceholderText("Confirm Password")
        self.confirm_in.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_in.setStyleSheet(INPUT_STYLE)
        self.confirm_in.setFixedHeight(INPUT_H)

        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet(f"font-size: 11px; color: {ACCENT_RED}; background: transparent; font-family: {FONT_STACK};")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setMinimumHeight(18)

        signup_btn = QPushButton("SIGN UP")
        signup_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        signup_btn.setFixedHeight(BTN_HERO)
        signup_btn.setStyleSheet(primary_button_qss("#FFFFFF", height=BTN_HERO))
        signup_btn.clicked.connect(self.attempt_signup)

        switch_row = QHBoxLayout()
        switch_row.setSpacing(6)
        has_acc = QLabel("Already registered?")
        has_acc.setStyleSheet(f"color: {T.c('text_secondary')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        switch_btn = QPushButton("Log in")
        switch_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        switch_btn.setStyleSheet(f"""
            QPushButton {{ background: transparent; border: none; color: #FFFFFF;
                font-size: 12px; font-weight: 600; text-decoration: underline; outline: none;
                font-family: {FONT_STACK}; }}
            QPushButton:hover {{ color: {ACCENT_ORANGE}; }}
        """)
        switch_btn.clicked.connect(lambda: self.controller.fade_to_page(0))
        switch_row.addStretch(1)
        switch_row.addWidget(has_acc)
        switch_row.addWidget(switch_btn)
        switch_row.addStretch(1)

        storage_row = _storage_actions(self.controller, self)

        layout.addStretch(1)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(6)
        layout.addWidget(self.fullname_in)
        layout.addWidget(self.username_in)
        layout.addWidget(self.password_in)
        layout.addWidget(self.rules_widget)
        layout.addWidget(self.confirm_in)
        layout.addWidget(self.status_lbl)
        layout.addSpacing(2)
        layout.addWidget(signup_btn)
        layout.addSpacing(4)
        layout.addLayout(switch_row)
        layout.addSpacing(4)
        layout.addLayout(storage_row)
        layout.addStretch(1)

    def attempt_signup(self):
        fullname = self.fullname_in.text().strip()
        username = self.username_in.text().strip()
        password = self.password_in.text().strip()
        confirm = self.confirm_in.text().strip()
        if not (fullname and username and password and confirm):
            self.status_lbl.setText("All fields are required.")
            return
        if password != confirm:
            self.status_lbl.setText("Passwords do not match.")
            return
        if not self.rules_widget.is_strong():
            self.status_lbl.setText("Password must satisfy all 5 security rules.")
            return

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        try:
            cur.execute("INSERT INTO users (fullname, username, password_hash) VALUES (?, ?, ?)",
                        (fullname, username, hash_password(password)))
            conn.commit()
        except sqlite3.IntegrityError:
            conn.close()
            self.status_lbl.setText("Username already exists.")
            return
        finally:
            conn.close()

        self.status_lbl.setText("")
        dlg = CredentialReminderDialog(username, password, parent=self)
        if dlg.exec():
            for w in [self.fullname_in, self.username_in, self.password_in, self.confirm_in]:
                w.clear()
            self.rules_widget.update_for("")
            self.controller.login_card.prefill_username(username)
            self.controller.fade_to_page(0)


# ==============================================================================
# § 8  OWNER PROFILE
# ==============================================================================
class OwnerDetailsSubPage(QWidget):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.chosen_dir = None
        self.init_ui()

    def init_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(18, 18, 18, 18)
        self.main_layout.setSpacing(14)

        # 1. VIEW MODE
        self.view_widget = QWidget()
        self.view_widget.setStyleSheet("background: transparent;")
        vl = QVBoxLayout(self.view_widget)
        vl.setContentsMargins(0,0,0,0)
        vl.setSpacing(14)

        hdr_row = QHBoxLayout()
        owner_ic = QLabel()
        owner_ic.setPixmap(SVGIconManager.pixmap("owner", color=ACCENT_INDIGO, size=20))
        hdr = QLabel("Owner Profile")
        hdr.setStyleSheet(f"font-size: {FS_SUB}; color: {T.c('text_primary')}; font-weight: 700; background: transparent; font-family: {FONT_STACK};")
        hdr_row.addWidget(owner_ic)
        hdr_row.addSpacing(6)
        hdr_row.addWidget(hdr)
        hdr_row.addStretch(1)
        vl.addLayout(hdr_row)

        self.card_profile = GlassCard()
        pl = QVBoxLayout(self.card_profile)
        pl.setContentsMargins(22, 20, 22, 20)
        pl.setSpacing(12)

        self.lbl_name = QLabel("Name: —")
        self.lbl_phone = QLabel("Phone: —")
        self.lbl_email = QLabel("Email: —")
        self.lbl_address = QLabel("Address: —")
        for w in [self.lbl_name, self.lbl_phone, self.lbl_email, self.lbl_address]:
            w.setStyleSheet(f"color: {T.c('text_primary')}; font-size: 13px; font-family: {FONT_STACK}; background: transparent;")
            pl.addWidget(w)
        vl.addWidget(self.card_profile)

        self.card_dir = GlassCard()
        dl = QVBoxLayout(self.card_dir)
        dl.setContentsMargins(22, 20, 22, 20)
        dl.setSpacing(10)

        dir_head = QHBoxLayout()
        fold_ic = QLabel()
        fold_ic.setPixmap(SVGIconManager.pixmap("folder", color=ACCENT_AMBER, size=18))
        dir_title = QLabel("Output Directory")
        dir_title.setStyleSheet(f"font-weight: 700; color: {ACCENT_AMBER}; font-size: 13px; background: transparent; font-family: {FONT_STACK};")
        dir_head.addWidget(fold_ic)
        dir_head.addSpacing(6)
        dir_head.addWidget(dir_title)
        dir_head.addStretch(1)
        dl.addLayout(dir_head)

        dir_row = QHBoxLayout()
        dir_row.setSpacing(10)
        self.lbl_dir = QLabel("Not configured. Falls back to ~/Desktop.")
        self.lbl_dir.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        self.lbl_dir.setWordWrap(True)

        browse = QPushButton("BROWSE")
        browse.setIcon(SVGIconManager.icon("folder", color=ACCENT_AMBER, size=14))
        browse.setCursor(Qt.CursorShape.PointingHandCursor)
        browse.setFixedHeight(BTN_INLINE)
        browse.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_INLINE))
        browse.clicked.connect(self.browse_dir)

        dir_row.addWidget(self.lbl_dir, 1)
        dir_row.addWidget(browse)
        dl.addLayout(dir_row)
        vl.addWidget(self.card_dir)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        edit_btn = QPushButton("EDIT DETAILS")
        edit_btn.setIcon(SVGIconManager.icon("edit", color="#04150E", size=16))
        edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        edit_btn.setFixedHeight(BTN_PRIMARY)
        edit_btn.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        edit_btn.clicked.connect(self.switch_to_edit)
        action_row.addWidget(edit_btn)

        del_btn = QPushButton("DELETE PROFILE")
        del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=16))
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.setFixedHeight(BTN_PRIMARY)
        del_btn.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_PRIMARY))
        del_btn.clicked.connect(self.delete_profile)
        action_row.addWidget(del_btn)
        vl.addLayout(action_row)
        vl.addStretch(1)

        # 2. FORM / EDIT MODE
        self.form_widget = QWidget()
        self.form_widget.setStyleSheet("background: transparent;")
        fl = QVBoxLayout(self.form_widget)
        fl.setContentsMargins(0,0,0,0)
        fl.setSpacing(10)

        fhdr = QLabel("Owner Profile Registration")
        fhdr.setStyleSheet(f"font-size: {FS_SUB}; color: {T.c('text_primary')}; font-weight: 700; background: transparent; font-family: {FONT_STACK};")
        self.status_lbl = QLabel("Fill required database fields.")
        self.status_lbl.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        fl.addWidget(fhdr)
        fl.addWidget(self.status_lbl)

        self.name_in = QLineEdit()
        self.name_in.setPlaceholderText("Full Name *")
        self.name_in.setStyleSheet(INPUT_STYLE)
        self.name_in.setFixedHeight(INPUT_H)

        self.phone_in = QLineEdit()
        self.phone_in.setPlaceholderText("Phone (10 digits) *")
        self.phone_in.setStyleSheet(INPUT_STYLE)
        self.phone_in.setFixedHeight(INPUT_H)
        self.phone_in.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,10}$"), self))

        self.email_in = QLineEdit()
        self.email_in.setPlaceholderText("Email (Optional)")
        self.email_in.setStyleSheet(INPUT_STYLE)
        self.email_in.setFixedHeight(INPUT_H)

        self.address_in = QTextEdit()
        self.address_in.setPlaceholderText("Permanent Address *")
        self.address_in.setStyleSheet(INPUT_STYLE)
        self.address_in.setMaximumHeight(70)

        geo1 = QHBoxLayout()
        geo1.setSpacing(10)
        self.city_in = QLineEdit()
        self.city_in.setPlaceholderText("City *")
        self.city_in.setStyleSheet(INPUT_STYLE)
        self.city_in.setFixedHeight(INPUT_H)

        self.state_in = QLineEdit()
        self.state_in.setPlaceholderText("State *")
        self.state_in.setStyleSheet(INPUT_STYLE)
        self.state_in.setFixedHeight(INPUT_H)
        geo1.addWidget(self.city_in)
        geo1.addWidget(self.state_in)

        geo2 = QHBoxLayout()
        geo2.setSpacing(10)
        self.country_in = QLineEdit()
        self.country_in.setPlaceholderText("Country *")
        self.country_in.setStyleSheet(INPUT_STYLE)
        self.country_in.setFixedHeight(INPUT_H)

        self.pin_in = QLineEdit()
        self.pin_in.setPlaceholderText("PIN (6 digits) *")
        self.pin_in.setStyleSheet(INPUT_STYLE)
        self.pin_in.setFixedHeight(INPUT_H)
        self.pin_in.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,6}$"), self))
        geo2.addWidget(self.country_in)
        geo2.addWidget(self.pin_in)

        dir_row2 = QHBoxLayout()
        dir_row2.setSpacing(10)
        self.lbl_dir_form = QLabel("Output Path: Default (~/Desktop)")
        self.lbl_dir_form.setStyleSheet(f"""
            background-color: {T.c('input_bg')};
            border: 1px solid {T.c('input_border')};
            border-radius: {RADIUS_INPUT}px;
            padding: 9px 13px;
            color: {T.c('text_secondary')};
            font-size: {FS_INPUT};
            font-family: {FONT_STACK};
        """)
        browse2 = QPushButton("BROWSE...")
        browse2.setIcon(SVGIconManager.icon("folder", color=ACCENT_AMBER, size=14))
        browse2.setCursor(Qt.CursorShape.PointingHandCursor)
        browse2.setFixedSize(120, BTN_PRIMARY)
        browse2.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_PRIMARY))
        browse2.clicked.connect(self.browse_dir_form)
        dir_row2.addWidget(self.lbl_dir_form, 1)
        dir_row2.addWidget(browse2)

        save_btn = QPushButton("SAVE PROFILE")
        save_btn.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.setFixedHeight(BTN_PRIMARY)
        save_btn.setStyleSheet(primary_button_qss(ACCENT_GREEN, height=BTN_PRIMARY))
        save_btn.clicked.connect(self.save_profile)

        fl.addWidget(self.name_in)
        fl.addWidget(self.phone_in)
        fl.addWidget(self.email_in)
        fl.addWidget(self.address_in)
        fl.addLayout(geo1)
        fl.addLayout(geo2)
        fl.addLayout(dir_row2)
        fl.addSpacing(4)
        fl.addWidget(save_btn)
        fl.addStretch(1)

        self.main_layout.addWidget(self.view_widget)
        self.main_layout.addWidget(self.form_widget)
        self.view_widget.hide()
        self.form_widget.show()

    def fade_swap(self, fn):
        if callable(fn):
            fn()
        self.update()

    def load_existing_profile(self):
        uid = self.controller.current_session_id
        if uid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT owner_name, phone, email, address, city, state, country, pin, excel_pdf_dir FROM profiles WHERE user_id=?", (uid,))
        row = cur.fetchone()
        conn.close()
        if row:
            for w, v in zip([self.name_in, self.phone_in, self.email_in], row[:3]):
                w.setText(v or "")
            self.address_in.setPlainText(row[3] or "")
            for w, v in zip([self.city_in, self.state_in, self.country_in, self.pin_in], row[4:8]):
                w.setText(v or "")
            self.chosen_dir = row[8]
            full = f"{row[3]}, {row[4]}, {row[5]}, {row[6]} - {row[7]}"
            self.lbl_name.setText(f"<b>Name:</b> {row[0]}")
            self.lbl_phone.setText(f"<b>Phone:</b> {row[1]}")
            self.lbl_email.setText(f"<b>Email:</b> {row[2] or '— (not provided)'}")
            self.lbl_address.setText(f"<b>Address:</b> {full}")
            if row[8]:
                self.lbl_dir.setText(row[8])
                self.lbl_dir.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 12px; font-weight: 600; background: transparent; font-family: {FONT_STACK};")
                self.lbl_dir_form.setText(row[8])
            else:
                self.lbl_dir.setText("Not configured. Falls back to ~/Desktop.")
                self.lbl_dir.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
                self.lbl_dir_form.setText("Output Path: Default (~/Desktop)")
            self.form_widget.hide()
            self.view_widget.show()
        else:
            for w in [self.name_in, self.phone_in, self.email_in, self.city_in, self.state_in, self.country_in, self.pin_in]:
                w.clear()
            self.address_in.clear()
            self.chosen_dir = None
            self.lbl_dir.setText("Not configured. Falls back to ~/Desktop.")
            self.lbl_dir.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
            self.lbl_dir_form.setText("Output Path: Default (~/Desktop)")
            self.status_lbl.setText("Fill required database fields.")
            self.status_lbl.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            self.view_widget.hide()
            self.form_widget.show()
        self.update()

    def browse_dir(self):
        uid = self.controller.current_session_id
        if uid is None: return
        d = QFileDialog.getExistingDirectory(self, "Select Output Directory", os.path.expanduser("~"))
        if d:
            self.chosen_dir = d
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("UPDATE profiles SET excel_pdf_dir=? WHERE user_id=?", (d, uid))
            conn.commit()
            conn.close()
            self.lbl_dir.setText(d)
            self.lbl_dir.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 12px; font-weight: 600; background: transparent; font-family: {FONT_STACK};")
            self.lbl_dir_form.setText(d)

    def browse_dir_form(self):
        d = QFileDialog.getExistingDirectory(self, "Select Output Directory", os.path.expanduser("~"))
        if d:
            self.chosen_dir = d
            self.lbl_dir_form.setText(d)

    def switch_to_edit(self):
        self.status_lbl.setText("Modifying active profile...")
        self.status_lbl.setStyleSheet(f"font-size: {FS_HELPER}; color: {ACCENT_AMBER}; background: transparent; font-family: {FONT_STACK};")
        self.view_widget.hide()
        self.form_widget.show()
        self.update()

    def save_profile(self):
        uid = self.controller.current_session_id
        if uid is None: return
        name = self.name_in.text().strip()
        phone = self.phone_in.text().strip()
        email = self.email_in.text().strip()
        address = self.address_in.toPlainText().strip()
        city = self.city_in.text().strip()
        state = self.state_in.text().strip()
        country = self.country_in.text().strip()
        pin = self.pin_in.text().strip()
        if not (name and phone and address and city and state and country and pin):
            self.status_lbl.setText("Mandatory fields missing (*).")
            self.status_lbl.setStyleSheet(f"color: {ACCENT_RED}; background: transparent; font-family: {FONT_STACK};")
            return
        if len(phone) != 10:
            self.status_lbl.setText("Phone must be exactly 10 digits.")
            self.status_lbl.setStyleSheet(f"color: {ACCENT_RED}; background: transparent; font-family: {FONT_STACK};")
            return
        if len(pin) != 6:
            self.status_lbl.setText("PIN must be exactly 6 digits.")
            self.status_lbl.setStyleSheet(f"color: {ACCENT_RED}; background: transparent; font-family: {FONT_STACK};")
            return

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("INSERT OR REPLACE INTO profiles (user_id, owner_name, phone, email, address, city, state, country, pin, excel_pdf_dir) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (uid, name, phone, email, address, city, state, country, pin, self.chosen_dir))
        conn.commit()
        conn.close()

        self.controller.dashboard_page.populate_verified_session(name)
        self.controller.dashboard_page.refresh_tab_states()
        self.load_existing_profile()
        self.update()

    def delete_profile(self):
        uid = self.controller.current_session_id
        if uid is None: return
        if QMessageBox.question(self, "Delete Profile", "Are you sure?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("DELETE FROM profiles WHERE user_id=?", (uid,))
        conn.commit()
        conn.close()
        self.chosen_dir = None
        self.controller.dashboard_page.refresh_tab_states()
        self.load_existing_profile()
        self.update()


# ==============================================================================
# § 9  PROPERTIES
# ==============================================================================
class PropertyDialog(GlassDialog):
    def __init__(self, user_id, property_data=None, parent=None):
        super().__init__(parent, width=480)
        self.user_id = user_id
        self.property_data = property_data
        self.init_ui()
        if property_data:
            self.populate()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        head = QHBoxLayout()
        icon_lbl = QLabel()
        icon_lbl.setPixmap(SVGIconManager.pixmap("properties", color=ACCENT_INDIGO, size=18))
        t = QLabel("PROPERTY REGISTRATION")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.4px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(icon_lbl)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        lay.addWidget(self._lbl("Property Name *"))
        self.name_in = QLineEdit()
        self.name_in.setPlaceholderText("e.g. Skyline Skyvilla")
        self.name_in.setStyleSheet(INPUT_STYLE)
        self.name_in.setFixedHeight(INPUT_H)
        lay.addWidget(self.name_in)

        addr_row = QHBoxLayout()
        addr_row.addWidget(self._lbl("Property Address *"), 1)
        use_owner = QPushButton("USE OWNER'S ADDRESS")
        use_owner.setIcon(SVGIconManager.icon("owner", color=ACCENT_INDIGO, size=14))
        use_owner.setCursor(Qt.CursorShape.PointingHandCursor)
        use_owner.setFixedHeight(BTN_INLINE)
        use_owner.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_INLINE))
        use_owner.clicked.connect(self.fetch_owner_addr)
        addr_row.addWidget(use_owner)
        lay.addLayout(addr_row)

        self.address_in = QTextEdit()
        self.address_in.setPlaceholderText("Street, area, landmark...")
        self.address_in.setStyleSheet(INPUT_STYLE)
        self.address_in.setMaximumHeight(64)
        lay.addWidget(self.address_in)

        # Clean 2-column paired layout for Type & BHK
        grid1 = QHBoxLayout()
        grid1.setSpacing(10)
        v_type = QVBoxLayout()
        v_type.setSpacing(4)
        v_type.addWidget(self._lbl("Asset Classification *"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Apartment", "Land", "Others"])
        self.type_combo.setStyleSheet(INPUT_STYLE)
        self.type_combo.setFixedHeight(INPUT_H)
        self.type_combo.currentTextChanged.connect(self.on_type)
        v_type.addWidget(self.type_combo)
        grid1.addLayout(v_type, 1)

        self.bhk_row = QWidget()
        self.bhk_row.setStyleSheet("background: transparent;")
        v_bhk = QVBoxLayout(self.bhk_row)
        v_bhk.setContentsMargins(0, 0, 0, 0)
        v_bhk.setSpacing(4)
        v_bhk.addWidget(self._lbl("Configuration"))
        self.bhk_combo = QComboBox()
        self.bhk_combo.addItems(["1 RK", "1 BHK", "2 BHK", "3 BHK", "4+ BHK"])
        self.bhk_combo.setStyleSheet(INPUT_STYLE)
        self.bhk_combo.setFixedHeight(INPUT_H)
        v_bhk.addWidget(self.bhk_combo)
        grid1.addWidget(self.bhk_row, 1)
        lay.addLayout(grid1)

        self.land_row = QWidget()
        self.land_row.setStyleSheet("background: transparent;")
        lr = QHBoxLayout(self.land_row)
        lr.setContentsMargins(0, 0, 0, 0)
        lr.setSpacing(10)
        v_sqft = QVBoxLayout()
        v_sqft.setSpacing(4)
        v_sqft.addWidget(self._lbl("Total Area (sq.ft)"))
        self.sqft_in = QLineEdit()
        self.sqft_in.setPlaceholderText("Total Area (sq.ft)")
        self.sqft_in.setStyleSheet(INPUT_STYLE)
        self.sqft_in.setFixedHeight(INPUT_H)
        self.sqft_in.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        v_sqft.addWidget(self.sqft_in)

        v_rate = QVBoxLayout()
        v_rate.setSpacing(4)
        v_rate.addWidget(self._lbl("Rate/sq.ft (₹)"))
        self.rate_sqft_in = QLineEdit()
        self.rate_sqft_in.setPlaceholderText("Rate/sq.ft (₹)")
        self.rate_sqft_in.setStyleSheet(INPUT_STYLE)
        self.rate_sqft_in.setFixedHeight(INPUT_H)
        self.rate_sqft_in.setValidator(QDoubleValidator(0.0, 1e7, 2, self))
        v_rate.addWidget(self.rate_sqft_in)

        self.sqft_in.textChanged.connect(self.land_calc)
        self.rate_sqft_in.textChanged.connect(self.land_calc)
        lr.addLayout(v_sqft)
        lr.addLayout(v_rate)
        lay.addWidget(self.land_row)
        self.land_row.hide()

        # Clean 2-column paired layout for Rent & Water Meter
        grid2 = QHBoxLayout()
        grid2.setSpacing(10)
        v_rent = QVBoxLayout()
        v_rent.setSpacing(4)
        v_rent.addWidget(self._lbl("Base Rent (₹) *"))
        self.rent_in = QLineEdit()
        self.rent_in.setPlaceholderText("Monthly Rent")
        self.rent_in.setStyleSheet(INPUT_STYLE)
        self.rent_in.setFixedHeight(INPUT_H)
        self.rent_in.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        v_rent.addWidget(self.rent_in)
        grid2.addLayout(v_rent, 1)

        v_meter = QVBoxLayout()
        v_meter.setSpacing(4)
        v_meter.addWidget(self._lbl("Water Meter Type"))
        self.meter_type_combo = QComboBox()
        self.meter_type_combo.addItems(["Common / Shared Meter", "Individual Meter"])
        self.meter_type_combo.setStyleSheet(INPUT_STYLE)
        self.meter_type_combo.setFixedHeight(INPUT_H)
        v_meter.addWidget(self.meter_type_combo)
        grid2.addLayout(v_meter, 1)
        lay.addLayout(grid2)

        # Utility Outlets row (Owner family members collected directly in Water Record)
        v_util = QVBoxLayout()
        v_util.setSpacing(4)
        v_util.addWidget(self._lbl("Utility Outlets"))
        util = QHBoxLayout()
        util.setSpacing(12)
        self.water_chk = ToggleCheckBox("Water Available")
        self.elec_chk = ToggleCheckBox("Electricity Available")
        util.addWidget(self.water_chk)
        util.addWidget(self.elec_chk)
        util.addStretch(1)
        v_util.addLayout(util)
        lay.addLayout(v_util)

        lay.addSpacing(8)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("SAVE PROPERTY")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save.clicked.connect(self.save)

        btn_row.addWidget(cancel, 1)
        btn_row.addWidget(save, 1)
        lay.addLayout(btn_row)

        base.addWidget(env)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def _plain(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"color: {T.c('text_soft')}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        return l

    def fetch_owner_addr(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT address, city, state, country, pin FROM profiles WHERE user_id=?", (self.user_id,))
        row = cur.fetchone()
        conn.close()
        if row:
            parts = [p for p in [row[0], row[1], row[2], row[3]] if p]
            s = ", ".join(parts)
            if row[4]: s += f" - {row[4]}"
            self.address_in.setPlainText(s)
        else:
            QMessageBox.information(self, "No Profile", "No owner profile to fetch.")

    def on_type(self, t):
        if t == "Apartment":
            self.bhk_row.show()
            self.land_row.hide()
        elif t == "Land":
            self.bhk_row.hide()
            self.land_row.show()
        else:
            self.bhk_row.hide()
            self.land_row.hide()

    def land_calc(self):
        try:
            a = float(self.sqft_in.text() or 0)
            r = float(self.rate_sqft_in.text() or 0)
            self.rent_in.setText(f"{a * r:.2f}")
        except ValueError:
            pass

    def populate(self):
        d = self.property_data
        self.name_in.setText(str(d[2]))
        self.address_in.setPlainText(str(d[3]))
        self.type_combo.setCurrentText(str(d[4]))
        self.on_type(str(d[4]))
        if d[5]: self.bhk_combo.setCurrentText(str(d[5]))
        if d[6]: self.sqft_in.setText(str(d[6]))
        if d[7]: self.rate_sqft_in.setText(str(d[7]))
        self.rent_in.setText(str(d[8]))
        self.water_chk.setChecked(bool(d[9]))
        mt = d[10] if len(d) > 10 else "Common"
        self.meter_type_combo.setCurrentText("Individual Meter" if mt == "Individual" else "Common / Shared Meter")
        self.elec_chk.setChecked(bool(d[11] if len(d) > 11 else 0))

    def save(self):
        name = self.name_in.text().strip()
        addr = self.address_in.toPlainText().strip()
        ptype = self.type_combo.currentText()
        rent_raw = self.rent_in.text().strip()
        if not (name and addr and rent_raw):
            QMessageBox.warning(self, "Validation", "Name, address and rent are required.")
            return
        try:
            rent = float(rent_raw)
        except ValueError:
            QMessageBox.warning(self, "Validation", "Rent must be numeric.")
            return
        bhk = self.bhk_combo.currentText() if ptype == "Apartment" else None
        sqft = float(self.sqft_in.text()) if (ptype == "Land" and self.sqft_in.text()) else None
        rate = float(self.rate_sqft_in.text()) if (ptype == "Land" and self.rate_sqft_in.text()) else None
        water = 1 if self.water_chk.isChecked() else 0
        meter = "Individual" if "Individual" in self.meter_type_combo.currentText() else "Common"
        elec = 1 if self.elec_chk.isChecked() else 0
        owner_members = int(self.property_data[13]) if (self.property_data and len(self.property_data) > 13 and self.property_data[13]) else 1

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        if self.property_data:
            cur.execute("""UPDATE properties SET name=?, address=?, property_type=?, bhk=?, area_sqft=?, rate_per_sqft=?,
                rent=?, water_available=?, water_meter_type=?, electricity_available=?, owner_members=? WHERE id=?""",
                (name, addr, ptype, bhk, sqft, rate, rent, water, meter, elec, owner_members, self.property_data[0]))
        else:
            cur.execute("""INSERT INTO properties (user_id, name, address, property_type, bhk, area_sqft, rate_per_sqft,
                rent, water_available, water_meter_type, electricity_available, owner_members)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (self.user_id, name, addr, ptype, bhk, sqft, rate, rent, water, meter, elec, owner_members))
        conn.commit()
        conn.close()
        self.accept()


class PropertyCardWidget(GlassCard):
    def __init__(self, prop, on_edit, on_delete, on_tap, parent=None):
        super().__init__(parent)
        self.prop = prop
        self.on_edit = on_edit
        self.on_delete = on_delete
        self.on_tap = on_tap
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.init_ui()

    def init_ui(self):
        d = self.prop
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 16, 20, 16)
        v.setSpacing(8)

        head = QHBoxLayout()
        self.name_lbl = QLabel(d[2])
        self.name_lbl.setStyleSheet(f"font-size: 16px; font-weight: 700; color: #F8FAFC; letter-spacing: 0.3px; font-family: {FONT_STACK};")

        badge_text = f"{d[4]} ({d[5]})".upper() if d[4] == "Apartment" and d[5] else d[4].upper()
        badge = GlassBadge(badge_text, "neutral")

        head.addWidget(self.name_lbl)
        head.addWidget(badge)
        head.addStretch(1)

        self.edit_btn = QPushButton()
        self.edit_btn.setIcon(SVGIconManager.icon("edit", color="#F8FAFC", size=14))
        self.edit_btn.setFixedSize(BTN_ICON, BTN_ICON)
        self.edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_btn.setStyleSheet(round_button_qss("#ffffff", size=BTN_ICON))
        self.edit_btn.clicked.connect(lambda: self.on_edit(self.prop))

        self.del_btn = QPushButton()
        self.del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
        self.del_btn.setFixedSize(BTN_ICON, BTN_ICON)
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        self.del_btn.clicked.connect(lambda: self.on_delete(d[0]))

        head.addWidget(self.edit_btn)
        head.addWidget(self.del_btn)
        v.addLayout(head)

        addr = QLabel(d[3])
        addr.setWordWrap(True)
        addr.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; font-family: {FONT_STACK};")
        v.addWidget(addr)

        summary = QHBoxLayout()
        rent = QLabel(f"<b style='font-size:16px; color:#34D399;'>{fmt_money(d[8])}</b> <span style='font-size:11px; color:{T.c('text_faint')};'>/ month</span>")
        summary.addWidget(rent)
        summary.addStretch(1)

        mt = d[10] if len(d) > 10 else "Common"
        meter_tag = "Shared Pump" if mt == "Common" else "Individual"
        w_col = ACCENT_INDIGO if d[9] else T.c("text_faint")
        e_col = ACCENT_AMBER if (d[11] if len(d) > 11 else 0) else T.c("text_faint")

        util_w = QLabel(f"<span style='color:{w_col}; font-size:12px; font-weight:600;'>Water ({meter_tag})</span> &nbsp; <span style='color:{e_col}; font-size:12px; font-weight:600;'>Electricity</span>")
        summary.addWidget(util_w)
        v.addLayout(summary)

        hint = QLabel("Click card for detailed audit")
        hint.setAlignment(Qt.AlignmentFlag.AlignRight)
        hint.setStyleSheet(f"color: {T.c('text_muted')}; font-size: 10px; font-family: {FONT_STACK};")
        v.addWidget(hint)

    def mousePressEvent(self, event):
        if self.edit_btn.geometry().contains(event.pos()) or self.del_btn.geometry().contains(event.pos()):
            super().mousePressEvent(event)
            return
        self.on_tap(self.prop)
        super().mousePressEvent(event)


class PropertyDetailDialog(GlassDialog):
    def __init__(self, user_id, prop, parent=None):
        super().__init__(parent, width=620, height=640)
        self.user_id = user_id
        self.prop = prop
        self.init_ui()

    def init_ui(self):
        d = self.prop
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("properties", color="#FFFFFF", size=20))
        t = QLabel(d[2])
        t.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: #FFFFFF; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(8)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; margin: 4px 2px 4px 0; }}
            QScrollBar::handle:vertical {{ background: rgba(255, 255, 255, 0.22); border-radius: 3px; min-height: 24px; }}
            QScrollBar::handle:vertical:hover {{ background: rgba(255, 255, 255, 0.38); }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        content = QWidget()
        content.setStyleSheet("background: transparent;")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(0, 0, 8, 20)
        cl.setSpacing(4)

        sections = [
            ("IDENTITY", [("Property ID", f"PROP-{d[0]:04d}"), ("Name", d[2]), ("Type", d[4]), ("BHK", d[5] or "—")]),
            ("FINANCIALS", [("Base Rent", fmt_money(d[8])), ("Area (sq.ft)", f"{d[6]:g}" if d[6] else "—"),
                            ("Rate per sq.ft", fmt_money(d[7]) if d[7] else "—"),
                            ("Est. Annual Yield", fmt_money((d[8] or 0) * 12))]),
            ("UTILITIES", [("Water Supply", "Available" if d[9] else "Not available"),
                           ("Water Meter", (d[10] if len(d) > 10 else "Common")),
                           ("Electric Supply", "Available" if (d[11] if len(d) > 11 else 0) else "Not available")]),
            ("LOCATION", [("Address", d[3])]),
        ]
        for i, (sec, rows) in enumerate(sections):
            if i > 0:
                divider = QFrame()
                divider.setFrameShape(QFrame.Shape.HLine)
                divider.setStyleSheet("border: none; background-color: rgba(255, 255, 255, 0.10); height: 1px;")
                cl.addWidget(divider)
            cl.addWidget(self._section(sec, rows))
        scroll.setWidget(content)
        lay.addWidget(scroll, 1)
        base.addWidget(env)

    def _section(self, title, rows):
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 6, 0, 6)
        v.setSpacing(3)
        t = QLabel(title)
        t.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {ACCENT_INDIGO}; letter-spacing: 1.0px; background: transparent; padding: 0 0 6px 0; font-family: {FONT_STACK};")
        v.addWidget(t)
        for label, value in rows:
            row = QHBoxLayout()
            row.setSpacing(14)
            l1 = QLabel(label)
            l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            l1.setFixedWidth(200)
            l2 = QLabel(str(value))
            l2.setStyleSheet(f"font-size: 12px; color: {T.c('text_primary')}; font-weight: 500; background: transparent; font-family: {FONT_STACK};")
            l2.setWordWrap(True)
            row.addWidget(l1)
            row.addWidget(l2, 1)
            v.addLayout(row)
        return w


class PropertyManagementSubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.init_ui()

    def init_ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(14)

        top = QHBoxLayout()
        tb = QVBoxLayout()
        tb.setSpacing(2)
        t = QLabel("Properties Management")
        t.setStyleSheet(f"font-size: {FS_SUB}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        s = QLabel("Register and manage your rental assets")
        s.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        tb.addWidget(t)
        tb.addWidget(s)
        top.addLayout(tb)
        top.addStretch(1)

        add = QPushButton("ADD PROPERTY")
        add.setIcon(SVGIconManager.icon("add", color="#04150E", size=16))
        add.setCursor(Qt.CursorShape.PointingHandCursor)
        add.setFixedHeight(BTN_SECONDARY)
        add.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_SECONDARY))
        add.clicked.connect(self.open_add)
        top.addWidget(add)
        v.addLayout(top)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.cards = QWidget()
        self.cards.setStyleSheet("background: transparent;")
        self.cards_lay = QVBoxLayout(self.cards)
        self.cards_lay.setContentsMargins(0, 4, 6, 4)
        self.cards_lay.setSpacing(12)
        scroll.setWidget(self.cards)
        v.addWidget(scroll)

    def load_properties(self):
        uid = self.controller.current_session_id
        if uid is None: return
        while self.cards_lay.count():
            it = self.cards_lay.takeAt(0)
            if it.widget(): it.widget().deleteLater()
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT id, user_id, name, address, property_type, bhk, area_sqft, rate_per_sqft,
                              rent, water_available, water_meter_type, electricity_available, card_color, owner_members
                       FROM properties WHERE user_id=? ORDER BY id DESC""", (uid,))
        rows = cur.fetchall()
        conn.close()
        if not rows:
            e = QLabel("No rental properties registered yet.\nClick 'ADD PROPERTY' to begin.")
            e.setAlignment(Qt.AlignmentFlag.AlignCenter)
            e.setStyleSheet(f"color: {T.c('text_faint')}; font-size: {FS_INPUT}; margin-top: 90px; line-height: 160%; background: transparent; font-family: {FONT_STACK};")
            self.cards_lay.addWidget(e)
        else:
            for r in rows:
                self.cards_lay.addWidget(PropertyCardWidget(r, self.open_edit, self.delete_prop, self.show_detail))
        self.cards_lay.addStretch(1)

    def open_add(self):
        dlg = PropertyDialog(self.controller.current_session_id, parent=self)
        if dlg.exec():
            self.load_properties()
            self.controller.dashboard_page.refresh_tab_states()

    def open_edit(self, data):
        dlg = PropertyDialog(self.controller.current_session_id, data, parent=self)
        if dlg.exec():
            self.load_properties()
            self.controller.dashboard_page.refresh_tab_states()

    def delete_prop(self, pid):
        if QMessageBox.question(self, "Delete Property", "Delete this property record?",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("DELETE FROM properties WHERE id=?", (pid,))
        conn.commit()
        conn.close()
        self.load_properties()
        self.controller.dashboard_page.refresh_tab_states()

    def show_detail(self, prop):
        PropertyDetailDialog(self.controller.current_session_id, prop, parent=self).exec()


# ==============================================================================
# § 10  TENANTS
# ==============================================================================
class TenantDialog(GlassDialog):
    def __init__(self, user_id, tenant_data=None, parent=None):
        super().__init__(parent, width=500)
        self.user_id = user_id
        self.tenant_data = tenant_data
        self.init_ui()
        if tenant_data:
            self.populate()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(9)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("tenants", color=ACCENT_INDIGO, size=18))
        t = QLabel("TENANT ONBOARDING")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.4px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        lay.addWidget(self._lbl("Select Property *"))
        self.prop_combo = QComboBox()
        self.prop_combo.setStyleSheet(INPUT_STYLE)
        self.prop_combo.setFixedHeight(INPUT_H)
        lay.addWidget(self.prop_combo)

        lay.addWidget(self._lbl("Tenant Name *"))
        self.name_in = QLineEdit()
        self.name_in.setPlaceholderText("Full legal name")
        self.name_in.setStyleSheet(INPUT_STYLE)
        self.name_in.setFixedHeight(INPUT_H)
        lay.addWidget(self.name_in)

        lay.addWidget(self._lbl("Phone (10 digits) *"))
        self.phone_in = QLineEdit()
        self.phone_in.setPlaceholderText("Contact phone")
        self.phone_in.setStyleSheet(INPUT_STYLE)
        self.phone_in.setFixedHeight(INPUT_H)
        self.phone_in.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,10}$"), self))
        lay.addWidget(self.phone_in)

        lay.addWidget(self._lbl("Permanent Address *"))
        self.address_in = QTextEdit()
        self.address_in.setPlaceholderText("Permanent residential address")
        self.address_in.setStyleSheet(INPUT_STYLE)
        self.address_in.setMaximumHeight(58)
        lay.addWidget(self.address_in)

        row1 = QHBoxLayout()
        row1.setSpacing(10)
        v1 = QVBoxLayout()
        v1.setSpacing(3)
        v1.addWidget(self._lbl("Monthly Rent (₹) *"))
        self.rent_in = QLineEdit()
        self.rent_in.setStyleSheet(INPUT_STYLE)
        self.rent_in.setFixedHeight(INPUT_H)
        self.rent_in.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        v1.addWidget(self.rent_in)

        v2 = QVBoxLayout()
        v2.setSpacing(3)
        v2.addWidget(self._lbl("Security Deposit (₹) *"))
        self.deposit_in = QLineEdit()
        self.deposit_in.setStyleSheet(INPUT_STYLE)
        self.deposit_in.setFixedHeight(INPUT_H)
        self.deposit_in.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        v2.addWidget(self.deposit_in)
        row1.addLayout(v1)
        row1.addLayout(v2)
        lay.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setSpacing(10)
        v3 = QVBoxLayout()
        v3.setSpacing(3)
        v3.addWidget(self._lbl("Entry Date *"))
        self.date_in = QDateEdit(QDate.currentDate())
        self.date_in.setCalendarPopup(True)
        setup_calendar_widget(self.date_in.calendarWidget())
        fix_stepper_margins(self.date_in)
        self.date_in.setDisplayFormat("dd/MMM/yyyy")
        self.date_in.setStyleSheet(INPUT_STYLE)
        self.date_in.setFixedHeight(INPUT_H)
        v3.addWidget(self.date_in)

        v4 = QVBoxLayout()
        v4.setSpacing(3)
        v4.addWidget(self._lbl("Occupants *"))
        self.occ_in = QSpinBox()
        setup_spinbox(self.occ_in)
        self.occ_in.setRange(1, 50)
        self.occ_in.setValue(1)
        self.occ_in.setStyleSheet(INPUT_STYLE)
        self.occ_in.setFixedHeight(INPUT_H)
        v4.addWidget(self.occ_in)
        row2.addLayout(v3)
        row2.addLayout(v4)
        lay.addLayout(row2)

        lay.addSpacing(8)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("SAVE TENANT")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save.clicked.connect(self.save_tenant)

        btn_row.addWidget(cancel, 1)
        btn_row.addWidget(save, 1)
        lay.addLayout(btn_row)

        base.addWidget(env)
        self.load_properties()

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def load_properties(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT id, name, rent FROM properties WHERE user_id=? ORDER BY id DESC", (self.user_id,))
        rows = cur.fetchall()
        conn.close()
        self.prop_combo.blockSignals(True)
        self.prop_combo.clear()
        for pid, pname, prent in rows:
            self.prop_combo.addItem(pname, pid)
        self.prop_combo.blockSignals(False)
        try:
            self.prop_combo.currentIndexChanged.disconnect()
        except TypeError:
            pass
        self.prop_combo.currentIndexChanged.connect(self.autofill_rent)
        if not self.tenant_data:
            self.autofill_rent()

    def autofill_rent(self):
        if self.tenant_data: return
        pid = self.prop_combo.currentData()
        if pid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT rent FROM properties WHERE id=?", (pid,))
        r = cur.fetchone()
        conn.close()
        if r:
            self.rent_in.setText(f"{r[0]:.2f}")

    def populate(self):
        d = self.tenant_data
        for i in range(self.prop_combo.count()):
            if self.prop_combo.itemData(i) == d[2]:
                self.prop_combo.setCurrentIndex(i)
                break
        self.name_in.setText(str(d[4]))
        self.phone_in.setText(str(d[5]))
        self.address_in.setPlainText(str(d[6]))
        self.deposit_in.setText(str(d[7]))
        self.rent_in.setText(str(d[8]))
        try:
            y, m, day = map(int, str(d[9]).split("-"))
            self.date_in.setDate(QDate(y, m, day))
        except Exception:
            pass
        self.occ_in.setValue(int(d[10]) if d[10] else 1)

    def save_tenant(self):
        pid = self.prop_combo.currentData()
        if pid is None:
            QMessageBox.warning(self, "Validation", "Register a property first.")
            return
        name = self.name_in.text().strip()
        phone = self.phone_in.text().strip()
        addr = self.address_in.toPlainText().strip()
        dep_raw = self.deposit_in.text().strip()
        rent_raw = self.rent_in.text().strip()
        if not (name and phone and addr and dep_raw and rent_raw):
            QMessageBox.warning(self, "Validation", "All required fields must be filled.")
            return
        if len(phone) != 10:
            QMessageBox.warning(self, "Validation", "Phone must be exactly 10 digits.")
            return
        try:
            dep = float(dep_raw)
            rent = float(rent_raw)
        except ValueError:
            QMessageBox.warning(self, "Validation", "Rent and deposit must be numeric.")
            return
        entry = self.date_in.date().toString("yyyy-MM-dd")
        occ = self.occ_in.value()

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        if self.tenant_data:
            cur.execute("""UPDATE tenants SET property_id=?, name=?, phone=?, permanent_address=?,
                security_deposit=?, monthly_rent=?, entry_date=?, occupants_count=? WHERE id=?""",
                (pid, name, phone, addr, dep, rent, entry, occ, self.tenant_data[0]))
        else:
            cur.execute("SELECT id FROM tenants WHERE property_id=? AND is_active=1", (pid,))
            if cur.fetchone():
                conn.close()
                QMessageBox.warning(self, "Occupied", "This property already has an active tenant.")
                return
            cur.execute("SELECT tenant_uid FROM tenants")
            used = {r[0] for r in cur.fetchall()}
            n = len(used) + 1
            while f"TNT-{n:05d}" in used:
                n += 1
            uid_str = f"TNT-{n:05d}"
            cur.execute("""INSERT INTO tenants (user_id, property_id, tenant_uid, name, phone, permanent_address,
                security_deposit, monthly_rent, entry_date, occupants_count, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                (self.user_id, pid, uid_str, name, phone, addr, dep, rent, entry, occ))
            tid_to_build = cur.lastrowid
        conn.commit()
        conn.close()

        try:
            rebuild_tenant_ledger(self.user_id, tid_to_build)
        except Exception as e:
            log.error(f"Tenant initial ledger build failed: {e}")

        try:
            record_tenant_entry_in_water_ledger(self.user_id, name, entry, occ)
        except Exception as e:
            log.error(f"Tenant initial water ledger entry failed: {e}")

        self.accept()


class TenantCardWidget(GlassCard):
    def __init__(self, tup, prop_name, on_edit, on_delete, on_toggle, on_tap, parent=None):
        super().__init__(parent)
        self.tup = tup
        self.prop_name = prop_name
        self.on_edit = on_edit
        self.on_delete = on_delete
        self.on_toggle = on_toggle
        self.on_tap = on_tap
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.init_ui()

    def init_ui(self):
        d = self.tup
        active = bool(d[11])
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 16, 20, 16)
        v.setSpacing(8)

        head = QHBoxLayout()
        self.name_lbl = QLabel(d[4])
        self.name_lbl.setStyleSheet(f"font-size: 15px; font-weight: 700; color: #F8FAFC; font-family: {FONT_STACK};")

        status_badge = GlassBadge("ACTIVE" if active else "INACTIVE", "success" if active else "danger")
        uid_badge = GlassBadge(d[3], "neutral")

        head.addWidget(self.name_lbl)
        head.addWidget(status_badge)
        head.addWidget(uid_badge)
        head.addStretch(1)

        self.edit_btn = QPushButton()
        self.edit_btn.setIcon(SVGIconManager.icon("edit", color="#F8FAFC", size=14))
        self.edit_btn.setFixedSize(BTN_ICON, BTN_ICON)
        self.edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_btn.setStyleSheet(round_button_qss("#ffffff", size=BTN_ICON))
        self.edit_btn.clicked.connect(lambda: self.on_edit(self.tup))

        self.toggle_btn = QPushButton()
        self.toggle_btn.setIcon(SVGIconManager.icon("power", color=ACCENT_INDIGO, size=14))
        self.toggle_btn.setFixedSize(BTN_ICON, BTN_ICON)
        self.toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_btn.setStyleSheet(round_button_qss(ACCENT_INDIGO, size=BTN_ICON))
        self.toggle_btn.clicked.connect(lambda: self.on_toggle(d[0], 0 if active else 1, d[2]))

        self.del_btn = QPushButton()
        self.del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
        self.del_btn.setFixedSize(BTN_ICON, BTN_ICON)
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        self.del_btn.clicked.connect(lambda: self.on_delete(d[0]))

        head.addWidget(self.edit_btn)
        head.addWidget(self.toggle_btn)
        head.addWidget(self.del_btn)
        v.addLayout(head)

        prop = QLabel(f"Property: {self.prop_name}")
        prop.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; font-family: {FONT_STACK};")
        v.addWidget(prop)

        srow = QHBoxLayout()
        srow.addWidget(QLabel(f"<b style='font-size:15px; color:#34D399;'>{fmt_money(d[8])}</b> <span style='font-size:11px; color:{T.c('text_faint')};'>/ month</span>"))
        srow.addSpacing(14)
        srow.addWidget(QLabel(f"<span style='font-size:11px; color:{T.c('text_secondary')};'>Deposit: {fmt_money(d[7])}</span>"))
        srow.addSpacing(14)
        srow.addWidget(QLabel(f"<span style='font-size:11px; color:{T.c('text_secondary')};'>{d[10]} occupants</span>"))
        srow.addStretch(1)
        v.addLayout(srow)

        hint = QLabel("Click card for detailed account view")
        hint.setAlignment(Qt.AlignmentFlag.AlignRight)
        hint.setStyleSheet(f"color: {T.c('text_muted')}; font-size: 10px; font-family: {FONT_STACK};")
        v.addWidget(hint)

    def mousePressEvent(self, event):
        if (self.edit_btn.geometry().contains(event.pos()) or
            self.del_btn.geometry().contains(event.pos()) or
            self.toggle_btn.geometry().contains(event.pos())):
            super().mousePressEvent(event)
            return
        self.on_tap(self.tup, self.prop_name)
        super().mousePressEvent(event)


class TenantDetailDialog(GlassDialog):
    def __init__(self, user_id, tup, prop_name, parent=None):
        super().__init__(parent, width=620, height=640)
        self.user_id = user_id
        self.tup = tup
        self.prop_name = prop_name
        self.init_ui()

    def init_ui(self):
        d = self.tup
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("tenants", color="#FFFFFF", size=20))
        t = QLabel(d[4])
        t.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: #FFFFFF; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(8)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; margin: 4px 2px 4px 0; }}
            QScrollBar::handle:vertical {{ background: rgba(255, 255, 255, 0.22); border-radius: 3px; min-height: 24px; }}
            QScrollBar::handle:vertical:hover {{ background: rgba(255, 255, 255, 0.38); }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        content = QWidget()
        content.setStyleSheet("background: transparent;")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(0, 0, 8, 20)
        cl.setSpacing(4)

        bal = get_tenant_balance(d[0])
        sections = [
            ("IDENTITY", [("Tenant UID", d[3]), ("Name", d[4]), ("Property", self.prop_name),
                          ("Phone", d[5]), ("Entry Date", d[9]), ("Status", "Active" if d[11] else "Inactive")]),
            ("ACCOUNT", [("Running Balance", fmt_money(bal)), ("Monthly Rent", fmt_money(d[8])),
                         ("Security Deposit", fmt_money(d[7])), ("Occupants", d[10])]),
            ("PERMANENT ADDRESS", [("Address", d[6])]),
        ]
        for i, (sec, rows) in enumerate(sections):
            if i > 0:
                divider = QFrame()
                divider.setFrameShape(QFrame.Shape.HLine)
                divider.setStyleSheet(f"border: none; background-color: {T.c('divider')}; height: 1px;")
                cl.addWidget(divider)
            cl.addWidget(self._section(sec, rows))
        scroll.setWidget(content)
        lay.addWidget(scroll, 1)
        base.addWidget(env)

    def _section(self, title, rows):
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 6, 0, 6)
        v.setSpacing(3)
        t = QLabel(title)
        t.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {ACCENT_INDIGO}; letter-spacing: 1.0px; background: transparent; padding: 0 0 6px 0; font-family: {FONT_STACK};")
        v.addWidget(t)
        for label, value in rows:
            row = QHBoxLayout()
            row.setSpacing(14)
            l1 = QLabel(label)
            l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            l1.setFixedWidth(200)
            l2 = QLabel(str(value))
            l2.setStyleSheet(f"font-size: 12px; color: {T.c('text_primary')}; font-weight: 500; background: transparent; font-family: {FONT_STACK};")
            l2.setWordWrap(True)
            row.addWidget(l1)
            row.addWidget(l2, 1)
            v.addLayout(row)
        return w


class TenantManagementSubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.init_ui()

    def init_ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(14)

        top = QHBoxLayout()
        tb = QVBoxLayout()
        tb.setSpacing(2)
        t = QLabel("Tenant Onboarding & Allocation")
        t.setStyleSheet(f"font-size: {FS_SUB}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        s = QLabel("Allocate tenants, track leases, enforce single occupancy")
        s.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        tb.addWidget(t)
        tb.addWidget(s)
        top.addLayout(tb)
        top.addStretch(1)

        add = QPushButton("ONBOARD TENANT")
        add.setIcon(SVGIconManager.icon("add", color="#04150E", size=16))
        add.setCursor(Qt.CursorShape.PointingHandCursor)
        add.setFixedHeight(BTN_SECONDARY)
        add.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_SECONDARY))
        add.clicked.connect(self.open_add)
        top.addWidget(add)
        v.addLayout(top)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.cards = QWidget()
        self.cards.setStyleSheet("background: transparent;")
        self.cards_lay = QVBoxLayout(self.cards)
        self.cards_lay.setContentsMargins(0, 4, 6, 4)
        self.cards_lay.setSpacing(12)
        scroll.setWidget(self.cards)
        v.addWidget(scroll)

    def load_tenants(self):
        uid = self.controller.current_session_id
        if uid is None: return
        while self.cards_lay.count():
            it = self.cards_lay.takeAt(0)
            if it.widget(): it.widget().deleteLater()
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.id, t.user_id, t.property_id, t.tenant_uid, t.name, t.phone,
                       t.permanent_address, t.security_deposit, t.monthly_rent,
                       t.entry_date, t.occupants_count, t.is_active, p.name
                       FROM tenants t JOIN properties p ON t.property_id = p.id
                       WHERE t.user_id=? ORDER BY t.is_active DESC, t.id DESC""", (uid,))
        rows = cur.fetchall()
        conn.close()
        if not rows:
            e = QLabel("No tenants onboarded yet.\nClick 'ONBOARD TENANT' to begin.")
            e.setAlignment(Qt.AlignmentFlag.AlignCenter)
            e.setStyleSheet(f"color: {T.c('text_faint')}; font-size: {FS_INPUT}; margin-top: 90px; line-height: 160%; background: transparent; font-family: {FONT_STACK};")
            self.cards_lay.addWidget(e)
        else:
            for r in rows:
                self.cards_lay.addWidget(TenantCardWidget(r[:12], r[12], self.open_edit, self.delete_t, self.toggle_active, self.show_detail))
        self.cards_lay.addStretch(1)

    def open_add(self):
        dlg = TenantDialog(self.controller.current_session_id, parent=self)
        if dlg.exec():
            self.load_tenants()
            self.controller.dashboard_page.refresh_tab_states()

    def open_edit(self, data):
        dlg = TenantDialog(self.controller.current_session_id, data, parent=self)
        if dlg.exec():
            self.load_tenants()
            self.controller.dashboard_page.refresh_tab_states()

    def toggle_active(self, tid, new_status, prop_id):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        if new_status == 1:
            cur.execute("SELECT id FROM tenants WHERE property_id=? AND is_active=1 AND id!=?", (prop_id, tid))
            if cur.fetchone():
                QMessageBox.warning(self, "Occupied", "This property already has another active tenant.")
                conn.close()
                return
        cur.execute("UPDATE tenants SET is_active=? WHERE id=?", (1 if new_status else 0, tid))
        conn.commit()
        conn.close()
        self.load_tenants()
        self.controller.dashboard_page.refresh_tab_states()

    def delete_t(self, tid):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.name, t.tenant_uid, p.address, p.name
                       FROM tenants t JOIN properties p ON t.property_id=p.id
                       WHERE t.id=?""", (tid,))
        row = cur.fetchone()
        if not row:
            conn.close()
            return
        tname, tuid, addr, pname = row

        reply = QMessageBox.question(
            self, "Delete Tenant",
            f"Delete {tname} AND ALL their bills, payments, and records?\n\n"
            f"This will permanently remove their history and files.\n"
            f"Action cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            conn.close()
            return

        cur.execute("DELETE FROM tenant_bills WHERE tenant_id=?", (tid,))
        cur.execute("DELETE FROM tenant_transactions WHERE tenant_id=?", (tid,))
        cur.execute("DELETE FROM tenants WHERE id=?", (tid,))
        conn.commit()
        conn.close()

        try:
            tdir = get_tenant_dir(self.controller.current_session_id, addr, pname, tname, tuid)
            if os.path.exists(tdir):
                shutil.rmtree(tdir, ignore_errors=True)
                log.info(f"Deleted tenant folder: {tdir}")
        except Exception as e:
            log.error(f"Folder delete failed: {e}")

        self.load_tenants()
        self.controller.dashboard_page.refresh_tab_states()
        self.controller.dashboard_page.refresh_live_metrics()

    def show_detail(self, tup, prop_name):
        TenantDetailDialog(self.controller.current_session_id, tup, prop_name, parent=self).exec()


# ==============================================================================
# § 11  WATER RECORD
# ==============================================================================
class ExtraPersonDialog(GlassDialog):
    def __init__(self, parent=None, existing=None):
        super().__init__(parent, width=420)
        self.result_data = None
        self.init_ui(existing)

    def init_ui(self, existing):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(26, 22, 26, 22)
        lay.setSpacing(10)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("water", color=ACCENT_AMBER, size=18))
        t = QLabel("EXTRA OCCUPANT (ONE-OFF)")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        lay.addLayout(head)

        hint = QLabel("For someone who shared the pump this month but is not in "
                      "the active tenant list — e.g. a tenant who left mid-month.")
        hint.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; background: transparent; font-family: {FONT_STACK};")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        lay.addWidget(self._lbl("Name / Label"))
        self.name_in = QLineEdit()
        self.name_in.setPlaceholderText("e.g. Sanjib (leaving)")
        self.name_in.setStyleSheet(INPUT_STYLE)
        self.name_in.setFixedHeight(INPUT_H)
        lay.addWidget(self.name_in)

        lay.addWidget(self._lbl("Headcount"))
        self.hc_in = QSpinBox()
        setup_spinbox(self.hc_in)
        self.hc_in.setRange(1, 50)
        self.hc_in.setValue(1)
        self.hc_in.setStyleSheet(INPUT_STYLE)
        self.hc_in.setFixedHeight(INPUT_H)
        lay.addWidget(self.hc_in)

        if existing:
            self.name_in.setText(existing.get("name", ""))
            self.hc_in.setValue(int(existing.get("headcount", 1)))

        lay.addSpacing(6)
        row = QHBoxLayout()
        row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        ok = QPushButton("SAVE")
        ok.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        ok.setCursor(Qt.CursorShape.PointingHandCursor)
        ok.setFixedHeight(BTN_PRIMARY)
        ok.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        ok.clicked.connect(self._on_ok)

        row.addWidget(cancel, 1)
        row.addWidget(ok, 1)
        lay.addLayout(row)
        base.addWidget(env)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def _on_ok(self):
        name = self.name_in.text().strip()
        if not name:
            QMessageBox.warning(self, "Required", "Name is required.")
            return
        self.result_data = {"name": name, "headcount": int(self.hc_in.value())}
        self.accept()


class WaterBillsViewerDialog(GlassDialog):
    def __init__(self, controller, billing_month, parent=None):
        super().__init__(parent, width=640, height=640)
        self.controller = controller
        self.billing_month = billing_month
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("water", color=ACCENT_INDIGO, size=22))
        t = QLabel(f"Water Allocation — {self.billing_month}")
        t.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        uid = self.controller.current_session_id
        rec = fetch_campus_water_record(uid, self.billing_month)
        if not rec:
            empty = QLabel(
                f"No water record exists for {self.billing_month}.\n\n"
                f"Save a pump reading in this tab first."
            )
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty.setStyleSheet(f"color: {T.c('text_secondary')}; font-size: 12px; padding: 40px; background: transparent; font-family: {FONT_STACK};")
            empty.setWordWrap(True)
            lay.addWidget(empty)
            base.addWidget(env)
            return

        try:
            breakdown = json.loads(rec["breakdown_json"]) if rec["breakdown_json"] else {}
        except Exception:
            breakdown = {}

        owner_family = int(breakdown.get("owner_family", 0) or 0)
        tenant_families = breakdown.get("tenant_families", {}) or {}
        extra_persons = breakdown.get("extra_persons", []) or []

        total_bill = float(rec["total_bill"])
        total_persons = int(rec["total_persons"])
        per_head = float(rec["rate_per_person"])
        prev_r = float(rec["prev_reading"])
        pres_r = float(rec["pres_reading"])
        rate_e = float(rec["rate_elec"])
        units = max(0.0, pres_r - prev_r)

        sum_card = GlassCard()
        sl = QGridLayout(sum_card)
        sl.setContentsMargins(20, 16, 20, 16)
        sl.setHorizontalSpacing(24)
        sl.setVerticalSpacing(6)

        def add_kv(row, col, label, value, color=None, bold=False):
            if color is None: color = T.c("text_primary")
            l1 = QLabel(label)
            l1.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.5px; background: transparent; font-family: {FONT_STACK};")
            l2 = QLabel(value)
            weight = "800" if bold else "600"
            l2.setStyleSheet(f"font-size: 13px; font-weight: {weight}; color: {color}; background: transparent; font-family: {FONT_STACK};")
            sl.addWidget(l1, row * 2, col)
            sl.addWidget(l2, row * 2 + 1, col)

        add_kv(0, 0, "READING DATE", rec["reading_date"] or "—")
        add_kv(0, 1, "UNITS CONSUMED", f"{units:,.2f}")
        add_kv(0, 2, "RATE / UNIT", f"₹{rate_e:,.2f}")
        add_kv(1, 0, "PUMP BILL", fmt_money(total_bill), ACCENT_AMBER, bold=True)
        add_kv(1, 1, "TOTAL PERSONS", str(total_persons))
        add_kv(1, 2, "PER-HEAD COST", fmt_money(per_head), ACCENT_GREEN, bold=True)
        lay.addWidget(sum_card)

        bill_lbl = QLabel("PER-PARTY ALLOCATED SHARES")
        bill_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.8px; background: transparent; font-family: {FONT_STACK};")
        lay.addWidget(bill_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 20px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        cont = QWidget()
        cont.setStyleSheet("background: transparent;")
        cl = QVBoxLayout(cont)
        cl.setContentsMargins(0, 0, 6, 0)
        cl.setSpacing(6)

        hdr = QFrame()
        hdr.setStyleSheet("QFrame { background-color: rgba(255, 255, 255, 0.05); border-radius: 8px; }")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(14, 8, 14, 8)
        hl.setSpacing(10)
        for text, width in [("Party", 220), ("Family", 70), ("Guests", 70), ("Heads", 70), ("Share", 110)]:
            lab = QLabel(text.upper())
            lab.setStyleSheet(f"font-size: 9px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.5px; background: transparent; font-family: {FONT_STACK};")
            lab.setFixedWidth(width)
            if text in ("Family", "Guests", "Heads", "Share"):
                lab.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            hl.addWidget(lab)
        hl.addStretch(1)
        cl.addWidget(hdr)

        def add_row(name, family, guests, share_color=None):
            if share_color is None: share_color = T.c("text_primary")
            heads = family + guests
            share = heads * per_head
            row = QFrame()
            row.setStyleSheet("""
                QFrame {
                    background-color: rgba(40, 40, 40, 0.60);
                    border: 1px solid rgba(255, 255, 255, 0.18);
                    border-radius: 12px;
                }
                QFrame:hover {
                    background-color: rgba(50, 50, 50, 0.72);
                    border: 1px solid rgba(255, 255, 255, 0.32);
                }
            """)
            rl = QHBoxLayout(row)
            rl.setContentsMargins(14, 8, 14, 8)
            rl.setSpacing(10)

            n = QLabel(name)
            n.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {share_color}; background: transparent; font-family: {FONT_STACK};")
            n.setFixedWidth(220)
            rl.addWidget(n)

            for val in [str(family), str(guests), str(heads)]:
                v = QLabel(val)
                v.setStyleSheet(f"font-size: 12px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
                v.setFixedWidth(70)
                v.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                rl.addWidget(v)

            s = QLabel(fmt_money(share))
            s.setStyleSheet(f"font-size: 12px; font-weight: 700; color: {ACCENT_GREEN}; background: transparent; font-family: {FONT_STACK};")
            s.setFixedWidth(110)
            s.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            rl.addWidget(s)

            rl.addStretch(1)
            cl.addWidget(row)

        add_row("Owner Family", owner_family, 0, ACCENT_INDIGO)
        for t_name, t_family in tenant_families.items():
            add_row(t_name, int(t_family), 0)
        for ep in extra_persons:
            add_row(f"{ep.get('name', 'Extra')} (one-off)", 0, int(ep.get("headcount", 0)), ACCENT_AMBER)

        total_row = QFrame()
        total_row.setStyleSheet("QFrame { background-color: rgba(16, 185, 129, 0.10); border: 1px solid rgba(16, 185, 129, 0.35); border-radius: 8px; }")
        tl = QHBoxLayout(total_row)
        tl.setContentsMargins(14, 8, 14, 8)
        tl.setSpacing(10)

        tot_name = QLabel("TOTAL")
        tot_name.setStyleSheet(f"font-size: 12px; font-weight: 800; color: {ACCENT_GREEN}; background: transparent; font-family: {FONT_STACK};")
        tot_name.setFixedWidth(220)
        tl.addWidget(tot_name)

        empty = QLabel(""); empty.setFixedWidth(70); tl.addWidget(empty)
        empty = QLabel(""); empty.setFixedWidth(70); tl.addWidget(empty)

        tot_heads = QLabel(str(total_persons))
        tot_heads.setStyleSheet(f"font-size: 12px; font-weight: 800; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        tot_heads.setFixedWidth(70)
        tot_heads.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        tl.addWidget(tot_heads)

        tot_share = QLabel(fmt_money(total_bill))
        tot_share.setStyleSheet(f"font-size: 13px; font-weight: 800; color: {ACCENT_GREEN}; background: transparent; font-family: {FONT_STACK};")
        tot_share.setFixedWidth(110)
        tot_share.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        tl.addWidget(tot_share)
        tl.addStretch(1)
        cl.addWidget(total_row)

        if rec.get("remarks"):
            note_lbl = QLabel("REMARKS")
            note_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.8px; background: transparent; font-family: {FONT_STACK}; margin-top: 8px;")
            cl.addWidget(note_lbl)
            note_box = QLabel(rec["remarks"])
            note_box.setStyleSheet(f"background: rgba(245,158,11,0.08); border: 1px solid rgba(245,158,11,0.30); border-radius: 8px; padding: 10px 12px; color: {ACCENT_AMBER}; font-size: 12px; font-family: {FONT_STACK};")
            note_box.setWordWrap(True)
            cl.addWidget(note_box)

        cl.addStretch(1)
        scroll.setWidget(cont)
        lay.addWidget(scroll, 1)

        footer = QHBoxLayout()
        footer.addStretch(1)
        close_btn = QPushButton("CLOSE")
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setFixedHeight(BTN_SECONDARY)
        close_btn.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_SECONDARY))
        close_btn.clicked.connect(self.accept)
        footer.addWidget(close_btn)
        lay.addLayout(footer)

        base.addWidget(env)


class WaterRecordSubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.extra_persons = []
        self.tenant_rows = []
        self.owner_family = 1
        self.owner_name = "Owner"
        self.computed_total_bill = 0.0
        self.computed_total_persons = 0
        self.computed_rate_per_person = 0.0
        self.init_ui()

    def _lbl(self, text, color=None):
        if color is None: color = T.c("text_faint")
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {color}; letter-spacing: 0.5px; background: transparent; font-family: {FONT_STACK};")
        return l

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 12, 18, 18)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        c = QWidget()
        c.setStyleSheet(card_qss() + input_qss())
        v = QVBoxLayout(c)
        v.setSpacing(14)
        v.setContentsMargins(0, 0, 6, 0)

        hdr = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("water", color=ACCENT_INDIGO, size=24))
        tl = QLabel("Campus Water Record")
        tl.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        sub = QLabel("One pump, one reading, one per-head — shared across the campus")
        sub.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        vt = QVBoxLayout()
        vt.setSpacing(2)
        vt.addWidget(tl)
        vt.addWidget(sub)
        hdr.addWidget(ic)
        hdr.addSpacing(6)
        hdr.addLayout(vt)
        hdr.addStretch(1)
        v.addLayout(hdr)

        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        cl.addWidget(self._lbl("BILLING CYCLE *"))

        self.lbl_lock = QLabel("")
        self.lbl_lock.setStyleSheet(f"""
            background-color: rgba(245,158,11,0.12); color: {ACCENT_AMBER};
            border: 1px solid rgba(245,158,11,0.35); padding: 8px 12px;
            border-radius: {RADIUS_INPUT}px; font-size: 11px; font-weight: 600;
            font-family: {FONT_STACK};
        """)
        self.lbl_lock.setWordWrap(True)
        self.lbl_lock.hide()
        cl.addWidget(self.lbl_lock)

        row = QHBoxLayout()
        row.setSpacing(14)
        vm = QVBoxLayout()
        vm.setSpacing(4)
        vm.addWidget(self._lbl("BILL MONTH *"))
        self.combo_month = QComboBox()
        setup_combobox(self.combo_month)
        self.combo_month.addItems(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
        self.combo_month.setCurrentIndex(QDate.currentDate().month() - 1)
        self.combo_month.setStyleSheet(INPUT_STYLE)
        self.combo_month.setFixedHeight(INPUT_H)
        self.combo_month.currentIndexChanged.connect(self.on_month_changed)
        vm.addWidget(self.combo_month)

        vy = QVBoxLayout()
        vy.setSpacing(4)
        vy.addWidget(self._lbl("BILL YEAR *"))
        self.combo_year = QComboBox()
        setup_combobox(self.combo_year)
        cy = QDate.currentDate().year()
        for y in range(cy - 2, cy + 5): self.combo_year.addItem(str(y))
        self.combo_year.setCurrentText(str(cy))
        self.combo_year.setStyleSheet(INPUT_STYLE)
        self.combo_year.setFixedHeight(INPUT_H)
        self.combo_year.currentIndexChanged.connect(self.on_month_changed)
        vy.addWidget(self.combo_year)

        vd = QVBoxLayout()
        vd.setSpacing(4)
        vd.addWidget(self._lbl("READING DATE *"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setStyleSheet(INPUT_STYLE)
        self.in_date.setFixedHeight(INPUT_H)
        vd.addWidget(self.in_date)

        row.addLayout(vm, 1)
        row.addLayout(vy, 1)
        row.addLayout(vd, 1)
        cl.addLayout(row)
        v.addWidget(card)

        pc = GlassCard()
        pl = QVBoxLayout(pc)
        pl.setContentsMargins(20, 16, 20, 16)
        pl.setSpacing(12)

        phead = QHBoxLayout()
        pt = QLabel("Pump Electricity Meter")
        pt.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {ACCENT_AMBER}; font-family: {FONT_STACK};")
        phead.addWidget(pt)
        phead.addStretch(1)
        self.btn_autofill = QPushButton("AUTO-FILL FROM LAST MONTH")
        self.btn_autofill.setIcon(SVGIconManager.icon("refresh", color=ACCENT_INDIGO, size=14))
        self.btn_autofill.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_autofill.setFixedHeight(BTN_INLINE)
        self.btn_autofill.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_INLINE))
        self.btn_autofill.clicked.connect(self.autofill_from_last_month)
        phead.addWidget(self.btn_autofill)
        pl.addLayout(phead)

        r1 = QHBoxLayout()
        r1.setSpacing(12)
        v1 = QVBoxLayout(); v1.setSpacing(4); v1.addWidget(self._lbl("PREVIOUS READING"))
        self.in_prev = QLineEdit("0"); self.in_prev.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_prev.setStyleSheet(INPUT_STYLE); self.in_prev.setFixedHeight(INPUT_H)
        self.in_prev.textChanged.connect(self.recalc)
        v1.addWidget(self.in_prev)

        v2 = QVBoxLayout(); v2.setSpacing(4); v2.addWidget(self._lbl("PRESENT READING"))
        self.in_pres = QLineEdit("0"); self.in_pres.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_pres.setStyleSheet(INPUT_STYLE); self.in_pres.setFixedHeight(INPUT_H)
        self.in_pres.textChanged.connect(self.recalc)
        v2.addWidget(self.in_pres)

        v3 = QVBoxLayout(); v3.setSpacing(4); v3.addWidget(self._lbl("RATE / UNIT (₹)"))
        self.in_rate = QLineEdit("9.50"); self.in_rate.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
        self.in_rate.setStyleSheet(INPUT_STYLE); self.in_rate.setFixedHeight(INPUT_H)
        self.in_rate.textChanged.connect(self.recalc)
        v3.addWidget(self.in_rate)
        r1.addLayout(v1, 1); r1.addLayout(v2, 1); r1.addLayout(v3, 1)
        pl.addLayout(r1)

        self.lbl_pump_computed = QLabel("Units consumed: 0.00  ·  Total pump bill: ₹0.00")
        self.lbl_pump_computed.setStyleSheet(f"""
            background-color: rgba(245, 158, 11, 0.10);
            color: {ACCENT_AMBER};
            border: 1px solid rgba(245, 158, 11, 0.28);
            border-radius: {RADIUS_INPUT}px;
            padding: 8px 14px;
            font-size: 12px; font-weight: 700;
            font-family: {FONT_STACK};
        """)
        pl.addWidget(self.lbl_pump_computed)
        v.addWidget(pc)

        hc = GlassCard()
        hl = QVBoxLayout(hc)
        hl.setContentsMargins(20, 16, 20, 16)
        hl.setSpacing(10)
        ht = QLabel("Campus Headcount Breakdown")
        ht.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {ACCENT_INDIGO}; font-family: {FONT_STACK};")
        hl.addWidget(ht)

        om_row = QHBoxLayout()
        om_row.setSpacing(14)
        vom = QVBoxLayout()
        vom.setSpacing(4)
        vom.addWidget(self._lbl("OWNER FAMILY MEMBERS *"))
        self.in_owner = QSpinBox()
        self.in_owner.setRange(0, 50)
        self.in_owner.setValue(1)
        self.in_owner.setStyleSheet(INPUT_STYLE)
        self.in_owner.setFixedHeight(INPUT_H)
        self.in_owner.setFixedWidth(120)
        setup_spinbox(self.in_owner)
        self.in_owner.valueChanged.connect(self.recalc)
        vom.addWidget(self.in_owner)
        om_row.addLayout(vom)

        om_info = QLabel("Owner family headcount on campus · Shared pump reading will be allocated proportionally")
        om_info.setStyleSheet(f"color: {T.c('text_dim')}; font-size: 12px; font-family: {FONT_STACK}; margin-top: 14px;")
        om_row.addWidget(om_info)
        om_row.addStretch(1)
        hl.addLayout(om_row)

        hl.addWidget(self._lbl("ACTIVE TENANTS (FAMILY SIZE — EDITABLE FOR THIS MONTH)"))
        self.tenant_list_widget = QWidget()
        self.tenant_list_widget.setStyleSheet("background: transparent;")
        self.tenant_list_lay = QVBoxLayout(self.tenant_list_widget)
        self.tenant_list_lay.setContentsMargins(0, 0, 0, 0)
        self.tenant_list_lay.setSpacing(6)
        hl.addWidget(self.tenant_list_widget)

        hl.addSpacing(4)
        hl.addWidget(self._lbl("EXTRA PERSONS (THIS MONTH ONLY)"))
        self.extra_list_widget = QWidget()
        self.extra_list_widget.setStyleSheet("background: transparent;")
        self.extra_list_lay = QVBoxLayout(self.extra_list_widget)
        self.extra_list_lay.setContentsMargins(0, 0, 0, 0)
        self.extra_list_lay.setSpacing(6)
        hl.addWidget(self.extra_list_widget)

        add_extra_btn = QPushButton("ADD EXTRA PERSON")
        add_extra_btn.setIcon(SVGIconManager.icon("add", color=ACCENT_AMBER, size=14))
        add_extra_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_extra_btn.setFixedHeight(BTN_INLINE)
        add_extra_btn.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_INLINE))
        add_extra_btn.clicked.connect(self.add_extra_person)
        hl.addWidget(add_extra_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        self.lbl_headcount_summary = QLabel("Total persons: 0  ·  Per-head cost: ₹0.00")
        self.lbl_headcount_summary.setStyleSheet(f"""
            background-color: rgba(16, 185, 129, 0.10);
            color: {ACCENT_GREEN};
            border: 1px solid rgba(16, 185, 129, 0.30);
            border-radius: {RADIUS_INPUT}px;
            padding: 10px 16px;
            font-size: 13px; font-weight: 700;
            font-family: {FONT_STACK};
        """)
        hl.addWidget(self.lbl_headcount_summary)
        v.addWidget(hc)

        rc = GlassCard()
        rl = QVBoxLayout(rc)
        rl.setContentsMargins(20, 16, 20, 16)
        rl.setSpacing(10)
        rt = QLabel("Remarks")
        rt.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        rl.addWidget(rt)
        self.in_remarks = QLineEdit()
        self.in_remarks.setPlaceholderText("Optional note for this month's pump reading…")
        self.in_remarks.setStyleSheet(INPUT_STYLE)
        self.in_remarks.setFixedHeight(INPUT_H)
        rl.addWidget(self.in_remarks)
        v.addWidget(rc)

        bc = GlassCard()
        bl = QHBoxLayout(bc)
        bl.setContentsMargins(20, 16, 20, 16)
        bl.setSpacing(10)

        self.btn_save = QPushButton("SAVE WATER RECORD")
        self.btn_save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        self.btn_save.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_save.setFixedHeight(BTN_PRIMARY)
        self.btn_save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        self.btn_save.clicked.connect(self.save_record)

        self.btn_view_bills = QPushButton("VIEW MONTH'S BILLS")
        self.btn_view_bills.setIcon(SVGIconManager.icon("file", color=ACCENT_AMBER, size=16))
        self.btn_view_bills.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_view_bills.setFixedHeight(BTN_PRIMARY)
        self.btn_view_bills.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_PRIMARY))
        self.btn_view_bills.clicked.connect(self.view_month_bills)

        self.btn_open_ledger = QPushButton("OPEN EXCEL LEDGER")
        self.btn_open_ledger.setIcon(SVGIconManager.icon("file_excel", color=ACCENT_INDIGO, size=16))
        self.btn_open_ledger.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_open_ledger.setFixedHeight(BTN_PRIMARY)
        self.btn_open_ledger.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        self.btn_open_ledger.clicked.connect(self.open_ledger)

        bl.addWidget(self.btn_save)
        bl.addWidget(self.btn_view_bills)
        bl.addWidget(self.btn_open_ledger)
        bl.addStretch(1)
        v.addWidget(bc)

        scroll.setWidget(c)
        root.addWidget(scroll)

    def refresh_page_data(self):
        uid = self.controller.current_session_id
        if uid is None: return

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT owner_name FROM profiles WHERE user_id=?", (uid,))
        prow = cur.fetchone()
        self.owner_name = prow[0] if (prow and prow[0]) else "Owner"

        cur.execute("""SELECT COALESCE(MAX(owner_members), 1) FROM properties WHERE user_id=?""", (uid,))
        orow = cur.fetchone()
        self.owner_family = int(orow[0]) if (orow and orow[0]) else 1

        cur.execute("""SELECT id, name, occupants_count FROM tenants
                       WHERE user_id=? AND is_active=1 ORDER BY tenant_uid""", (uid,))
        trows = cur.fetchall()
        conn.close()

        self.tenant_rows = [
            {"tenant_id": r[0], "name": r[1], "family": int(r[2] or 0)}
            for r in trows
        ]
        self.in_owner.setValue(self.owner_family)
        self.render_tenant_list()
        self.on_month_changed()

    def on_month_changed(self):
        uid = self.controller.current_session_id
        if uid is None: return
        billing_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        rec = fetch_campus_water_record(uid, billing_month)
        if rec:
            self.in_prev.setText(str(rec["prev_reading"]))
            self.in_pres.setText(str(rec["pres_reading"]))
            self.in_rate.setText(str(rec["rate_elec"]))
            try:
                self.in_date.setDate(QDate.fromString(rec["reading_date"], "yyyy-MM-dd"))
            except Exception: pass
            self.in_remarks.setText(rec["remarks"] or "")
            try:
                b = json.loads(rec["breakdown_json"]) if rec["breakdown_json"] else {}
            except Exception:
                b = {}
            if b:
                self.in_owner.setValue(int(b.get("owner_family", self.owner_family)))
                for tr in self.tenant_rows:
                    if tr["name"] in b.get("tenant_families", {}):
                        tr["family"] = int(b["tenant_families"][tr["name"]])
                self.extra_persons = b.get("extra_persons", []) or []
                self.render_tenant_list()
                self.render_extra_list()
            self.lbl_lock.setText(f"Locked for {billing_month}. Saving will UPDATE it. Past bills remain frozen.")
            self.lbl_lock.show()
        else:
            self.lbl_lock.setText(f"No record yet for {billing_month}. Saving will CREATE it.")
            self.lbl_lock.show()
        self.recalc()

    def render_tenant_list(self):
        while self.tenant_list_lay.count():
            it = self.tenant_list_lay.takeAt(0)
            if it.widget(): it.widget().deleteLater()
        if not self.tenant_rows:
            lab = QLabel("— No active tenants on the campus —")
            lab.setStyleSheet(f"font-size: 12px; color: {T.c('text_secondary')}; padding: 6px; font-family: {FONT_STACK};")
            self.tenant_list_lay.addWidget(lab)
            return
        for tr in self.tenant_rows:
            row = QFrame()
            row.setStyleSheet("""
                QFrame {
                    background-color: rgba(40, 40, 40, 0.60);
                    border: 1px solid rgba(255, 255, 255, 0.18);
                    border-radius: 12px;
                }
                QFrame:hover {
                    background-color: rgba(50, 50, 50, 0.72);
                    border: 1px solid rgba(255, 255, 255, 0.32);
                }
            """)
            rl = QHBoxLayout(row)
            rl.setContentsMargins(16, 10, 16, 10)
            rl.setSpacing(14)

            name_icon = QLabel()
            name_icon.setPixmap(SVGIconManager.pixmap("owner", color=ACCENT_INDIGO, size=18))
            rl.addWidget(name_icon)

            name = QLabel(f"<b>{tr['name']}</b>")
            name.setStyleSheet(f"font-size: 13px; font-family: {FONT_STACK}; color: #FFFFFF;")
            name.setFixedWidth(160)
            rl.addWidget(name)

            badge = GlassBadge("ACTIVE LEASE", "success")
            rl.addWidget(badge)

            rl.addSpacing(16)

            occ_lbl = QLabel("Family Occupants:")
            occ_lbl.setStyleSheet(f"color: {T.c('text_secondary')}; font-size: 12px; font-weight: 500; font-family: {FONT_STACK};")
            rl.addWidget(occ_lbl)

            spin = QSpinBox()
            spin.setRange(0, 50)
            spin.setValue(int(tr["family"]))
            spin.setStyleSheet(INPUT_STYLE)
            spin.setFixedHeight(34)
            spin.setFixedWidth(90)
            setup_spinbox(spin)
            spin.valueChanged.connect(lambda v, n=tr["name"]: self._on_tenant_family_change(n, v))
            rl.addWidget(spin)

            persons_lbl = QLabel("persons")
            persons_lbl.setStyleSheet(f"color: {T.c('text_secondary')}; font-size: 12px; font-weight: 500; font-family: {FONT_STACK};")
            rl.addWidget(persons_lbl)

            rl.addStretch(1)
            self.tenant_list_lay.addWidget(row)

    def _on_tenant_family_change(self, tenant_name, value):
        for tr in self.tenant_rows:
            if tr["name"] == tenant_name:
                tr["family"] = int(value)
                break
        self.recalc()

    def render_extra_list(self):
        while self.extra_list_lay.count():
            it = self.extra_list_lay.takeAt(0)
            if it.widget(): it.widget().deleteLater()
        if not self.extra_persons:
            lab = QLabel("— None —")
            lab.setStyleSheet(f"font-size: 12px; color: {T.c('text_muted')}; padding: 4px; font-family: {FONT_STACK};")
            self.extra_list_lay.addWidget(lab)
            return
        for i, ep in enumerate(self.extra_persons):
            row = QFrame()
            row.setStyleSheet("""
                QFrame {
                    background-color: rgba(40, 40, 40, 0.60);
                    border: 1px solid rgba(255, 106, 0, 0.35);
                    border-radius: 12px;
                }
                QFrame:hover {
                    background-color: rgba(50, 50, 50, 0.72);
                    border: 1px solid rgba(255, 106, 0, 0.55);
                }
            """)
            rl = QHBoxLayout(row)
            rl.setContentsMargins(14, 8, 14, 8)
            rl.setSpacing(12)
            name = QLabel(f"<b style='color:{T.c('text_primary')};'>{ep['name']}</b>")
            name.setStyleSheet(f"font-size: 12px; font-family: {FONT_STACK};")
            hc = QLabel(f"{ep['headcount']} persons")
            hc.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-family: {FONT_STACK};")

            edit_btn = QPushButton()
            edit_btn.setIcon(SVGIconManager.icon("edit", color="#F8FAFC", size=14))
            edit_btn.setFixedSize(BTN_ICON, BTN_ICON)
            edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            edit_btn.setStyleSheet(round_button_qss("#ffffff", size=BTN_ICON))
            edit_btn.clicked.connect(lambda _, idx=i: self.edit_extra_person(idx))

            del_btn = QPushButton()
            del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
            del_btn.setFixedSize(BTN_ICON, BTN_ICON)
            del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            del_btn.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
            del_btn.clicked.connect(lambda _, idx=i: self.remove_extra_person(idx))

            rl.addWidget(name)
            rl.addStretch(1)
            rl.addWidget(hc)
            rl.addWidget(edit_btn)
            rl.addWidget(del_btn)
            self.extra_list_lay.addWidget(row)

    def add_extra_person(self):
        dlg = ExtraPersonDialog(self)
        if dlg.exec() and dlg.result_data:
            self.extra_persons.append(dlg.result_data)
            self.render_extra_list()
            self.recalc()

    def edit_extra_person(self, idx):
        if idx < 0 or idx >= len(self.extra_persons): return
        dlg = ExtraPersonDialog(self, existing=self.extra_persons[idx])
        if dlg.exec() and dlg.result_data:
            self.extra_persons[idx] = dlg.result_data
            self.render_extra_list()
            self.recalc()

    def remove_extra_person(self, idx):
        if 0 <= idx < len(self.extra_persons):
            self.extra_persons.pop(idx)
            self.render_extra_list()
            self.recalc()

    def autofill_from_last_month(self):
        uid = self.controller.current_session_id
        if uid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT pres_reading, rate_elec FROM campus_water_records
                       WHERE user_id=? ORDER BY id DESC LIMIT 1""", (uid,))
        r = cur.fetchone()
        conn.close()
        if r:
            self.in_prev.setText(str(r[0]))
            self.in_rate.setText(str(r[1]))
        else:
            QMessageBox.information(self, "No history", "No previous water record exists.")

    def recalc(self):
        try:
            prev = float(self.in_prev.text() or 0)
            pres = float(self.in_pres.text() or 0)
            rate = float(self.in_rate.text() or 0)
            units = max(0.0, pres - prev)
            self.computed_total_bill = units * rate
            self.lbl_pump_computed.setText(
                f"Units consumed: {units:,.2f}  ·  Total pump bill: {fmt_money(self.computed_total_bill)}")
        except ValueError:
            self.computed_total_bill = 0.0

        try:
            self.owner_family = int(self.in_owner.value())
        except Exception:
            self.owner_family = 1
        tenant_total = sum(int(t["family"]) for t in self.tenant_rows)
        extra_total = sum(int(ep["headcount"]) for ep in self.extra_persons)
        self.computed_total_persons = self.owner_family + tenant_total + extra_total

        if self.computed_total_persons > 0:
            self.computed_rate_per_person = self.computed_total_bill / self.computed_total_persons
        else:
            self.computed_rate_per_person = 0.0

        self.lbl_headcount_summary.setText(
            f"Total Campus Persons: {self.computed_total_persons}  ·  "
            f"Per-Head Water Share: {fmt_money(self.computed_rate_per_person)}")

    def save_record(self):
        uid = self.controller.current_session_id
        if uid is None: return
        if self.computed_total_persons <= 0:
            QMessageBox.warning(self, "Validation", "Total persons must be > 0.")
            return

        billing_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        reading_date = self.in_date.date().toString("yyyy-MM-dd")
        prev = float(self.in_prev.text() or 0)
        pres = float(self.in_pres.text() or 0)
        rate = float(self.in_rate.text() or 0)
        if pres < prev:
            QMessageBox.warning(self, "Validation", "Present reading cannot be less than previous.")
            return
        remarks = self.in_remarks.text().strip()

        breakdown = {
            "owner_family": self.owner_family,
            "tenant_families": {t["name"]: int(t["family"]) for t in self.tenant_rows},
            "extra_persons": self.extra_persons,
        }
        breakdown_json = json.dumps(breakdown)

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""INSERT INTO campus_water_records
            (user_id, billing_month, reading_date, prev_reading, pres_reading, rate_elec,
             total_bill, total_persons, rate_per_person, breakdown_json, remarks)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id, billing_month) DO UPDATE SET
                reading_date=excluded.reading_date,
                prev_reading=excluded.prev_reading,
                pres_reading=excluded.pres_reading,
                rate_elec=excluded.rate_elec,
                total_bill=excluded.total_bill,
                total_persons=excluded.total_persons,
                rate_per_person=excluded.rate_per_person,
                breakdown_json=excluded.breakdown_json,
                remarks=excluded.remarks""",
            (uid, billing_month, reading_date, prev, pres, rate,
             self.computed_total_bill, self.computed_total_persons,
             self.computed_rate_per_person, breakdown_json, remarks))
        conn.commit()
        conn.close()

        try:
            tenants_data = [{"name": t["name"], "family": int(t["family"]), "guests": 0}
                            for t in self.tenant_rows]
            ensure_all_water_sheets(
                uid, billing_month, reading_date,
                prev, pres, rate,
                self.owner_name, self.owner_family,
                tenants_data, remarks=remarks
            )
            if hasattr(self.controller, "toast"):
                self.controller.toast(f"Campus water record saved ({billing_month})", "check", ACCENT_GREEN)
            QMessageBox.information(self, "Saved",
                                    f"Campus water record for {billing_month} saved.\n\n"
                                    f"Per-head: {fmt_money(self.computed_rate_per_person)}\n"
                                    f"Total persons: {self.computed_total_persons}\n"
                                    f"Total pump bill: {fmt_money(self.computed_total_bill)}\n\n"
                                    f"All tenant sheets updated.")
        except Exception as e:
            log.error(f"Master water Excel write failed: {e}")
            QMessageBox.warning(self, "Saved (Excel warning)", f"DB saved but Excel write had an issue:\n\n{e}")

    def view_month_bills(self):
        uid = self.controller.current_session_id
        if uid is None: return
        billing_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        dlg = WaterBillsViewerDialog(self.controller, billing_month, parent=self)
        dlg.exec()

    def open_ledger(self):
        uid = self.controller.current_session_id
        if uid is None: return
        path = get_master_water_path(uid)
        if not os.path.exists(path):
            QMessageBox.information(self, "Not yet created", "Save a water record first to create the Excel.")
            return
        try:
            if sys.platform.startswith("win"): os.startfile(path)
            elif sys.platform == "darwin": subprocess.run(["open", path])
            else: subprocess.run(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, "Could not open", f"{e}\n\nPath:\n{path}")


# ==============================================================================
# § 12  BILLING TERMINAL
# ==============================================================================
class BillingEntrySubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.tenant_cache = {}
        self.elec_total = 0.0
        self.water_total = 0.0
        self.water_per_head = 0.0
        self.running_balance = 0.0
        self.notes = ""
        self._gen_in_progress = False
        self._current_water_record = None
        self._is_common_water = False
        self.init_ui()

    def _lbl(self, text, color=None):
        if color is None: color = T.c("text_faint")
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {color}; letter-spacing: 0.5px; background: transparent; font-family: {FONT_STACK};")
        return l

    def init_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 12, 18, 18)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        c = QWidget()
        c.setStyleSheet(card_qss() + input_qss())
        v = QVBoxLayout(c)
        v.setSpacing(14)
        v.setContentsMargins(0, 0, 6, 0)

        hdr = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("billing", color=ACCENT_INDIGO, size=24))
        tl = QLabel("Generate Bill Statement")
        tl.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        hdr.addWidget(ic)
        hdr.addSpacing(6)
        hdr.addWidget(tl)
        hdr.addStretch(1)
        v.addLayout(hdr)

        # 1. Tenant Selection Card
        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(8)
        cl.addWidget(self._lbl("SELECT TENANT *"))
        self.combo_tenant = QComboBox()
        self.combo_tenant.setFixedHeight(INPUT_H)
        self.combo_tenant.currentIndexChanged.connect(self.on_tenant_selected)
        cl.addWidget(self.combo_tenant)

        self.lbl_notice = QLabel("Select a tenant above")
        self.lbl_notice.setStyleSheet(f"""
            background-color: {T.c('input_bg')}; color: {T.c('text_dim')};
            border: 1px solid {T.c('input_border')}; padding: 10px 12px;
            border-radius: {RADIUS_INPUT}px; font-size: 12px; font-weight: 500;
            font-family: {FONT_STACK};
        """)
        self.lbl_notice.setWordWrap(True)
        cl.addWidget(self.lbl_notice)
        v.addWidget(card)

        # 2. Cycle Details & Base Rent
        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        title = QLabel("Billing Cycle & Lease Terms")
        title.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        cl.addWidget(title)

        row = QHBoxLayout()
        row.setSpacing(10)
        vm = QVBoxLayout(); vm.setSpacing(4); vm.addWidget(self._lbl("BILL MONTH *"))
        self.combo_month = QComboBox()
        setup_combobox(self.combo_month)
        self.combo_month.addItems(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
        self.combo_month.setCurrentIndex(QDate.currentDate().month() - 1)
        self.combo_month.setFixedHeight(INPUT_H)
        self.combo_month.currentIndexChanged.connect(self.on_month_year_changed)
        vm.addWidget(self.combo_month)

        vy = QVBoxLayout(); vy.setSpacing(4); vy.addWidget(self._lbl("BILL YEAR *"))
        self.combo_year = QComboBox()
        setup_combobox(self.combo_year)
        cy = QDate.currentDate().year()
        for y in range(cy - 2, cy + 5): self.combo_year.addItem(str(y))
        self.combo_year.setCurrentText(str(cy))
        self.combo_year.setFixedHeight(INPUT_H)
        self.combo_year.currentIndexChanged.connect(self.on_month_year_changed)
        vy.addWidget(self.combo_year)

        vd = QVBoxLayout(); vd.setSpacing(4); vd.addWidget(self._lbl("BILL DATE *"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setFixedHeight(INPUT_H)
        vd.addWidget(self.in_date)
        row.addLayout(vm, 1)
        row.addLayout(vy, 1)
        row.addLayout(vd, 1)
        cl.addLayout(row)

        cl.addWidget(self._lbl("MONTHLY BASE RENT (₹) *"))
        self.in_rent = QLineEdit()
        self.in_rent.setPlaceholderText("auto-filled")
        self.in_rent.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_rent.setFixedHeight(INPUT_H)
        self.in_rent.textChanged.connect(self.recalc)
        cl.addWidget(self.in_rent)

        sr = QHBoxLayout()
        self.lbl_prev = QLabel("<b>Previous Balance: ₹0.00</b>")
        self.lbl_prev.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        self.lbl_credit = QLabel("<b>Advance Credit: ₹0.00</b>")
        self.lbl_credit.setStyleSheet(f"color: {ACCENT_GREEN}; font-size: 12px; background: transparent; font-family: {FONT_STACK};")
        sr.addWidget(self.lbl_prev)
        sr.addStretch(1)
        sr.addWidget(self.lbl_credit)
        cl.addLayout(sr)
        v.addWidget(card)

        # 3. Electricity Meter
        ec = GlassCard()
        el = QVBoxLayout(ec)
        el.setContentsMargins(20, 16, 20, 16)
        el.setSpacing(10)
        et = QLabel("Electricity Meter")
        et.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {ACCENT_AMBER}; font-family: {FONT_STACK};")
        el.addWidget(et)

        r1 = QHBoxLayout()
        r1.setSpacing(10)
        v1 = QVBoxLayout(); v1.setSpacing(4); v1.addWidget(self._lbl("PREVIOUS READING"))
        self.in_eprev = QLineEdit("0"); self.in_eprev.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_eprev.setFixedHeight(INPUT_H); v1.addWidget(self.in_eprev)
        v2 = QVBoxLayout(); v2.setSpacing(4); v2.addWidget(self._lbl("CURRENT READING"))
        self.in_epres = QLineEdit("0"); self.in_epres.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_epres.setFixedHeight(INPUT_H); v2.addWidget(self.in_epres)
        r1.addLayout(v1); r1.addLayout(v2)
        el.addLayout(r1)

        r2 = QHBoxLayout()
        r2.setSpacing(10)
        v3 = QVBoxLayout(); v3.setSpacing(4); v3.addWidget(self._lbl("RATE PER UNIT (₹)"))
        self.in_erate = QLineEdit("8.50"); self.in_erate.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
        self.in_erate.setFixedHeight(INPUT_H); v3.addWidget(self.in_erate)
        r2.addLayout(v3); r2.addStretch(1)
        el.addLayout(r2)

        be = QPushButton("Calculate Electricity")
        be.setIcon(SVGIconManager.icon("bolt", color="#04150E", size=16))
        be.setCursor(Qt.CursorShape.PointingHandCursor)
        be.setFixedHeight(BTN_PRIMARY)
        be.setStyleSheet(primary_button_qss(ACCENT_AMBER, height=BTN_PRIMARY))
        be.clicked.connect(self.calc_elec)
        el.addWidget(be)

        self.lbl_eres = QLabel("<b>Electricity: ₹0.00</b>")
        self.lbl_eres.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 13px; font-weight: 700; font-family: {FONT_STACK};")
        el.addWidget(self.lbl_eres)
        v.addWidget(ec)

        # 4. Water Meter Card
        wc = GlassCard()
        wl = QVBoxLayout(wc)
        wl.setContentsMargins(20, 16, 20, 16)
        wl.setSpacing(10)
        self.w_title = QLabel("Water (Campus Shared Pump)")
        self.w_title.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {ACCENT_INDIGO}; font-family: {FONT_STACK};")
        wl.addWidget(self.w_title)

        self.warn_banner = QLabel("")
        self.warn_banner.setStyleSheet(f"""
            background-color: rgba(245,158,11,0.12); color: {ACCENT_AMBER};
            border: 1px solid rgba(245,158,11,0.45); padding: 10px 14px;
            border-radius: {RADIUS_INPUT}px; font-size: 12px; font-weight: 600;
            font-family: {FONT_STACK};
        """)
        self.warn_banner.setWordWrap(True)
        self.warn_banner.hide()
        wl.addWidget(self.warn_banner)

        self.btn_go_water = QPushButton("GO TO WATER RECORD")
        self.btn_go_water.setIcon(SVGIconManager.icon("water", color=ACCENT_INDIGO, size=16))
        self.btn_go_water.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_go_water.setFixedHeight(BTN_PRIMARY)
        self.btn_go_water.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        self.btn_go_water.clicked.connect(self.goto_water_record)
        self.btn_go_water.hide()
        wl.addWidget(self.btn_go_water)

        self.common_widget = QWidget()
        self.common_widget.setStyleSheet("background: transparent;")
        cw = QVBoxLayout(self.common_widget)
        cw.setContentsMargins(0, 0, 0, 0)
        cw.setSpacing(8)

        self.lbl_shared_badge = QLabel("")
        self.lbl_shared_badge.setStyleSheet(f"""
            background-color: rgba(16,185,129,0.12); color: {ACCENT_INDIGO};
            border: 1px solid rgba(16,185,129,0.35); padding: 6px 10px;
            border-radius: 8px; font-size: 11px; font-weight: 700;
            font-family: {FONT_STACK};
        """)
        self.lbl_shared_badge.setWordWrap(True)
        cw.addWidget(self.lbl_shared_badge)

        rr = QHBoxLayout()
        rr.setSpacing(10)
        vph = QVBoxLayout(); vph.setSpacing(4); vph.addWidget(self._lbl("PER-HEAD COST"))
        self.lbl_per_head_val = QLabel("—")
        self.lbl_per_head_val.setStyleSheet(f"""
            background-color: {T.c('input_bg')};
            border: 1px solid {T.c('input_border')};
            border-radius: {RADIUS_INPUT}px;
            padding: 9px 13px; color: {ACCENT_GREEN};
            font-size: 13px; font-weight: 700;
            font-family: {FONT_STACK};
        """)
        self.lbl_per_head_val.setFixedHeight(INPUT_H)
        vph.addWidget(self.lbl_per_head_val)

        vfl = QVBoxLayout(); vfl.setSpacing(4); vfl.addWidget(self._lbl("FLAT FAMILY OCCUPANTS"))
        self.lbl_flat_val = QLabel("—")
        self.lbl_flat_val.setStyleSheet(f"""
            background-color: {T.c('input_bg')};
            border: 1px solid {T.c('input_border')};
            border-radius: {RADIUS_INPUT}px;
            padding: 9px 13px; color: {T.c('text_primary')};
            font-size: 13px; font-weight: 700;
            font-family: {FONT_STACK};
        """)
        self.lbl_flat_val.setFixedHeight(INPUT_H)
        vfl.addWidget(self.lbl_flat_val)
        rr.addLayout(vph)
        rr.addLayout(vfl)
        cw.addLayout(rr)

        vg = QVBoxLayout()
        vg.setSpacing(4)
        vg.addWidget(self._lbl("GUEST(S) THIS MONTH (EDITABLE)"))
        self.in_guests = QLineEdit("0")
        self.in_guests.setValidator(QDoubleValidator(0, 50, 0, self))
        self.in_guests.setStyleSheet(input_qss())
        self.in_guests.setFixedHeight(INPUT_H)
        self.in_guests.textChanged.connect(self.recalc_common_water)
        vg.addWidget(self.in_guests)
        cw.addLayout(vg)

        self.lbl_common_water_result = QLabel("Water share: ₹0.00")
        self.lbl_common_water_result.setStyleSheet(f"color: {ACCENT_INDIGO}; font-size: 13px; font-weight: 700; font-family: {FONT_STACK};")
        cw.addWidget(self.lbl_common_water_result)
        wl.addWidget(self.common_widget)

        self.individual_widget = QWidget()
        self.individual_widget.setStyleSheet("background: transparent;")
        iw = QVBoxLayout(self.individual_widget)
        iw.setContentsMargins(0, 0, 0, 0)
        iw.setSpacing(8)

        rr2 = QHBoxLayout()
        rr2.setSpacing(10)
        vw = QVBoxLayout(); vw.setSpacing(4); vw.addWidget(self._lbl("PREVIOUS READING"))
        self.in_wprev = QLineEdit("0"); self.in_wprev.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_wprev.setFixedHeight(INPUT_H); vw.addWidget(self.in_wprev)
        vc = QVBoxLayout(); vc.setSpacing(4); vc.addWidget(self._lbl("CURRENT READING"))
        self.in_wpres = QLineEdit("0"); self.in_wpres.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_wpres.setFixedHeight(INPUT_H); vc.addWidget(self.in_wpres)
        rr2.addLayout(vw)
        rr2.addLayout(vc)
        iw.addLayout(rr2)

        vr = QVBoxLayout(); vr.setSpacing(4); vr.addWidget(self._lbl("RATE PER UNIT (₹)"))
        self.in_wrate = QLineEdit("8.50"); self.in_wrate.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
        self.in_wrate.setFixedHeight(INPUT_H); vr.addWidget(self.in_wrate)
        iw.addLayout(vr)

        bw = QPushButton("Calculate Water")
        bw.setIcon(SVGIconManager.icon("water", color="#04150E", size=16))
        bw.setCursor(Qt.CursorShape.PointingHandCursor)
        bw.setFixedHeight(BTN_PRIMARY)
        bw.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        bw.clicked.connect(self.calc_water_individual)
        iw.addWidget(bw)

        self.lbl_wres = QLabel("<b>Water: ₹0.00</b>")
        self.lbl_wres.setStyleSheet(f"color: {ACCENT_INDIGO}; font-size: 13px; font-weight: 700; font-family: {FONT_STACK};")
        iw.addWidget(self.lbl_wres)
        self.individual_widget.hide()
        wl.addWidget(self.individual_widget)
        v.addWidget(wc)

        # 5. Concessions & Discounts
        cc = GlassCard()
        cl = QVBoxLayout(cc)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        atitle = QLabel("Adjustments & Concessions")
        atitle.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        cl.addWidget(atitle)

        self.chk_concession = ToggleCheckBox("Apply Concession / Discount")
        self.chk_concession.toggled.connect(self.toggle_concession)
        cl.addWidget(self.chk_concession)

        self.row_cf = QHBoxLayout()
        self.row_cf.setSpacing(10)
        def mk_cf(label):
            vv = QVBoxLayout(); vv.setSpacing(4); vv.addWidget(self._lbl(label))
            e = QLineEdit("0.00"); e.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
            e.setStyleSheet(input_qss())
            e.setFixedHeight(INPUT_H); e.textChanged.connect(self.recalc); vv.addWidget(e)
            return vv, e
        c1, self.in_dr = mk_cf("RENT CONCESSION (₹)")
        c2, self.in_de = mk_cf("ELEC. DISC. (₹)")
        c3, self.in_dw = mk_cf("WATER DISC. (₹)")
        c4, self.in_dm = mk_cf("OTHER MISC (₹)")
        self.row_cf.addLayout(c1)
        self.row_cf.addLayout(c2)
        self.row_cf.addLayout(c3)
        self.row_cf.addLayout(c4)
        cl.addLayout(self.row_cf)

        cl.addWidget(self._lbl("GENERAL DEDUCTION (₹)"))
        self.in_general = QLineEdit("0.00")
        self.in_general.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_general.setStyleSheet(input_qss())
        self.in_general.setFixedHeight(INPUT_H)
        self.in_general.textChanged.connect(self.recalc)
        cl.addWidget(self.in_general)

        cl.addWidget(self._lbl("DEDUCTION REASON"))
        self.in_reason = QLineEdit()
        self.in_reason.setPlaceholderText("Reason for discount...")
        self.in_reason.setStyleSheet(input_qss())
        self.in_reason.setFixedHeight(INPUT_H)
        cl.addWidget(self.in_reason)

        self.chk_render = ToggleCheckBox("Render discount on customer PDF copy")
        self.chk_render.setChecked(True)
        cl.addWidget(self.chk_render)
        v.addWidget(cc)

        # 6. Owner-Only Audit Notes
        nc = GlassCard()
        nl = QVBoxLayout(nc)
        nl.setContentsMargins(20, 16, 20, 16)
        nl.setSpacing(10)
        nt = QLabel("Owner Internal Notes")
        nt.setStyleSheet(f"font-size: {FS_CARD_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        nl.addWidget(nt)

        hint = QLabel("Owner-only audit log. These notes appear exclusively on the OWNER PDF and Excel ledger — "
                      "never on the tenant's customer invoice.")
        hint.setStyleSheet(f"font-size: 11px; color: {T.c('text_secondary')}; font-family: {FONT_STACK};")
        hint.setWordWrap(True)
        nl.addWidget(hint)

        nl.addWidget(self._lbl("RENT & ELECTRICITY NOTE"))
        self.in_common_note = QTextEdit()
        self.in_common_note.setPlaceholderText("e.g. Rate renegotiated from Nov; advance carried forward…")
        self.in_common_note.setStyleSheet(INPUT_STYLE)
        self.in_common_note.setMaximumHeight(64)
        nl.addWidget(self.in_common_note)

        nl.addWidget(self._lbl("WATER NOTE"))
        self.in_water_note = QTextEdit()
        self.in_water_note.setPlaceholderText("e.g. 2 guests stayed for 10 days; charged as 6 heads for Oct…")
        self.in_water_note.setStyleSheet(INPUT_STYLE)
        self.in_water_note.setMaximumHeight(64)
        nl.addWidget(self.in_water_note)
        v.addWidget(nc)

        # 7. Grand Total & Action Bar
        bc = GlassCard()
        bl = QHBoxLayout(bc)
        bl.setContentsMargins(20, 16, 20, 16)
        vt = QVBoxLayout()
        vt.setSpacing(6)
        self.lbl_grand = QLabel("TOTAL BILL: ₹0.00")
        self.lbl_grand.setStyleSheet(f"font-size: 22px; font-weight: 900; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        rec = QPushButton("Recalculate")
        rec.setIcon(SVGIconManager.icon("refresh", color="#F8FAFC", size=14))
        rec.setCursor(Qt.CursorShape.PointingHandCursor)
        rec.setFixedHeight(BTN_INLINE)
        rec.setStyleSheet(pill_button_qss("#ffffff", height=BTN_INLINE))
        rec.clicked.connect(self.recalc)
        vt.addWidget(self.lbl_grand)
        vt.addWidget(rec)
        bl.addLayout(vt)
        bl.addStretch(1)

        self.btn_save = QPushButton("Save & Record Bill")
        self.btn_save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        self.btn_save.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_save.setFixedHeight(BTN_PRIMARY)
        self.btn_save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        self.btn_save.clicked.connect(self.save_bill)
        bl.addWidget(self.btn_save)
        v.addWidget(bc)

        scroll.setWidget(c)
        root.addWidget(scroll)
        self.toggle_concession()

    def toggle_concession(self):
        if not hasattr(self, "lbl_grand"): return
        on = self.chk_concession.isChecked()
        for w in [self.in_dr, self.in_de, self.in_dw, self.in_dm, self.in_reason, self.chk_render]:
            w.setEnabled(on)
        self.recalc()

    def open_notes(self):
        text, ok = QInputDialog.getMultiLineText(self, "Remarks", "Enter remarks:")
        if ok and text: self.notes = text

    def refresh_page_data(self):
        uid = self.controller.current_session_id
        if uid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.id, t.name, t.tenant_uid, t.monthly_rent, t.occupants_count,
                       p.name, p.address, p.water_meter_type, p.id, p.owner_members
                       FROM tenants t JOIN properties p ON t.property_id = p.id
                       WHERE t.user_id=? AND t.is_active=1 ORDER BY t.id DESC""", (uid,))
        rows = cur.fetchall()
        conn.close()
        self.tenant_cache = {}
        self.combo_tenant.blockSignals(True)
        self.combo_tenant.clear()
        self.combo_tenant.addItem("-- Select Active Tenant --", None)
        for tid, tn, tuid, trent, tocc, pn, pa, wm, pid, ohc in rows:
            self.combo_tenant.addItem(f"{tn} ({pn} - {tuid})", tid)
            self.tenant_cache[tid] = {
                "name": tn, "uid": tuid, "rent": trent, "occupants": tocc,
                "prop_name": pn, "address": pa, "water_meter_type": wm or "Common",
                "property_id": pid, "owner_members": ohc or 1,
            }
        self.combo_tenant.blockSignals(False)

    def on_month_year_changed(self):
        if self.combo_tenant.currentData() is not None:
            self.refresh_water_source()

    def on_tenant_selected(self):
        tid = self.combo_tenant.currentData()
        if not tid or tid not in self.tenant_cache:
            self.lbl_notice.setText("Select a tenant above")
            self.in_rent.clear()
            self._current_water_record = None
            self._is_common_water = False
            self.common_widget.hide(); self.individual_widget.hide()
            self.warn_banner.hide(); self.btn_go_water.hide()
            return
        info = self.tenant_cache[tid]
        self.lbl_notice.setText(f"Active lease: <b>{info['name']}</b> at <b>{info['prop_name']}</b> ({info['address']})")
        self.lbl_notice.setStyleSheet(f"""
            background-color: rgba(16,185,129,0.12); color: {T.c('text_primary')};
            border: 1px solid rgba(16,185,129,0.30); padding: 10px 12px;
            border-radius: {RADIUS_INPUT}px; font-size: 12px; font-weight: 500;
            font-family: {FONT_STACK};
        """)
        self.in_rent.setText(f"{info['rent']:.2f}")

        is_common = info["water_meter_type"] != "Individual"
        self._is_common_water = is_common

        if is_common:
            self.common_widget.show()
            self.individual_widget.hide()
            self.w_title.setText("Water (Campus Shared Pump)")
            self.refresh_water_source()
        else:
            self.common_widget.hide()
            self.individual_widget.show()
            self.w_title.setText("Water (Individual Meter)")
            self.warn_banner.hide(); self.btn_go_water.hide()
            self._current_water_record = None
            self.water_total = 0.0
            self.water_per_head = 0.0

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT elec_pres, water_pres FROM tenant_bills WHERE tenant_id=? ORDER BY id DESC LIMIT 1", (tid,))
        last = cur.fetchone()
        conn.close()
        if last:
            self.in_eprev.setText(str(last[0]))
            if not is_common and last[1] is not None:
                self.in_wprev.setText(str(last[1]))
        else:
            self.in_eprev.setText("0")
            if not is_common: self.in_wprev.setText("0")

        self.running_balance = get_tenant_balance(tid)
        if self.running_balance > 0:
            self.lbl_prev.setText(f"<b>Previous Balance: {fmt_money(self.running_balance)}</b>")
            self.lbl_credit.setText("<b>Advance Credit: ₹0.00</b>")
        elif self.running_balance < 0:
            self.lbl_prev.setText("<b>Previous Balance: ₹0.00</b>")
            self.lbl_credit.setText(f"<b>Advance Credit: {fmt_money(-self.running_balance)}</b>")
        else:
            self.lbl_prev.setText("<b>Previous Balance: ₹0.00</b>")
            self.lbl_credit.setText("<b>Advance Credit: ₹0.00</b>")

        self.calc_elec()
        if not is_common:
            self.calc_water_individual()
        else:
            self.recalc_common_water()

    def refresh_water_source(self):
        if not self._is_common_water: return
        uid = self.controller.current_session_id
        billing_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        rec = fetch_campus_water_record(uid, billing_month)
        self._current_water_record = rec
        tid = self.combo_tenant.currentData()
        info = self.tenant_cache.get(tid, {})
        if rec:
            self.lbl_per_head_val.setText(_pdf_money(rec["rate_per_person"]))
            self.lbl_flat_val.setText(f"{info.get('occupants', 0)} persons")
            self.lbl_shared_badge.setText(
                f"Locked for {billing_month}  ·  "
                f"Campus bill {_pdf_money(rec['total_bill'])} ÷ "
                f"{rec['total_persons']} persons")
            self.warn_banner.hide()
            self.btn_go_water.hide()
            self.recalc_common_water()
        else:
            self.lbl_per_head_val.setText("—")
            self.lbl_flat_val.setText(f"{info.get('occupants', 0)} persons")
            self.lbl_shared_badge.setText("")
            self.warn_banner.setText(
                f"No campus water record exists for {billing_month}. "
                f"Record the pump reading once in the Water Record tab, then come back."
            )
            self.warn_banner.show()
            self.btn_go_water.show()
            self.water_total = 0.0; self.water_per_head = 0.0
            self.lbl_common_water_result.setText("Water share: ₹0.00")
            self.recalc()

    def recalc_common_water(self):
        if not self._is_common_water: return
        tid = self.combo_tenant.currentData()
        if not tid or tid not in self.tenant_cache:
            self.water_total = 0.0; self.water_per_head = 0.0
            self.lbl_common_water_result.setText("Water share: ₹0.00")
            self.recalc(); return
        info = self.tenant_cache[tid]
        if not self._current_water_record:
            self.water_total = 0.0; self.water_per_head = 0.0
            self.lbl_common_water_result.setText("Water share: ₹0.00")
            self.recalc(); return
        try:
            guests = float(self.in_guests.text() or 0)
        except ValueError:
            guests = 0.0
        per_head = float(self._current_water_record["rate_per_person"])
        self.water_per_head = per_head
        self.water_total = round(per_head * (info["occupants"] + guests), 2)
        self.lbl_common_water_result.setText(
            f"Water share: {fmt_money(self.water_total)}  "
            f"(= ({info['occupants']} + {int(guests)}) × {fmt_money(per_head)})"
        )
        self.recalc()

    def goto_water_record(self):
        self.controller.dashboard_page.route_water_record()

    def calc_water_individual(self):
        if self._is_common_water: return
        try:
            prev = float(self.in_wprev.text() or 0)
            pres = float(self.in_wpres.text() or 0)
            rate = float(self.in_wrate.text() or 0)
            units = max(0.0, pres - prev)
            self.water_total = round(units * rate, 2)
            self.water_per_head = 0.0
            self.lbl_wres.setText(f"Water: {fmt_money(self.water_total)}")
        except ValueError:
            pass
        self.recalc()

    def calc_elec(self):
        try:
            prev = float(self.in_eprev.text() or 0)
            pres = float(self.in_epres.text() or 0)
            rate = float(self.in_erate.text() or 0)
            units = max(0.0, pres - prev)
            self.elec_total = round(units * rate, 2)
            self.lbl_eres.setText(f"Electricity: {fmt_money(self.elec_total)}")
        except ValueError:
            pass
        self.recalc()

    def recalc(self):
        if not hasattr(self, "lbl_grand"): return 0.0, 0.0
        try:
            rent = float(self.in_rent.text() or 0)
            discounts = 0.0
            if self.chk_concession.isChecked():
                discounts += float(self.in_dr.text() or 0)
                discounts += float(self.in_de.text() or 0)
                discounts += float(self.in_dw.text() or 0)
                discounts += float(self.in_dm.text() or 0)
            discounts += float(self.in_general.text() or 0)
            total = round(rent + self.elec_total + self.water_total - discounts + self.running_balance, 2)
            self.lbl_grand.setText(f"TOTAL BILL: {fmt_money(total)}")
            return total, discounts
        except ValueError:
            return 0.0, 0.0

    def save_bill(self):
        if self._gen_in_progress:
            QMessageBox.information(self, "Please wait", "Generation in progress...")
            return
        tid = self.combo_tenant.currentData()
        if not tid or tid not in self.tenant_cache:
            QMessageBox.warning(self, "Selection", "Select a tenant first.")
            return
        info = self.tenant_cache[tid]

        is_common = info["water_meter_type"] != "Individual"
        if is_common and not self._current_water_record:
            QMessageBox.warning(self, "No Water Record",
                                f"Record the campus water reading for "
                                f"{self.combo_month.currentText()} {self.combo_year.currentText()} first.")
            return

        b_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        b_date = self.in_date.date().toString("yyyy-MM-dd")
        year_str = self.combo_year.currentText()
        month_folder = f"{self.in_date.date().month():02d}_{sanitize_fs_name(b_month)}"

        grand_total, discounts = self.recalc()
        base_rent = float(self.in_rent.text() or 0)
        guests = int(float(self.in_guests.text() or 0)) if is_common else 0

        eprev = float(self.in_eprev.text() or 0); epres = float(self.in_epres.text() or 0)
        erate = float(self.in_erate.text() or 0)
        eunits = max(0.0, epres - eprev)

        if is_common:
            wprev = float(self._current_water_record["prev_reading"])
            wpres = float(self._current_water_record["pres_reading"])
            wrate = float(self._current_water_record["rate_per_person"])
            wunits = max(0.0, wpres - wprev)
            total_persons_snapshot = int(self._current_water_record["total_persons"])
        else:
            wprev = float(self.in_wprev.text() or 0); wpres = float(self.in_wpres.text() or 0)
            wrate = float(self.in_wrate.text() or 0)
            wunits = max(0.0, wpres - wprev)
            total_persons_snapshot = 0

        common_note_val = self.in_common_note.toPlainText().strip()
        water_note_val = self.in_water_note.toPlainText().strip()
        render_customer = self.chk_render.isChecked()

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""INSERT INTO tenant_bills
            (user_id, tenant_id, billing_month, bill_date, base_rent,
             elec_prev, elec_pres, elec_units, elec_rate, elec_total,
             water_prev, water_pres, water_units, water_rate, water_total,
             discount_total, prev_due, total_payable, amount_paid, balance_remaining,
             payment_status, settled_status, remarks, guests_count,
             owner_members_snapshot, campus_persons_snapshot,
             common_note, water_note)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, 'Unpaid', 'unpaid', ?, ?, ?, ?, ?, ?)""",
            (self.controller.current_session_id, tid, b_month, b_date, base_rent,
             eprev, epres, eunits, erate, self.elec_total,
             wprev, wpres, wunits, wrate, self.water_total,
             discounts, self.running_balance, grand_total, grand_total,
             self.in_reason.text().strip(), guests,
             0, total_persons_snapshot,
             common_note_val, water_note_val))
        bill_id = cur.lastrowid

        cur.execute("""INSERT INTO tenant_transactions
            (user_id, tenant_id, property_id, txn_date, txn_type, description, amount, linked_bill_id)
            VALUES (?, ?, ?, ?, 'bill', ?, ?, ?)""",
            (self.controller.current_session_id, tid, info["property_id"],
             b_date, f"Bill for {b_month}", grand_total, bill_id))
        conn.commit()
        conn.close()

        if is_common:
            try:
                update_single_water_sheet_guests(
                    self.controller.current_session_id, b_month, "tenant", info["name"], guests,
                    bill_note=water_note_val
                )
            except Exception as e:
                log.error(f"Water sheet guest update failed: {e}")

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT owner_name, phone FROM profiles WHERE user_id=?",
                    (self.controller.current_session_id,))
        p_row = cur.fetchone()
        conn.close()
        owner_name_full = p_row[0] if (p_row and p_row[0]) else "Property Administrator"
        owner_phone = p_row[1] if (p_row and p_row[1]) else "--"

        tenant_dir = get_tenant_dir(self.controller.current_session_id, info["address"],
                                    info["prop_name"], info["name"], info["uid"])
        ledger_path = os.path.join(tenant_dir, "Tenant_Master_Ledger.xlsx")
        gen_bill_dir = get_bill_dir(self.controller.current_session_id, info["address"],
                                    info["prop_name"], info["name"], info["uid"],
                                    year_str, b_month, "Generated")
        unpaid_bill_dir = get_bill_dir(self.controller.current_session_id, info["address"],
                                       info["prop_name"], info["name"], info["uid"],
                                       year_str, b_month, "Unpaid")

        payload = {
            "bill_id": bill_id, "owner_name": owner_name_full, "owner_phone": owner_phone,
            "property_name": info["prop_name"], "property_address": info["address"],
            "tenant_name": info["name"], "tenant_uid": info["uid"], "occupants_count": info["occupants"],
            "billing_month": b_month, "bill_date": b_date, "base_rent": base_rent,
            "elec_prev": eprev, "elec_pres": epres, "elec_rate": erate, "elec_total": self.elec_total,
            "water_prev": wprev, "water_pres": wpres, "water_rate": wrate, "water_total": self.water_total,
            "is_common_water": is_common,
            "discount_total": discounts, "discount_reason": self.in_reason.text().strip(),
            "render_discount_customer": render_customer,
            "prev_due": self.running_balance, "total_payable": grand_total, "amount_paid": 0.0,
            "balance_remaining": grand_total, "remarks": self.in_reason.text().strip(),
            "guests_count": guests, "total_persons": total_persons_snapshot,
            "owner_members_snapshot": 0,
            "common_note": common_note_val, "water_note": water_note_val,
        }

        self._gen_in_progress = True
        self.btn_save.setEnabled(False)
        self.btn_save.setText("Generating Documents...")

        def task_ledger():
            rebuild_tenant_ledger(self.controller.current_session_id, tid)
            return ("ledger", ledger_path)

        def task_bill_pdfs():
            generate_bill_pdfs(gen_bill_dir, payload, is_paid=False)
            cust, own_ = generate_bill_pdfs(unpaid_bill_dir, payload, is_paid=False)
            return ("pdf_bill", (cust, own_))

        self.worker = FileGenerator([task_ledger, task_bill_pdfs])
        self.worker.finished_ok.connect(lambda results: self._on_gen_done(results, unpaid_bill_dir, info))
        self.worker.failed.connect(self._on_gen_failed)
        self.worker.start()

    def _restore_save_button(self):
        self._gen_in_progress = False
        self.btn_save.setEnabled(True)
        self.btn_save.setText("Save & Record Bill")

    def _on_gen_done(self, results, gen_dir, info):
        self._restore_save_button()
        files_summary = []
        for kind, path in results:
            if path:
                if isinstance(path, tuple):
                    for p in path: files_summary.append(f"• {p}")
                else:
                    files_summary.append(f"• {path}")
        QMessageBox.information(self, "Bill Recorded",
                                f"Bill saved for {info['name']}.\n\nGenerated Files:\n" + "\n".join(files_summary))
        self.refresh_page_data()
        try:
            self.controller.dashboard_page.refresh_live_metrics()
        except Exception:
            pass

    def _on_gen_failed(self, err):
        self._restore_save_button()
        QMessageBox.critical(self, "Generation Failed",
                             f"Bill was saved, but file generation encountered an issue:\n\n{err[:800]}")


# ==============================================================================
# § 13  RECORDS & FINANCIAL TRANSACTIONS
# ==============================================================================
def fetch_tenant_ledger(tenant_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT id, txn_date, txn_type, description, amount, payment_method,
                          reference_no, linked_bill_id, affects_deposit
                   FROM tenant_transactions WHERE tenant_id=?
                   ORDER BY txn_date ASC, id ASC""", (tenant_id,))
    rows = cur.fetchall()
    conn.close()
    return rows


def fetch_unpaid_bills(tenant_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT id, billing_month, bill_date, total_payable, amount_paid, balance_remaining
                   FROM tenant_bills
                   WHERE tenant_id=? AND balance_remaining > 0.001
                     AND payment_status != 'Paid'
                     AND settled_status NOT IN ('closed', 'closed_manual', 'settled', 'settled_deposit')
                   ORDER BY bill_date DESC""", (tenant_id,))
    rows = cur.fetchall()
    conn.close()
    return rows


def fetch_bill_payments(bill_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT txn_date, payment_method, reference_no, amount
                   FROM tenant_transactions
                   WHERE linked_bill_id=? AND txn_type IN ('payment','advance')
                   ORDER BY txn_date ASC""", (bill_id,))
    rows = cur.fetchall()
    conn.close()
    return [{"date": r[0], "method": r[1] or "—", "ref": r[2] or "—", "amount": r[3] or 0.0} for r in rows]


def update_bill_settled_status(bill_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("SELECT total_payable, amount_paid, balance_remaining, settled_status FROM tenant_bills WHERE id=?", (bill_id,))
    row = cur.fetchone()
    if not row: conn.close(); return
    total, paid, rem, current = row
    if current in ("closed_manual", "closed"): conn.close(); return
    if (rem is not None and rem <= 0.001) or (paid is not None and total is not None and paid >= total - 0.001 and total > 0):
        status = "closed"
        rem = 0.0
        pay_st = "Paid"
    elif paid and paid > 0.001:
        status = "partial"
        pay_st = "Partial"
    else:
        status = "unpaid"
        pay_st = "Unpaid"
    cur.execute("UPDATE tenant_bills SET settled_status=?, balance_remaining=?, payment_status=? WHERE id=?", (status, rem, pay_st, bill_id))
    conn.commit()
    conn.close()


def apply_payment_to_unpaid_bills(tenant_id, amount):
    applied = []
    remaining = float(amount)
    if remaining <= 0:
        return applied
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""SELECT id, total_payable, amount_paid, balance_remaining
                   FROM tenant_bills
                   WHERE tenant_id=? AND balance_remaining > 0.001
                     AND payment_status != 'Paid'
                     AND settled_status NOT IN ('closed', 'closed_manual', 'settled', 'settled_deposit')
                   ORDER BY bill_date ASC, id ASC""", (tenant_id,))
    bills = cur.fetchall()
    for bid, total, paid, bal in bills:
        if remaining <= 0.001:
            break
        need = float(bal)
        take = min(remaining, need)
        new_paid = float(paid or 0.0) + take
        new_bal = max(0.0, need - take)
        if new_bal <= 0.001:
            new_bal = 0.0
            new_status = "Paid"
            new_settled = "closed"
        else:
            new_status = "Partial" if new_paid > 0.001 else "Unpaid"
            new_settled = "partial"
        cur.execute("""UPDATE tenant_bills
                       SET amount_paid=?, balance_remaining=?, payment_status=?, settled_status=?
                       WHERE id=?""",
                    (new_paid, new_bal, new_status, new_settled, bid))
        applied.append((bid, take))
        remaining -= take
    conn.commit()
    conn.close()
    return applied


def regenerate_paid_bill_pdfs(bill_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""
        SELECT b.user_id, b.billing_month, b.bill_date, b.base_rent,
               b.elec_prev, b.elec_pres, b.elec_rate, b.elec_total,
               b.water_prev, b.water_pres, b.water_rate, b.water_total,
               b.discount_total, b.prev_due, b.total_payable, b.amount_paid,
               b.balance_remaining, b.guests_count, b.owner_members_snapshot, b.remarks,
               t.name, t.tenant_uid, t.permanent_address, t.occupants_count,
               p.name, p.address, p.water_meter_type, pr.owner_name, pr.phone,
               b.common_note, b.water_note
        FROM tenant_bills b
        JOIN tenants t ON b.tenant_id = t.id
        JOIN properties p ON t.property_id = p.id
        LEFT JOIN profiles pr ON pr.user_id = b.user_id
        WHERE b.id=?
    """, (bill_id,))
    row = cur.fetchone()
    if not row: conn.close(); return
    (uid, b_month, b_date, base_rent,
     eprev, epres, erate, e_total,
     wprev, wpres, wrate, w_total,
     discount, prev_due, total_payable, amount_paid,
     balance_remaining, guests, owner_snapshot, remarks,
     t_name, t_uid, t_addr, t_occ,
     p_name, p_addr, w_meter, owner_name, owner_phone,
     common_note, water_note) = row
    conn.close()

    payments = fetch_bill_payments(bill_id)
    method = payments[-1]["method"] if payments else None
    pay_date = payments[-1]["date"] if payments else None

    year_str = b_date.split("-")[0]
    unpaid_dir = get_bill_dir(uid, p_addr, p_name, t_name, t_uid, year_str, b_month, "Unpaid")
    paid_dir = get_bill_dir(uid, p_addr, p_name, t_name, t_uid, year_str, b_month, "Paid")

    payload = {
        "bill_id": bill_id, "owner_name": owner_name or "Property Administrator",
        "owner_phone": owner_phone or "—",
        "property_name": p_name, "property_address": p_addr,
        "tenant_name": t_name, "tenant_uid": t_uid, "occupants_count": t_occ,
        "billing_month": b_month, "bill_date": b_date, "base_rent": base_rent,
        "elec_prev": eprev, "elec_pres": epres, "elec_rate": erate, "elec_total": e_total,
        "water_prev": wprev, "water_pres": wpres, "water_rate": wrate, "water_total": w_total,
        "is_common_water": (w_meter or "Common") != "Individual",
        "discount_total": discount or 0, "discount_reason": "",
        "render_discount_customer": True,
        "prev_due": prev_due or 0, "total_payable": total_payable, "amount_paid": amount_paid,
        "balance_remaining": balance_remaining, "remarks": remarks or "",
        "guests_count": guests or 0,
        "owner_members_snapshot": owner_snapshot or 1,
        "total_persons": (owner_snapshot or 1) + (t_occ or 1) + (guests or 0),
        "payment_method": method,
        "payment_date": pay_date,
        "payments": payments,
        "common_note": common_note or "",
        "water_note": water_note or "",
    }

    # If remaining balance <= 0, clear Unpaid Bills and generate Paid Bill in Paid Receipts
    if (balance_remaining or 0) <= 0.001:
        if os.path.isdir(unpaid_dir):
            for fname in os.listdir(unpaid_dir):
                if fname.endswith(".pdf"):
                    try: os.remove(os.path.join(unpaid_dir, fname))
                    except Exception: pass
            try:
                if not os.listdir(unpaid_dir):
                    os.rmdir(unpaid_dir)
            except Exception: pass
        ensure_dir(paid_dir)
        try:
            generate_bill_pdfs(paid_dir, payload, is_paid=True)
        except Exception as e:
            log.error(f"Paid bill PDF generation failed: {e}")
    else:
        # If partial payment made and there is remaining due, update Unpaid Bills with remaining balance
        ensure_dir(unpaid_dir)
        try:
            generate_bill_pdfs(unpaid_dir, payload, is_paid=False)
        except Exception as e:
            log.error(f"Unpaid remaining bill PDF failed: {e}")

        try:
            if (w_meter or "Common") != "Individual" and (water_note or "").strip():
                update_single_water_sheet_guests(
                    uid, b_month, "tenant", t_name,
                    guests or 0,
                    bill_note=water_note or ""
                )
        except Exception as e:
            log.error(f"Water ledger bill-note refresh failed for bill {bill_id}: {e}")


def recompute_tenant_from_transactions(tenant_id):
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()

    cur.execute("""SELECT t.user_id, t.name, t.tenant_uid,
                          p.name, p.address
                   FROM tenants t JOIN properties p ON t.property_id = p.id
                   WHERE t.id=?""", (tenant_id,))
    row = cur.fetchone()
    if not row:
        conn.close(); return
    uid, t_name, t_uid, p_name, p_addr = row

    cur.execute("""UPDATE tenant_bills
                   SET amount_paid = 0,
                       balance_remaining = CASE WHEN settled_status = 'closed_manual' THEN 0 ELSE total_payable END,
                       payment_status = CASE WHEN settled_status = 'closed_manual' THEN 'Paid' ELSE 'Unpaid' END,
                       settled_status = CASE WHEN settled_status = 'closed_manual' THEN 'closed_manual' ELSE 'unpaid' END
                   WHERE tenant_id=?""", (tenant_id,))

    cur.execute("""SELECT linked_bill_id, COALESCE(SUM(amount),0)
                   FROM tenant_transactions
                   WHERE tenant_id=? AND txn_type IN ('payment','advance')
                     AND linked_bill_id IS NOT NULL
                   GROUP BY linked_bill_id""", (tenant_id,))
    linked_payments = dict(cur.fetchall())

    for bill_id, paid_amt in linked_payments.items():
        cur.execute("SELECT total_payable, settled_status FROM tenant_bills WHERE id=?", (bill_id,))
        b = cur.fetchone()
        if not b: continue
        total, current_status = b
        total = round(float(total or 0.0), 2)
        paid_amt = round(float(paid_amt or 0.0), 2)
        new_paid = min(paid_amt, total)
        new_bal = round(max(0.0, total - new_paid), 2)
        if new_bal <= 0.009:
            new_bal = 0.0
            new_status = "Paid"
            settled = current_status if current_status == "closed_manual" else "closed"
        else:
            new_status = "Partial" if new_paid > 0.009 else "Unpaid"
            settled = current_status if current_status == "closed_manual" else (
                "partial" if new_paid > 0.009 else "unpaid"
            )
        cur.execute("""UPDATE tenant_bills
                       SET amount_paid=?, balance_remaining=?,
                           payment_status=?, settled_status=?
                       WHERE id=?""",
                    (new_paid, new_bal, new_status, settled, bill_id))

    cur.execute("""SELECT COALESCE(SUM(amount),0)
                   FROM tenant_transactions
                   WHERE tenant_id=? AND txn_type IN ('payment','advance')
                     AND linked_bill_id IS NULL""", (tenant_id,))
    unlinked_pool = cur.fetchone()[0] or 0.0
    if unlinked_pool > 0.001:
        cur.execute("""SELECT id, total_payable, amount_paid, settled_status
                       FROM tenant_bills
                       WHERE tenant_id=? AND balance_remaining > 0.001
                         AND settled_status NOT IN ('closed', 'closed_manual')
                       ORDER BY bill_date ASC, id ASC""", (tenant_id,))
        for bid, total, already, current_status in cur.fetchall():
            if unlinked_pool <= 0.001: break
            need = float(total) - float(already or 0.0)
            if need <= 0.001: continue
            take = min(unlinked_pool, need)
            new_paid = float(already or 0.0) + take
            new_bal = max(0.0, float(total) - new_paid)
            if new_bal <= 0.001:
                new_bal = 0.0
                new_status = "Paid"
                settled = current_status if current_status == "closed_manual" else "closed"
            else:
                new_status = "Partial" if new_paid > 0.001 else "Unpaid"
                settled = current_status if current_status == "closed_manual" else (
                    "partial" if new_paid > 0.001 else "unpaid"
                )
            cur.execute("""UPDATE tenant_bills
                           SET amount_paid=?, balance_remaining=?,
                               payment_status=?, settled_status=?
                           WHERE id=?""",
                        (new_paid, new_bal, new_status, settled, bid))
            unlinked_pool -= take

    conn.commit()

    cur.execute("""SELECT DISTINCT linked_bill_id
                   FROM tenant_transactions
                   WHERE tenant_id=? AND linked_bill_id IS NOT NULL
                     AND txn_type IN ('payment','advance')""", (tenant_id,))
    affected_bills = [r[0] for r in cur.fetchall()]
    conn.close()

    try:
        rebuild_tenant_ledger(uid, tenant_id)
    except Exception as e:
        log.error(f"[recompute] ledger rebuild failed for tenant {tenant_id}: {e}")

    for bid in affected_bills:
        try:
            regenerate_paid_bill_pdfs(bid)
        except Exception as e:
            log.error(f"[recompute] bill PDF regen failed for bill {bid}: {e}")

    try:
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        for bid in affected_bills:
            cur.execute("""SELECT b.billing_month, b.water_note,
                                  t.name, b.guests_count
                           FROM tenant_bills b
                           JOIN tenants t ON b.tenant_id = t.id
                           WHERE b.id=?""", (bid,))
            r = cur.fetchone()
            if not r: continue
            b_month, w_note, t_name2, guests_cnt = r
            if (w_note or "").strip():
                update_single_water_sheet_guests(
                    uid, b_month, "tenant", t_name2,
                    guests_cnt or 0,
                    bill_note=w_note or ""
                )
        conn.close()
    except Exception as e:
        log.error(f"[recompute] water ledger bill-note refresh failed: {e}")


def delete_bill_pdfs(user_id, address, property_name, tenant_name, tenant_uid, year_str, b_month):
    """Removes all generated and unpaid bill PDFs for a specific bill and prunes empty folders."""
    for kind in ["Generated", "Unpaid", "Paid"]:
        dir_path = get_bill_dir(user_id, address, property_name, tenant_name, tenant_uid, year_str, b_month, kind)
        if os.path.isdir(dir_path):
            for fname in os.listdir(dir_path):
                if (fname.startswith("Bill_") or fname.startswith("Paid_Bill_")) and fname.endswith(".pdf"):
                    try:
                        os.remove(os.path.join(dir_path, fname))
                    except Exception as e:
                        log.error(f"Failed to remove {fname} in {dir_path}: {e}")
            try:
                if not os.listdir(dir_path):
                    os.rmdir(dir_path)
            except Exception:
                pass
            parent_month_dir = os.path.dirname(dir_path)
            try:
                if os.path.isdir(parent_month_dir) and not os.listdir(parent_month_dir):
                    os.rmdir(parent_month_dir)
            except Exception:
                pass
            parent_year_dir = os.path.dirname(parent_month_dir)
            try:
                if os.path.isdir(parent_year_dir) and not os.listdir(parent_year_dir):
                    os.rmdir(parent_year_dir)
            except Exception:
                pass


def regenerate_bill_documents(user_id, bill_id, old_year=None, old_month=None):
    """Regenerates Generated Bill & Unpaid Bills PDFs, updates water ledger, and rebuilds tenant ledger."""
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""
        SELECT b.user_id, b.billing_month, b.bill_date, b.base_rent,
               b.elec_prev, b.elec_pres, b.elec_rate, b.elec_total,
               b.water_prev, b.water_pres, b.water_rate, b.water_total,
               b.discount_total, b.prev_due, b.total_payable, b.amount_paid,
               b.balance_remaining, b.guests_count, b.owner_members_snapshot, b.remarks,
               t.name, t.tenant_uid, t.permanent_address, t.occupants_count,
               p.name, p.address, p.water_meter_type, pr.owner_name, pr.phone,
               b.common_note, b.water_note, b.tenant_id
        FROM tenant_bills b
        JOIN tenants t ON b.tenant_id = t.id
        JOIN properties p ON t.property_id = p.id
        LEFT JOIN profiles pr ON pr.user_id = b.user_id
        WHERE b.id=?
    """, (bill_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return
    (uid, b_month, b_date, base_rent,
     eprev, epres, erate, e_total,
     wprev, wpres, wrate, w_total,
     discount, prev_due, total_payable, amount_paid,
     balance_remaining, guests, owner_snapshot, remarks,
     t_name, t_uid, t_addr, t_occ,
     p_name, p_addr, w_meter, owner_name, owner_phone,
     common_note, water_note, tenant_id) = row
    conn.close()

    new_year = b_date.split("-")[0] if b_date else str(datetime.now().year)

    if old_month and old_year and (str(old_month) != str(b_month) or str(old_year) != str(new_year)):
        try:
            delete_bill_pdfs(uid, p_addr, p_name, t_name, t_uid, str(old_year), str(old_month))
        except Exception as e:
            log.error(f"Error cleaning up old bill PDFs on month change: {e}")

    gen_bill_dir = get_bill_dir(uid, p_addr, p_name, t_name, t_uid, new_year, b_month, "Generated")
    unpaid_bill_dir = get_bill_dir(uid, p_addr, p_name, t_name, t_uid, new_year, b_month, "Unpaid")

    payload = {
        "bill_id": bill_id,
        "owner_name": owner_name or "Property Administrator",
        "owner_phone": owner_phone or "—",
        "property_name": p_name,
        "property_address": p_addr,
        "tenant_name": t_name,
        "tenant_uid": t_uid,
        "occupants_count": t_occ,
        "billing_month": b_month,
        "bill_date": b_date,
        "base_rent": base_rent,
        "elec_prev": eprev,
        "elec_pres": epres,
        "elec_rate": erate,
        "elec_total": e_total,
        "water_prev": wprev,
        "water_pres": wpres,
        "water_rate": wrate,
        "water_total": w_total,
        "is_common_water": (w_meter or "Common") != "Individual",
        "discount_total": discount or 0,
        "discount_reason": remarks or "",
        "render_discount_customer": True,
        "prev_due": prev_due or 0,
        "total_payable": total_payable,
        "amount_paid": amount_paid or 0.0,
        "balance_remaining": balance_remaining if balance_remaining is not None else total_payable,
        "remarks": remarks or "",
        "guests_count": guests or 0,
        "owner_members_snapshot": owner_snapshot or 1,
        "total_persons": (owner_snapshot or 1) + (t_occ or 1) + (guests or 0),
        "common_note": common_note or "",
        "water_note": water_note or "",
    }

    ensure_dir(gen_bill_dir)
    generate_bill_pdfs(gen_bill_dir, payload, is_paid=False)

    if (balance_remaining or 0) <= 0.001:
        if os.path.isdir(unpaid_bill_dir):
            for fname in os.listdir(unpaid_bill_dir):
                if (fname.startswith("Bill_") or fname.startswith("Paid_Bill_")) and fname.endswith(".pdf"):
                    try: os.remove(os.path.join(unpaid_bill_dir, fname))
                    except Exception: pass
            try:
                if not os.listdir(unpaid_bill_dir):
                    os.rmdir(unpaid_bill_dir)
            except Exception: pass
    else:
        ensure_dir(unpaid_bill_dir)
        generate_bill_pdfs(unpaid_bill_dir, payload, is_paid=False)

    if (w_meter or "Common") != "Individual":
        try:
            update_single_water_sheet_guests(
                uid, b_month, "tenant", t_name,
                guests or 0,
                bill_note=water_note or ""
            )
        except Exception as e:
            log.error(f"Error updating water ledger on bill edit: {e}")

    try:
        rebuild_tenant_ledger(uid, tenant_id)
    except Exception as e:
        log.error(f"Error rebuilding tenant ledger on bill edit: {e}")


def delete_bill_record(user_id, bill_id):
    """Deletes a generated bill, unlinks payments, cleans up PDFs, updates water ledger, and rebuilds tenant ledger."""
    conn = sqlite3.connect(get_db_path())
    cur = conn.cursor()
    cur.execute("""
        SELECT b.tenant_id, b.billing_month, b.bill_date, b.total_payable,
               t.name, t.tenant_uid, p.name, p.address, p.water_meter_type
        FROM tenant_bills b
        JOIN tenants t ON b.tenant_id = t.id
        JOIN properties p ON t.property_id = p.id
        WHERE b.id=?
    """, (bill_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Bill record not found."

    (tenant_id, b_month, b_date, total_payable,
     t_name, t_uid, p_name, p_addr, w_meter) = row

    old_values = {
        "bill_id": bill_id, "billing_month": b_month, "bill_date": b_date,
        "total_payable": total_payable, "tenant_id": tenant_id
    }

    # 1. Unlink any payments attached to this bill so tenant retains their payment records
    cur.execute("""
        UPDATE tenant_transactions
        SET linked_bill_id = NULL
        WHERE linked_bill_id = ? AND txn_type IN ('payment', 'advance')
    """, (bill_id,))

    # 2. Delete the bill transaction
    cur.execute("""
        DELETE FROM tenant_transactions
        WHERE linked_bill_id = ? AND txn_type = 'bill'
    """, (bill_id,))

    # 3. Delete the bill record
    cur.execute("DELETE FROM tenant_bills WHERE id = ?", (bill_id,))
    conn.commit()
    conn.close()

    # 4. Log edit history
    log_edit(user_id, tenant_id, bill_id, "delete_bill", old_values, None)

    # 5. Clean up PDFs for this bill
    year_str = b_date.split("-")[0] if b_date else str(datetime.now().year)
    try:
        delete_bill_pdfs(user_id, p_addr, p_name, t_name, t_uid, year_str, b_month)
    except Exception as e:
        log.error(f"Error deleting bill PDFs: {e}")

    # 6. Reset guests and water note in Master_Water_Ledger.xlsx if common water
    if (w_meter or "Common") != "Individual":
        try:
            update_single_water_sheet_guests(user_id, b_month, "tenant", t_name, 0, bill_note="")
        except Exception as e:
            log.error(f"Error resetting water ledger on bill delete: {e}")

    # 7. Recompute tenant transactions & rebuild Excel ledger
    try:
        recompute_tenant_from_transactions(tenant_id)
    except Exception as e:
        log.error(f"Error recomputing tenant after bill delete: {e}")

    return True, f"Bill for {b_month} deleted successfully."


def delete_receipt_pdfs_for_txn(user_id, tenant_id, txn_id):
    """Deletes existing receipt PDFs for a given transaction id."""
    try:
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.name, t.tenant_uid, p.name, p.address
                       FROM tenants t
                       JOIN properties p ON t.property_id = p.id
                       WHERE t.id = ?""", (tenant_id,))
        row = cur.fetchone()
        conn.close()
        if not row: return
        t_name, t_uid, p_name, p_addr = row
        tag = f"R{txn_id:05d}"
        tenant_dir = get_tenant_dir(user_id, p_addr, p_name, t_name, t_uid)
        if os.path.isdir(tenant_dir):
            for root_d, dirs, files in os.walk(tenant_dir):
                for fname in files:
                    if tag in fname and fname.endswith(".pdf"):
                        try:
                            os.remove(os.path.join(root_d, fname))
                        except Exception: pass
    except Exception as e:
        log.error(f"Error in delete_receipt_pdfs_for_txn: {e}")


def regenerate_receipt_pdfs_for_txn(user_id, txn_id):
    """Regenerates the receipt PDFs for a transaction with updated date/amount."""
    try:
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT tt.tenant_id, tt.property_id, tt.txn_date, tt.txn_type,
                              tt.description, tt.amount, tt.payment_method,
                              tt.reference_no, tt.linked_bill_id, tt.notes,
                              t.name, t.tenant_uid, p.name, p.address,
                              pr.owner_name, pr.phone
                       FROM tenant_transactions tt
                       JOIN tenants t ON tt.tenant_id = t.id
                       JOIN properties p ON tt.property_id = p.id
                       LEFT JOIN profiles pr ON pr.user_id = ?
                       WHERE tt.id=?""", (user_id, txn_id))
        r = cur.fetchone()
        if not r:
            conn.close()
            return
        (tid, pid, tdate, ttype, desc, amount, method,
         ref, linked_bill, notes, t_name, t_uid, p_name, p_addr,
         owner_name, owner_phone) = r

        tag = f"R{txn_id:05d}"
        tenant_dir = get_tenant_dir(user_id, p_addr, p_name, t_name, t_uid)
        if os.path.isdir(tenant_dir):
            for root_d, dirs, files in os.walk(tenant_dir):
                for fname in files:
                    if tag in fname and fname.endswith(".pdf"):
                        try:
                            os.remove(os.path.join(root_d, fname))
                        except Exception: pass

        y = int(tdate.split("-")[0]) if tdate else datetime.now().year
        b_month_target = None
        bill_total = 0.0
        bill_rem = 0.0
        if linked_bill:
            cur.execute("SELECT billing_month, bill_date, total_payable, balance_remaining FROM tenant_bills WHERE id=?", (linked_bill,))
            brow = cur.fetchone()
            if brow:
                b_month_target = brow[0]
                if brow[1]:
                    try: y = int(brow[1].split("-")[0])
                    except Exception: pass
                bill_total = brow[2] or 0.0
                bill_rem = brow[3] or 0.0
        conn.close()

        if not b_month_target:
            try:
                b_month_target = datetime.strptime(tdate, "%Y-%m-%d").strftime("%b %Y")
            except Exception:
                b_month_target = datetime.now().strftime("%b %Y")

        gen = get_bill_dir(user_id, p_addr, p_name, t_name, t_uid, y, b_month_target, "Receipts")
        bal_after = get_tenant_balance(tid)

        payload = {
            "txn_id": txn_id,
            "owner_name": owner_name or "Property Administrator",
            "owner_phone": owner_phone or "—",
            "property_name": p_name,
            "property_address": p_addr,
            "tenant_name": t_name,
            "tenant_uid": t_uid,
            "amount": amount,
            "payment_date": tdate,
            "paid_date": tdate,
            "payment_method": method or "—",
            "reference_no": ref or "",
            "description": desc or ("Advance payment" if ttype == "advance" else "Payment"),
            "balance_after": bal_after,
        }
        if linked_bill:
            payload["linked_bill_month"] = b_month_target
            payload["linked_bill_amount"] = bill_total
            payload["linked_bill_remaining"] = bill_rem

        ensure_dir(gen)
        generate_receipt_pdfs(gen, payload)
    except Exception as e:
        log.error(f"Error in regenerate_receipt_pdfs_for_txn: {e}")


class EditBillDialog(GlassDialog):
    """Full-featured modal dialog to review, edit, recalculate, and correct a recorded bill."""
    def __init__(self, user_id, bill_id, parent=None):
        super().__init__(parent, width=620, height=720)
        self.user_id = user_id
        self.bill_id = bill_id
        self.bill_data = None
        self._load_bill()
        self.init_ui()

    def _load_bill(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""
            SELECT b.id, b.user_id, b.tenant_id, b.billing_month, b.bill_date, b.base_rent,
                   b.elec_prev, b.elec_pres, b.elec_units, b.elec_rate, b.elec_total,
                   b.water_prev, b.water_pres, b.water_units, b.water_rate, b.water_total,
                   b.discount_total, b.prev_due, b.total_payable, b.amount_paid,
                   b.balance_remaining, b.payment_status, b.settled_status, b.remarks,
                   b.guests_count, b.common_note, b.water_note,
                   t.name, t.tenant_uid, t.occupants_count,
                   p.name, p.address, p.water_meter_type, p.id
            FROM tenant_bills b
            JOIN tenants t ON b.tenant_id = t.id
            JOIN properties p ON t.property_id = p.id
            WHERE b.id=?
        """, (self.bill_id,))
        self.bill_data = cur.fetchone()
        conn.close()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 22, 28, 22)
        lay.setSpacing(10)

        if not self.bill_data:
            lbl = QLabel("Bill statement not found.")
            lbl.setStyleSheet(f"color: {ACCENT_RED}; font-size: 13px; font-family: {FONT_STACK};")
            lay.addWidget(lbl)
            base.addWidget(env)
            return

        (bid, uid, tid, b_month, b_date, base_rent,
         eprev, epres, eunits, erate, etotal,
         wprev, wpres, wunits, wrate, wtotal,
         disc, prev_due, total_payable, amount_paid,
         balance_remaining, pay_status, settled_status, remarks,
         guests, common_note, water_note,
         t_name, t_uid, t_occ,
         p_name, p_addr, w_meter, pid) = self.bill_data

        self.tid = tid
        self.t_name = t_name
        self.t_uid = t_uid
        self.t_occ = t_occ or 1
        self.p_name = p_name
        self.p_addr = p_addr
        self.is_common_water = (w_meter or "Common") != "Individual"
        self.old_b_month = b_month
        self.old_b_date = b_date
        self.old_total_payable = total_payable
        self.amount_paid_snapshot = amount_paid or 0.0

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("edit", color=ACCENT_INDIGO, size=20))
        t = QLabel("EDIT BILL STATEMENT")
        t.setStyleSheet(f"font-size: 14px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        chip = QLabel(f"<b>{t_name}</b> ({t_uid}) — <b>{p_name}</b> &nbsp;·&nbsp; Bill #{bid} ({b_month})")
        chip.setStyleSheet(f"""
            background: rgba(16,185,129,0.12); color: {T.c('text_primary')};
            border: 1px solid rgba(16,185,129,0.30);
            border-radius: {RADIUS_INPUT}px; padding: 7px 12px;
            font-size: 12px; font-family: {FONT_STACK};
        """)
        lay.addWidget(chip)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 20px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        c_widget = QWidget()
        c_widget.setStyleSheet("background: transparent;")
        cl = QVBoxLayout(c_widget)
        cl.setContentsMargins(0, 0, 8, 0)
        cl.setSpacing(12)

        # 1. Billing Cycle & Date Card
        c1 = GlassCard()
        l1 = QVBoxLayout(c1)
        l1.setContentsMargins(16, 12, 16, 12)
        l1.setSpacing(8)
        l1.addWidget(self._lbl("BILLING CYCLE & DATE *"))

        row_dt = QHBoxLayout()
        row_dt.setSpacing(10)

        v_m = QVBoxLayout(); v_m.setSpacing(3); v_m.addWidget(self._lbl("MONTH"))
        self.combo_month = QComboBox()
        setup_combobox(self.combo_month)
        self.combo_month.addItems(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
        self.combo_month.setStyleSheet(INPUT_STYLE)
        self.combo_month.setFixedHeight(INPUT_H)
        v_m.addWidget(self.combo_month)

        v_y = QVBoxLayout(); v_y.setSpacing(3); v_y.addWidget(self._lbl("YEAR"))
        self.combo_year = QComboBox()
        setup_combobox(self.combo_year)
        cy = QDate.currentDate().year()
        for y in range(cy - 3, cy + 5): self.combo_year.addItem(str(y))
        self.combo_year.setStyleSheet(INPUT_STYLE)
        self.combo_year.setFixedHeight(INPUT_H)
        v_y.addWidget(self.combo_year)

        parts = b_month.split()
        if len(parts) >= 1 and parts[0] in [self.combo_month.itemText(i) for i in range(12)]:
            self.combo_month.setCurrentText(parts[0])
        if len(parts) >= 2 and parts[1] in [self.combo_year.itemText(i) for i in range(self.combo_year.count())]:
            self.combo_year.setCurrentText(parts[1])

        v_d = QVBoxLayout(); v_d.setSpacing(3); v_d.addWidget(self._lbl("BILL DATE"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        try:
            yy, mm, dd = map(int, b_date.split("-"))
            self.in_date.setDate(QDate(yy, mm, dd))
        except Exception: pass
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setStyleSheet(INPUT_STYLE)
        self.in_date.setFixedHeight(INPUT_H)
        v_d.addWidget(self.in_date)

        row_dt.addLayout(v_m, 1)
        row_dt.addLayout(v_y, 1)
        row_dt.addLayout(v_d, 1)
        l1.addLayout(row_dt)
        cl.addWidget(c1)

        # 2. Base Rent Card
        c2 = GlassCard()
        l2 = QVBoxLayout(c2)
        l2.setContentsMargins(16, 12, 16, 12)
        l2.setSpacing(6)
        l2.addWidget(self._lbl("MONTHLY BASE RENT (₹) *"))
        self.in_rent = QLineEdit(f"{float(base_rent or 0):.2f}")
        self.in_rent.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_rent.setStyleSheet(INPUT_STYLE)
        self.in_rent.setFixedHeight(INPUT_H)
        self.in_rent.textChanged.connect(self.recalc)
        l2.addWidget(self.in_rent)
        cl.addWidget(c2)

        # 3. Electricity Card
        c3 = GlassCard()
        l3 = QVBoxLayout(c3)
        l3.setContentsMargins(16, 12, 16, 12)
        l3.setSpacing(8)
        l3.addWidget(self._lbl("ELECTRICITY METER READINGS"))
        r_el = QHBoxLayout()
        r_el.setSpacing(10)
        v_eprev = QVBoxLayout(); v_eprev.setSpacing(3); v_eprev.addWidget(self._lbl("PREVIOUS READING"))
        self.in_eprev = QLineEdit(f"{float(eprev or 0):g}")
        self.in_eprev.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_eprev.setStyleSheet(INPUT_STYLE)
        self.in_eprev.setFixedHeight(INPUT_H)
        self.in_eprev.textChanged.connect(self.recalc)
        v_eprev.addWidget(self.in_eprev)

        v_epres = QVBoxLayout(); v_epres.setSpacing(3); v_epres.addWidget(self._lbl("PRESENT READING"))
        self.in_epres = QLineEdit(f"{float(epres or 0):g}")
        self.in_epres.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
        self.in_epres.setStyleSheet(INPUT_STYLE)
        self.in_epres.setFixedHeight(INPUT_H)
        self.in_epres.textChanged.connect(self.recalc)
        v_epres.addWidget(self.in_epres)

        v_erate = QVBoxLayout(); v_erate.setSpacing(3); v_erate.addWidget(self._lbl("RATE / UNIT (₹)"))
        self.in_erate = QLineEdit(f"{float(erate or 0):g}")
        self.in_erate.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
        self.in_erate.setStyleSheet(INPUT_STYLE)
        self.in_erate.setFixedHeight(INPUT_H)
        self.in_erate.textChanged.connect(self.recalc)
        v_erate.addWidget(self.in_erate)

        r_el.addLayout(v_eprev, 1)
        r_el.addLayout(v_epres, 1)
        r_el.addLayout(v_erate, 1)
        l3.addLayout(r_el)

        self.lbl_eres = QLabel("Electricity: ₹0.00")
        self.lbl_eres.setStyleSheet(f"color: {ACCENT_AMBER}; font-size: 12px; font-weight: 700; font-family: {FONT_STACK};")
        l3.addWidget(self.lbl_eres)
        cl.addWidget(c3)

        # 4. Water Card
        c4 = GlassCard()
        l4 = QVBoxLayout(c4)
        l4.setContentsMargins(16, 12, 16, 12)
        l4.setSpacing(8)

        if self.is_common_water:
            l4.addWidget(self._lbl("WATER (CAMPUS SHARED PUMP ALLOCATION)"))
            r_wt = QHBoxLayout()
            r_wt.setSpacing(10)

            v_wflat = QVBoxLayout(); v_wflat.setSpacing(3); v_wflat.addWidget(self._lbl("FAMILY HEADCOUNT"))
            self.lbl_wflat = QLabel(f"{self.t_occ} persons")
            self.lbl_wflat.setStyleSheet(f"""
                background-color: {T.c('input_bg')};
                border: 1px solid {T.c('input_border')};
                border-radius: {RADIUS_INPUT}px;
                padding: 9px 13px; color: {T.c('text_primary')};
                font-size: 13px; font-weight: 700;
                font-family: {FONT_STACK};
            """)
            self.lbl_wflat.setFixedHeight(INPUT_H)
            v_wflat.addWidget(self.lbl_wflat)

            v_wguest = QVBoxLayout(); v_wguest.setSpacing(3); v_wguest.addWidget(self._lbl("GUESTS THIS MONTH"))
            self.in_guests = QLineEdit(str(int(guests or 0)))
            self.in_guests.setValidator(QDoubleValidator(0, 100, 0, self))
            self.in_guests.setStyleSheet(INPUT_STYLE)
            self.in_guests.setFixedHeight(INPUT_H)
            self.in_guests.textChanged.connect(self.recalc)
            v_wguest.addWidget(self.in_guests)

            v_wrate = QVBoxLayout(); v_wrate.setSpacing(3); v_wrate.addWidget(self._lbl("RATE / PERSON (₹)"))
            self.in_wrate_ph = QLineEdit(f"{float(wrate or 0):.2f}")
            self.in_wrate_ph.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
            self.in_wrate_ph.setStyleSheet(INPUT_STYLE)
            self.in_wrate_ph.setFixedHeight(INPUT_H)
            self.in_wrate_ph.textChanged.connect(self.recalc)
            v_wrate.addWidget(self.in_wrate_ph)

            r_wt.addLayout(v_wflat, 1)
            r_wt.addLayout(v_wguest, 1)
            r_wt.addLayout(v_wrate, 1)
            l4.addLayout(r_wt)
        else:
            l4.addWidget(self._lbl("WATER (INDIVIDUAL METER READINGS)"))
            r_wt = QHBoxLayout()
            r_wt.setSpacing(10)

            v_wprev = QVBoxLayout(); v_wprev.setSpacing(3); v_wprev.addWidget(self._lbl("PREVIOUS READING"))
            self.in_wprev = QLineEdit(f"{float(wprev or 0):g}")
            self.in_wprev.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
            self.in_wprev.setStyleSheet(INPUT_STYLE)
            self.in_wprev.setFixedHeight(INPUT_H)
            self.in_wprev.textChanged.connect(self.recalc)
            v_wprev.addWidget(self.in_wprev)

            v_wpres = QVBoxLayout(); v_wpres.setSpacing(3); v_wpres.addWidget(self._lbl("PRESENT READING"))
            self.in_wpres = QLineEdit(f"{float(wpres or 0):g}")
            self.in_wpres.setValidator(QDoubleValidator(0.0, 1e9, 2, self))
            self.in_wpres.setStyleSheet(INPUT_STYLE)
            self.in_wpres.setFixedHeight(INPUT_H)
            self.in_wpres.textChanged.connect(self.recalc)
            v_wpres.addWidget(self.in_wpres)

            v_wrate = QVBoxLayout(); v_wrate.setSpacing(3); v_wrate.addWidget(self._lbl("RATE / UNIT (₹)"))
            self.in_wrate = QLineEdit(f"{float(wrate or 0):g}")
            self.in_wrate.setValidator(QDoubleValidator(0.0, 1e5, 2, self))
            self.in_wrate.setStyleSheet(INPUT_STYLE)
            self.in_wrate.setFixedHeight(INPUT_H)
            self.in_wrate.textChanged.connect(self.recalc)
            v_wrate.addWidget(self.in_wrate)

            r_wt.addLayout(v_wprev, 1)
            r_wt.addLayout(v_wpres, 1)
            r_wt.addLayout(v_wrate, 1)
            l4.addLayout(r_wt)

        self.lbl_wres = QLabel("Water: ₹0.00")
        self.lbl_wres.setStyleSheet(f"color: {ACCENT_INDIGO}; font-size: 12px; font-weight: 700; font-family: {FONT_STACK};")
        l4.addWidget(self.lbl_wres)
        cl.addWidget(c4)

        # 5. Concessions & Discounts Card
        c5 = GlassCard()
        l5 = QVBoxLayout(c5)
        l5.setContentsMargins(16, 12, 16, 12)
        l5.setSpacing(8)
        l5.addWidget(self._lbl("CONCESSION & ADJUSTMENTS"))

        r_disc = QHBoxLayout()
        r_disc.setSpacing(10)
        v_damt = QVBoxLayout(); v_damt.setSpacing(3); v_damt.addWidget(self._lbl("DISCOUNT AMOUNT (₹)"))
        self.in_discount = QLineEdit(f"{float(disc or 0):.2f}")
        self.in_discount.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_discount.setStyleSheet(INPUT_STYLE)
        self.in_discount.setFixedHeight(INPUT_H)
        self.in_discount.textChanged.connect(self.recalc)
        v_damt.addWidget(self.in_discount)

        v_dres = QVBoxLayout(); v_dres.setSpacing(3); v_dres.addWidget(self._lbl("DISCOUNT REASON"))
        self.in_reason = QLineEdit(remarks or "")
        self.in_reason.setPlaceholderText("Reason for discount...")
        self.in_reason.setStyleSheet(INPUT_STYLE)
        self.in_reason.setFixedHeight(INPUT_H)
        v_dres.addWidget(self.in_reason)
        r_disc.addLayout(v_damt, 1)
        r_disc.addLayout(v_dres, 2)
        l5.addLayout(r_disc)

        self.chk_render = ToggleCheckBox("Render discount on customer PDF copy")
        self.chk_render.setChecked(True)
        l5.addWidget(self.chk_render)
        cl.addWidget(c5)

        # 6. Previous Due Card
        c6 = GlassCard()
        l6 = QVBoxLayout(c6)
        l6.setContentsMargins(16, 12, 16, 12)
        l6.setSpacing(6)
        l6.addWidget(self._lbl("PREVIOUS BALANCE CARRIED FORWARD (₹)"))
        self.in_prev_due = QLineEdit(f"{float(prev_due or 0):.2f}")
        self.in_prev_due.setValidator(QDoubleValidator(-1e10, 1e10, 2, self))
        self.in_prev_due.setStyleSheet(INPUT_STYLE)
        self.in_prev_due.setFixedHeight(INPUT_H)
        self.in_prev_due.textChanged.connect(self.recalc)
        l6.addWidget(self.in_prev_due)
        cl.addWidget(c6)

        # 7. Owner Internal Notes Card
        c7 = GlassCard()
        l7 = QVBoxLayout(c7)
        l7.setContentsMargins(16, 12, 16, 12)
        l7.setSpacing(8)
        l7.addWidget(self._lbl("OWNER AUDIT NOTES (Internal Only)"))

        l7.addWidget(self._lbl("RENT & ELECTRICITY NOTE"))
        self.in_common_note = QTextEdit(common_note or "")
        self.in_common_note.setStyleSheet(INPUT_STYLE)
        self.in_common_note.setMaximumHeight(54)
        l7.addWidget(self.in_common_note)

        l7.addWidget(self._lbl("WATER NOTE"))
        self.in_water_note = QTextEdit(water_note or "")
        self.in_water_note.setStyleSheet(INPUT_STYLE)
        self.in_water_note.setMaximumHeight(54)
        l7.addWidget(self.in_water_note)
        cl.addWidget(c7)

        # 8. Financial Summary Card
        c8 = GlassCard()
        l8 = QVBoxLayout(c8)
        l8.setContentsMargins(18, 14, 18, 14)
        l8.setSpacing(6)

        self.lbl_grand = QLabel("TOTAL BILL: ₹0.00")
        self.lbl_grand.setStyleSheet(f"font-size: 20px; font-weight: 900; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        l8.addWidget(self.lbl_grand)

        r_bal_info = QHBoxLayout()
        self.lbl_paid_info = QLabel(f"Amount Paid So Far: {fmt_money(amount_paid or 0)}")
        self.lbl_paid_info.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {ACCENT_GREEN}; font-family: {FONT_STACK};")
        self.lbl_rem_info = QLabel("Remaining Balance: ₹0.00")
        self.lbl_rem_info.setStyleSheet(f"font-size: 12px; font-weight: 700; color: {ACCENT_RED}; font-family: {FONT_STACK};")
        r_bal_info.addWidget(self.lbl_paid_info)
        r_bal_info.addStretch(1)
        r_bal_info.addWidget(self.lbl_rem_info)
        l8.addLayout(r_bal_info)
        cl.addWidget(c8)

        scroll.setWidget(c_widget)
        lay.addWidget(scroll, 1)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        del_btn = QPushButton("DELETE BILL")
        del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.setFixedHeight(BTN_PRIMARY)
        del_btn.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_PRIMARY))
        del_btn.clicked.connect(self.delete_bill)

        cancel_btn = QPushButton("CANCEL")
        cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel_btn.setFixedHeight(BTN_PRIMARY)
        cancel_btn.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel_btn.clicked.connect(self.reject)

        save_btn = QPushButton("SAVE CHANGES")
        save_btn.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.setFixedHeight(BTN_PRIMARY)
        save_btn.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save_btn.clicked.connect(self.save_bill)

        btn_row.addWidget(del_btn, 1)
        btn_row.addWidget(cancel_btn, 1)
        btn_row.addWidget(save_btn, 2)
        lay.addLayout(btn_row)

        base.addWidget(env)
        self.recalc()

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.5px; background: transparent; font-family: {FONT_STACK};")
        return l

    def recalc(self):
        if not hasattr(self, "lbl_grand"): return 0.0, 0.0, 0.0
        try:
            eprev = float(self.in_eprev.text() or 0)
            epres = float(self.in_epres.text() or 0)
            erate = float(self.in_erate.text() or 0)
            eunits = max(0.0, epres - eprev)
            etotal = round(eunits * erate, 2)
            self.lbl_eres.setText(f"Units: {eunits:,.2f}  ·  Total: {fmt_money(etotal)}")
        except ValueError:
            etotal = 0.0

        if self.is_common_water:
            try:
                guests = float(self.in_guests.text() or 0)
                wrate = float(self.in_wrate_ph.text() or 0)
                wtotal = round((self.t_occ + guests) * wrate, 2)
                self.lbl_wres.setText(f"Water Share: {fmt_money(wtotal)}  ·  ({self.t_occ} family + {int(guests)} guests × {fmt_money(wrate)})")
            except ValueError:
                wtotal = 0.0
        else:
            try:
                wprev = float(self.in_wprev.text() or 0)
                wpres = float(self.in_wpres.text() or 0)
                wrate = float(self.in_wrate.text() or 0)
                wunits = max(0.0, wpres - wprev)
                wtotal = round(wunits * wrate, 2)
                self.lbl_wres.setText(f"Units: {wunits:,.2f}  ·  Total: {fmt_money(wtotal)}")
            except ValueError:
                wtotal = 0.0

        try:
            rent = float(self.in_rent.text() or 0)
            discount = float(self.in_discount.text() or 0)
            prev_due = float(self.in_prev_due.text() or 0)
            total = max(0.0, rent + etotal + wtotal - discount + prev_due)
            rem = max(0.0, total - self.amount_paid_snapshot)
            self.lbl_grand.setText(f"TOTAL BILL: {fmt_money(total)}")
            self.lbl_rem_info.setText(f"Remaining Balance: {fmt_money(rem)}")
            return total, etotal, wtotal
        except ValueError:
            return 0.0, 0.0, 0.0

    def save_bill(self):
        try:
            eprev = float(self.in_eprev.text() or 0)
            epres = float(self.in_epres.text() or 0)
            erate = float(self.in_erate.text() or 0)
        except ValueError:
            QMessageBox.warning(self, "Validation", "Electricity values must be numbers."); return
        if epres < eprev:
            QMessageBox.warning(self, "Validation", "Electricity present reading cannot be less than previous."); return

        if self.is_common_water:
            try:
                guests = int(float(self.in_guests.text() or 0))
                wrate = float(self.in_wrate_ph.text() or 0)
                wtotal = (self.t_occ + guests) * wrate
                wprev = 0.0; wpres = 0.0; wunits = 0.0
            except ValueError:
                QMessageBox.warning(self, "Validation", "Water values must be numbers."); return
        else:
            try:
                wprev = float(self.in_wprev.text() or 0)
                wpres = float(self.in_wpres.text() or 0)
                wrate = float(self.in_wrate.text() or 0)
                wunits = max(0.0, wpres - wprev)
                wtotal = wunits * wrate
                guests = 0
            except ValueError:
                QMessageBox.warning(self, "Validation", "Water values must be numbers."); return
            if wpres < wprev:
                QMessageBox.warning(self, "Validation", "Water present reading cannot be less than previous."); return

        try:
            rent = float(self.in_rent.text() or 0)
            discount = float(self.in_discount.text() or 0)
            prev_due = float(self.in_prev_due.text() or 0)
        except ValueError:
            QMessageBox.warning(self, "Validation", "Financial amounts must be numbers."); return

        if rent <= 0:
            QMessageBox.warning(self, "Validation", "Base rent must be greater than zero."); return

        eunits = max(0.0, epres - eprev)
        etotal = eunits * erate
        grand_total = round(max(0.0, rent + etotal + wtotal - discount + prev_due), 2)

        new_b_month = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        new_b_date = self.in_date.date().toString("yyyy-MM-dd")
        new_remarks = self.in_reason.text().strip()
        new_common_note = self.in_common_note.toPlainText().strip()
        new_water_note = self.in_water_note.toPlainText().strip()

        old_year = self.old_b_date.split("-")[0] if self.old_b_date else str(datetime.now().year)

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""UPDATE tenant_bills SET
            billing_month=?, bill_date=?, base_rent=?,
            elec_prev=?, elec_pres=?, elec_units=?, elec_rate=?, elec_total=?,
            water_prev=?, water_pres=?, water_units=?, water_rate=?, water_total=?,
            discount_total=?, prev_due=?, total_payable=?,
            remarks=?, guests_count=?,
            common_note=?, water_note=?
            WHERE id=?""",
            (new_b_month, new_b_date, rent,
             eprev, epres, eunits, erate, etotal,
             wprev, wpres, wunits, wrate, wtotal,
             discount, prev_due, grand_total,
             new_remarks, guests,
             new_common_note, new_water_note, self.bill_id))

        cur.execute("""UPDATE tenant_transactions SET
            txn_date=?, description=?, amount=?
            WHERE linked_bill_id=? AND txn_type='bill'""",
            (new_b_date, f"Bill for {new_b_month}", grand_total, self.bill_id))
        conn.commit()
        conn.close()

        log_edit(self.user_id, self.tid, self.bill_id, "edit_bill",
                 {"billing_month": self.old_b_month, "bill_date": self.old_b_date, "total_payable": self.old_total_payable},
                 {"billing_month": new_b_month, "bill_date": new_b_date, "total_payable": grand_total})

        # Update water ledger if month changed
        if self.is_common_water and self.old_b_month != new_b_month:
            try:
                update_single_water_sheet_guests(self.user_id, self.old_b_month, "tenant", self.t_name, 0, bill_note="")
            except Exception as e:
                log.error(f"Failed to reset old month in water ledger: {e}")

        # Recompute tenant payments and ledger
        try:
            recompute_tenant_from_transactions(self.tid)
        except Exception as e:
            log.error(f"Recompute after bill edit failed: {e}")

        # Regenerate PDFs for this bill (and clean up old month/year if changed)
        try:
            regenerate_bill_documents(self.user_id, self.bill_id, old_year=old_year, old_month=self.old_b_month)
        except Exception as e:
            log.error(f"PDF regeneration after bill edit failed: {e}")
            QMessageBox.warning(self, "PDF Warning", f"Bill updated in database and Excel, but PDF generation had an error:\n\n{e}")
            self.accept()
            return

        QMessageBox.information(
            self, "Bill Updated",
            f"Bill statement for {new_b_month} updated successfully.\n\n"
            f"• Tenant Master Excel ledger updated\n"
            f"• All bill PDFs regenerated with corrected figures\n"
            f"• Account balance and payment allocations reconciled"
        )
        self.accept()

    def delete_bill(self):
        if QMessageBox.question(
            self, "Delete Bill",
            f"Permanently delete the bill for {self.old_b_month} ({fmt_money(self.old_total_payable)})?\n\n"
            "This will remove the bill, update the tenant's running balance, rebuild the Excel ledger, "
            "and remove all generated bill PDFs.\n\nThis cannot be undone. Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return

        ok, msg = delete_bill_record(self.user_id, self.bill_id)
        if ok:
            QMessageBox.information(self, "Deleted", msg)
            self.accept()
        else:
            QMessageBox.warning(self, "Delete Failed", msg)


class EditPaymentDialog(GlassDialog):
    def __init__(self, user_id, txn_id, parent=None):
        super().__init__(parent, width=500)
        self.user_id = user_id
        self.txn_id = txn_id
        self.txn = None
        self._load_txn()
        self.init_ui()

    def _load_txn(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT tt.tenant_id, tt.property_id, tt.txn_date, tt.txn_type,
                              tt.description, tt.amount, tt.payment_method,
                              tt.reference_no, tt.linked_bill_id, tt.notes,
                              t.name, t.tenant_uid, p.name
                       FROM tenant_transactions tt
                       JOIN tenants t ON tt.tenant_id = t.id
                       JOIN properties p ON tt.property_id = p.id
                       WHERE tt.id=?""", (self.txn_id,))
        self.txn = cur.fetchone()
        conn.close()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        if not self.txn:
            lbl = QLabel("Transaction record not found.")
            lbl.setStyleSheet(f"color: {ACCENT_RED}; font-size: 12px;")
            lay.addWidget(lbl)
            base.addWidget(env)
            return

        (tid, pid, tdate, ttype, desc, amount, method,
         ref, linked_bill, notes, t_name, t_uid, p_name) = self.txn

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("edit", color=ACCENT_INDIGO, size=18))
        t = QLabel("EDIT PAYMENT RECORD")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        chip = QLabel(f"<b>{t_name}</b> ({t_uid}) — {p_name}")
        chip.setStyleSheet(f"""
            background: rgba(16,185,129,0.12); color: {T.c('text_primary')};
            border: 1px solid rgba(16,185,129,0.30);
            border-radius: {RADIUS_INPUT}px; padding: 8px 12px;
            font-size: 12px; font-family: {FONT_STACK};
        """)
        lay.addWidget(chip)

        lay.addWidget(self._lbl("Amount (₹) *"))
        self.in_amount = QLineEdit(f"{float(amount):.2f}")
        self.in_amount.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_amount.setStyleSheet(INPUT_STYLE)
        self.in_amount.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_amount)

        lay.addWidget(self._lbl("Paid Date *"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        try:
            y, m, d = map(int, tdate.split("-"))
            self.in_date.setDate(QDate(y, m, d))
        except Exception:
            pass
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setStyleSheet(INPUT_STYLE)
        self.in_date.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_date)

        lay.addWidget(self._lbl("Payment Method"))
        self.combo_method = QComboBox()
        self.combo_method.addItems(["Cash", "UPI", "Bank Transfer", "Cheque", "Card", "Other"])
        if method and method in [self.combo_method.itemText(i) for i in range(self.combo_method.count())]:
            self.combo_method.setCurrentText(method)
        self.combo_method.setStyleSheet(INPUT_STYLE)
        self.combo_method.setFixedHeight(INPUT_H)
        lay.addWidget(self.combo_method)

        lay.addWidget(self._lbl("Reference Number (Optional)"))
        self.in_ref = QLineEdit(ref or "")
        self.in_ref.setStyleSheet(INPUT_STYLE)
        self.in_ref.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_ref)

        lay.addWidget(self._lbl("Linked Bill"))
        self.combo_bill = QComboBox()
        self.combo_bill.setStyleSheet(INPUT_STYLE)
        self.combo_bill.setFixedHeight(INPUT_H)
        self.combo_bill.addItem("(Not linked — auto-apply to oldest unpaid)", None)
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT id, billing_month, total_payable, amount_paid,
                              balance_remaining, settled_status
                       FROM tenant_bills
                       WHERE tenant_id=? ORDER BY bill_date ASC, id ASC""", (tid,))
        for bid, bmonth, btotal, bpaid, brem, bstat in cur.fetchall():
            tag = ""
            if bstat == "closed_manual": tag = " · CLOSED"
            elif brem <= 0.001: tag = " · PAID"
            elif (bpaid or 0) > 0.001: tag = f" · PARTIAL (due {fmt_money(brem)})"
            else: tag = f" · UNPAID (due {fmt_money(btotal)})"
            self.combo_bill.addItem(f"{bmonth}{tag}", bid)
        conn.close()
        if linked_bill:
            for i in range(self.combo_bill.count()):
                if self.combo_bill.itemData(i) == linked_bill:
                    self.combo_bill.setCurrentIndex(i); break
        lay.addWidget(self.combo_bill)

        lay.addWidget(self._lbl("Notes"))
        self.in_notes = QTextEdit()
        self.in_notes.setStyleSheet(INPUT_STYLE)
        self.in_notes.setMaximumHeight(56)
        if notes: self.in_notes.setPlainText(notes)
        lay.addWidget(self.in_notes)

        warn = QLabel(
            "Saving will automatically rebuild this tenant's Master Excel ledger and "
            "regenerate every bill PDF impacted by this adjustment."
        )
        warn.setStyleSheet(f"""
            font-size: 11px; color: {ACCENT_AMBER};
            background: rgba(245,158,11,0.08);
            border: 1px solid rgba(245,158,11,0.30);
            border-radius: 10px; padding: 8px 12px;
            font-family: {FONT_STACK};
        """)
        warn.setWordWrap(True)
        lay.addWidget(warn)

        lay.addSpacing(4)
        row = QHBoxLayout()
        row.setSpacing(10)
        delete_btn = QPushButton("DELETE")
        delete_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
        delete_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        delete_btn.setFixedHeight(BTN_PRIMARY)
        delete_btn.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_PRIMARY))
        delete_btn.clicked.connect(self.delete_txn)

        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("SAVE CHANGES")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save.clicked.connect(self.save_changes)

        row.addWidget(delete_btn, 1)
        row.addWidget(cancel, 1)
        row.addWidget(save, 2)
        lay.addLayout(row)
        base.addWidget(env)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def save_changes(self):
        amt_raw = self.in_amount.text().strip()
        try:
            amt = float(amt_raw)
        except ValueError:
            QMessageBox.warning(self, "Invalid", "Amount must be numeric."); return
        if amt <= 0:
            QMessageBox.warning(self, "Invalid", "Amount must be > 0."); return

        self.in_date.interpretText()
        qdate = self.in_date.date()
        raw_text = self.in_date.text().strip()
        for fmt in ["dd/MMM/yyyy", "dd-MMM-yyyy", "dd/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd", "d/M/yyyy", "d-M-yyyy"]:
            parsed = QDate.fromString(raw_text, fmt)
            if parsed.isValid():
                qdate = parsed
                break
        new_date = qdate.toString("yyyy-MM-dd")
        new_method = self.combo_method.currentText()
        new_ref = self.in_ref.text().strip()
        new_notes = self.in_notes.toPlainText().strip()
        new_linked = self.combo_bill.currentData()
        tid = self.txn[0]

        old_values = {
            "amount": self.txn[5], "txn_date": self.txn[2],
            "payment_method": self.txn[6], "reference_no": self.txn[7],
            "linked_bill_id": self.txn[8], "notes": self.txn[9],
        }
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""UPDATE tenant_transactions
                       SET txn_date=?, amount=?, payment_method=?,
                           reference_no=?, linked_bill_id=?, notes=?
                       WHERE id=?""",
                    (new_date, amt, new_method, new_ref,
                     new_linked, new_notes, self.txn_id))
        conn.commit()
        conn.close()

        log_edit(self.user_id, tid, self.txn_id, "edit_payment", old_values,
                 {"amount": amt, "txn_date": new_date, "payment_method": new_method,
                  "reference_no": new_ref, "linked_bill_id": new_linked, "notes": new_notes})

        try:
            recompute_tenant_from_transactions(tid)
        except Exception as e:
            log.error(f"Recompute after edit failed: {e}")
            QMessageBox.warning(self, "Partial Save", f"Transaction updated, but recompute encountered an issue:\n\n{e}")
            self.accept()
            return

        try:
            regenerate_receipt_pdfs_for_txn(self.user_id, self.txn_id)
        except Exception as e:
            log.error(f"Receipt PDF regeneration failed in EditPaymentDialog: {e}")

        QMessageBox.information(self, "Saved", "Payment updated.\n\nLedger Excel and affected bill PDFs were regenerated.")
        self.accept()

    def delete_txn(self):
        if QMessageBox.question(
            self, "Delete Payment",
            "Permanently delete this payment?\n\n"
            "The tenant's Excel ledger and affected bill PDFs will be regenerated.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return
        tid = self.txn[0]
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("DELETE FROM tenant_transactions WHERE id=?", (self.txn_id,))
        conn.commit()
        conn.close()

        log_edit(self.user_id, tid, self.txn_id, "delete_payment",
                 {"amount": self.txn[5], "txn_date": self.txn[2]}, None)
        try:
            delete_receipt_pdfs_for_txn(self.user_id, tid, self.txn_id)
        except Exception as e:
            log.error(f"Receipt deletion error: {e}")

        try:
            recompute_tenant_from_transactions(tid)
        except Exception as e:
            log.error(f"Recompute after delete failed: {e}")
            QMessageBox.warning(self, "Partial Delete", f"Payment deleted, but recompute had an issue:\n\n{e}")
            self.accept()
            return

        QMessageBox.information(self, "Deleted", "Payment deleted and downstream documents rebuilt.")
        self.accept()


class RelinkPaymentDialog(GlassDialog):
    def __init__(self, user_id, txn_id, parent=None):
        super().__init__(parent, width=460)
        self.user_id = user_id
        self.txn_id = txn_id
        self.tid = None
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(26, 22, 26, 22)
        lay.setSpacing(10)

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT tenant_id, linked_bill_id, amount, txn_date
                       FROM tenant_transactions WHERE id=?""", (self.txn_id,))
        row = cur.fetchone()
        if not row:
            conn.close()
            lay.addWidget(QLabel("Transaction not found."))
            base.addWidget(env); return
        self.tid, current_linked, amt, tdate = row

        cur.execute("""SELECT id, billing_month, total_payable, amount_paid,
                              balance_remaining, settled_status
                       FROM tenant_bills WHERE tenant_id=?
                       ORDER BY bill_date ASC, id ASC""", (self.tid,))
        bills = cur.fetchall()
        conn.close()

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("link", color=ACCENT_INDIGO, size=18))
        t = QLabel("RELINK PAYMENT TO BILL")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        info = QLabel(f"Payment of {fmt_money(amt)} on {tdate}")
        info.setStyleSheet(f"font-size: 12px; color: {T.c('text_soft')}; font-family: {FONT_STACK};")
        lay.addWidget(info)

        self.combo = QComboBox()
        self.combo.setStyleSheet(INPUT_STYLE)
        self.combo.setFixedHeight(INPUT_H)
        self.combo.addItem("(Not linked — auto-apply to oldest unpaid)", None)
        for bid, bmonth, btotal, bpaid, brem, bstat in bills:
            tag = ""
            if bstat == "closed_manual": tag = " · CLOSED"
            elif brem <= 0.001: tag = " · PAID"
            elif (bpaid or 0) > 0.001: tag = f" · PARTIAL (due {fmt_money(brem)})"
            else: tag = f" · UNPAID (due {fmt_money(btotal)})"
            self.combo.addItem(f"{bmonth}{tag}", bid)
        if current_linked:
            for i in range(self.combo.count()):
                if self.combo.itemData(i) == current_linked:
                    self.combo.setCurrentIndex(i); break
        lay.addWidget(self.combo)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("SAVE LINK")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save.clicked.connect(self.do_save)

        btn_row.addWidget(cancel, 1)
        btn_row.addWidget(save, 1)
        lay.addLayout(btn_row)
        base.addWidget(env)

    def do_save(self):
        new_linked = self.combo.currentData()
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("UPDATE tenant_transactions SET linked_bill_id=? WHERE id=?",
                    (new_linked, self.txn_id))
        conn.commit()
        conn.close()
        log_edit(self.user_id, self.tid, self.txn_id, "relink_payment",
                 None, {"linked_bill_id": new_linked})
        try:
            recompute_tenant_from_transactions(self.tid)
        except Exception as e:
            log.error(f"Recompute after relink failed: {e}")
        self.accept()


class TenantFilesDialog(GlassDialog):
    def __init__(self, user_id, tenant_info=None, parent=None):
        super().__init__(parent, width=820, height=620)
        self.user_id = user_id
        self.tenant_info = tenant_info
        self.init_ui()
        self.scan_files()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("folder", color=ACCENT_INDIGO, size=20))
        t = QLabel("GENERATED FINANCIAL DOCUMENTS")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        filter_lbl = QLabel(
            f"Filtered scope: <b>{self.tenant_info['name']}</b> ({self.tenant_info['tenant_uid']})"
            if self.tenant_info else
            "Showing all portfolio documents. Double-click a file to launch."
        )
        filter_lbl.setStyleSheet(f"""
            font-size: 11px; color: {T.c('text_soft')};
            background: {T.c('input_bg')};
            border: 1px solid {T.c('input_border')};
            border-radius: 8px; padding: 8px 12px;
            font-family: {FONT_STACK};
        """)
        filter_lbl.setWordWrap(True)
        lay.addWidget(filter_lbl)

        refresh_row = QHBoxLayout()
        refresh_row.addStretch(1)
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setIcon(SVGIconManager.icon("refresh", color="#F8FAFC", size=14))
        refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        refresh_btn.setFixedHeight(BTN_INLINE)
        refresh_btn.setStyleSheet(pill_button_qss("#ffffff", height=BTN_INLINE))
        refresh_btn.clicked.connect(self.scan_files)
        refresh_row.addWidget(refresh_btn)
        lay.addLayout(refresh_row)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["File Name", "Modified", "Size"])
        self.tree.setColumnWidth(0, 480)
        self.tree.setColumnWidth(1, 160)
        self.tree.setColumnWidth(2, 90)
        self.tree.setRootIsDecorated(True)
        self.tree.setStyleSheet(f"""
            QTreeWidget {{
                background-color: {T.c('table_bg')};
                border: 1px solid {T.c('table_border')};
                border-radius: 12px; color: {T.c('text_primary')};
                font-size: 12px; font-family: {FONT_STACK};
                outline: none; padding: 4px;
            }}
            QTreeWidget::item {{ padding: 6px 4px; border: none; }}
            QTreeWidget::item:hover {{ background-color: {T.c('table_sel_bg')}; }}
            QTreeWidget::item:selected {{ background-color: {T.c('table_sel_bg')}; color: {T.c('text_primary')}; }}
            QHeaderView::section {{
                background-color: {T.c('table_hdr_bg')};
                color: {T.c('text_primary')}; border: none;
                padding: 8px 6px; font-weight: 700;
                font-size: 11px; font-family: {FONT_STACK};
            }}
        """)
        self.tree.itemDoubleClicked.connect(self.open_item)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        lay.addWidget(self.tree, 1)

        hint = QLabel("Double-click a file to open. Right-click for options.")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet(f"font-size: 11px; color: {T.c('text_muted')}; font-family: {FONT_STACK};")
        lay.addWidget(hint)
        base.addWidget(env)

    def scan_files(self):
        self.tree.clear()
        root = get_output_root(self.user_id)
        if not os.path.isdir(root):
            self.tree.addTopLevelItem(QTreeWidgetItem(["No output folder configured", "", ""]))
            return

        grouped = {}
        for dirpath, dirnames, filenames in os.walk(root):
            for fname in filenames:
                if not (fname.endswith(".pdf") or fname.endswith(".xlsx") or fname.endswith(".csv")):
                    continue
                full_path = os.path.join(dirpath, fname)
                rel = os.path.relpath(full_path, root)
                parts = rel.split(os.sep)

                location = next((p for p in parts if p.startswith("Location_")), None)
                prop = next((p for p in parts if p.startswith("Property_")), None)
                tenant = next((p for p in parts if p.startswith("Tenant_")), None)

                if location is None and prop is None and tenant is None:
                    grouped.setdefault("Campus-Wide", {}).setdefault(
                        "Shared Assets", {}
                    ).setdefault("Root", {}).setdefault(
                        self._category_for_location_level(fname), []
                    ).append(full_path)
                    continue

                if location and prop is None and tenant is None:
                    grouped.setdefault(location, {}).setdefault(
                        "Shared Assets", {}
                    ).setdefault("Building-Wide", {}).setdefault(
                        self._category_for_location_level(fname), []
                    ).append(full_path)
                    continue

                if self.tenant_info:
                    expected = f"Tenant_{sanitize_fs_name(self.tenant_info['name'] + '_' + self.tenant_info['tenant_uid'])}"
                    if tenant != expected:
                        continue

                loc_key = location or "Ungrouped"
                prop_key = prop or "Ungrouped"
                tenant_key = tenant or "Ungrouped"
                category = self._category_for(parts, fname)
                grouped.setdefault(loc_key, {}).setdefault(prop_key, {}).setdefault(
                    tenant_key, {}
                ).setdefault(category, []).append(full_path)

        if not grouped:
            self.tree.addTopLevelItem(QTreeWidgetItem(["No files generated yet.", "", ""]))
            return

        for loc_name in sorted(grouped.keys()):
            loc_item = QTreeWidgetItem([loc_name, "", ""])
            loc_item.setIcon(0, SVGIconManager.icon("properties", color=ACCENT_INDIGO, size=16))
            loc_item.setExpanded(True)
            self.tree.addTopLevelItem(loc_item)
            for prop_name in sorted(grouped[loc_name].keys()):
                prop_item = QTreeWidgetItem([prop_name, "", ""])
                prop_item.setIcon(0, SVGIconManager.icon("properties", color=ACCENT_INDIGO, size=16))
                prop_item.setExpanded(True)
                loc_item.addChild(prop_item)
                for tenant_name in sorted(grouped[loc_name][prop_name].keys()):
                    ten_item = QTreeWidgetItem([tenant_name, "", ""])
                    ten_item.setIcon(0, SVGIconManager.icon("owner", color=ACCENT_INDIGO, size=16))
                    ten_item.setExpanded(True)
                    prop_item.addChild(ten_item)
                    for cat_name in sorted(grouped[loc_name][prop_name][tenant_name].keys()):
                        files = sorted(grouped[loc_name][prop_name][tenant_name][cat_name])
                        cat_item = QTreeWidgetItem([f"{cat_name} ({len(files)})", "", ""])
                        cat_item.setIcon(0, SVGIconManager.icon("folder", color=ACCENT_AMBER, size=16))
                        cat_item.setExpanded(True)
                        ten_item.addChild(cat_item)
                        for fp in files:
                            try:
                                mtime = datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%d %b %Y · %H:%M")
                                size = self._human_size(os.path.getsize(fp))
                            except OSError:
                                mtime = "—"; size = "—"
                            f_item = QTreeWidgetItem([os.path.basename(fp), mtime, size])
                            f_item.setData(0, Qt.ItemDataRole.UserRole, fp)
                            icon_type = "file_excel" if fp.endswith(".xlsx") else ("file" if fp.endswith(".pdf") else "records")
                            f_item.setIcon(0, SVGIconManager.icon(icon_type, color=ACCENT_INDIGO, size=16))
                            f_item.setForeground(0, QBrush(QColor(ACCENT_INDIGO)))
                            cat_item.addChild(f_item)

    def _category_for(self, parts, fname):
        last_dir = parts[-2] if len(parts) >= 2 else ""
        if "Unpaid" in last_dir: return "Unpaid Bills"
        if "Paid" in last_dir or "Receipt" in last_dir: return "Paid Receipts"
        if "Generated" in last_dir: return "Generated Bill"
        if "Final" in last_dir: return "Final Settlement"
        if fname.startswith("Paid_Bill"): return "Paid Receipts"
        if fname.startswith("Bill_"): return "Generated Bill"
        if fname.startswith("Receipt_"): return "Paid Receipts"
        if fname.startswith("Final_Settlement"): return "Final Settlement"
        if fname == "Tenant_Master_Ledger.xlsx": return "Ledgers · Tenant"
        if fname.endswith(".csv"): return "Exports · CSV"
        if fname.startswith("Statement_"): return "Statements"
        if fname.startswith("Monthly_Summary"): return "Monthly Summaries"
        if fname.endswith(".xlsx"): return "Excel"
        return "Other"

    def _category_for_location_level(self, fname):
        if fname == "Master_Water_Ledger.xlsx": return "Master Water Ledger"
        if fname.endswith(".csv"): return "CSV Exports"
        if fname.startswith("Monthly_Summary"): return "Monthly Summaries"
        if fname.endswith(".xlsx"): return "Excel"
        if fname.endswith(".pdf"): return "PDF"
        return "Other"

    def _human_size(self, n):
        for unit in ["B", "KB", "MB", "GB"]:
            if n < 1024:
                return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} TB"

    def open_item(self, item, column):
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if not path or not os.path.exists(path): return
        try:
            if sys.platform.startswith("win"): os.startfile(path)
            elif sys.platform == "darwin": subprocess.run(["open", path])
            else: subprocess.run(["xdg-open", path])
        except Exception as e:
            QMessageBox.warning(self, "Could not open", f"{e}\n\nPath:\n{path}")

    def show_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item: return
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if not path: return
        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background-color: {T.c('menu_bg')}; color: {T.c('text_primary')};
                border: 1px solid {T.c('menu_border')};
                border-radius: 8px; padding: 6px;
                font-family: {FONT_STACK}; font-size: 12px;
            }}
            QMenu::item {{ padding: 6px 20px; border-radius: 6px; }}
            QMenu::item:selected {{ background-color: {T.c('menu_sel_bg')}; }}
        """)
        a_open = menu.addAction("Open File")
        a_folder = menu.addAction("Show in Folder")
        a_copy = menu.addAction("Copy Path")
        menu.addSeparator()
        a_del = menu.addAction("Delete File")
        action = menu.exec(self.tree.viewport().mapToGlobal(pos))
        if action == a_open: self.open_item(item, 0)
        elif action == a_folder: self.show_in_folder(path)
        elif action == a_copy: QApplication.clipboard().setText(path)
        elif action == a_del: self._delete_file(path)

    def show_in_folder(self, path):
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(path)])
        except Exception as e:
            QMessageBox.warning(self, "Could not open folder", str(e))

    def _delete_file(self, path):
        if QMessageBox.question(self, "Delete File",
                                f"Delete this file?\n\n{os.path.basename(path)}",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            self.scan_files()
        except Exception as e:
            QMessageBox.warning(self, "Delete failed", str(e))


class PaymentDialog(GlassDialog):
    def __init__(self, user_id, tenant_info, parent=None, is_advance=False):
        super().__init__(parent, width=480)
        self.user_id = user_id
        self.info = tenant_info
        self.is_advance = is_advance
        self.unpaid_bills = [] if is_advance else fetch_unpaid_bills(tenant_info["id"])
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("dollar", color=ACCENT_GREEN if not self.is_advance else ACCENT_INDIGO, size=20))
        t = QLabel("RECORD ADVANCE" if self.is_advance else "RECORD PAYMENT")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        chip = QLabel(f"<b>{self.info['name']}</b> — {self.info['property_name']}")
        chip.setStyleSheet(f"""
            background: rgba(16,185,129,0.12); color: {T.c('text_primary')};
            border: 1px solid rgba(16,185,129,0.30);
            border-radius: {RADIUS_INPUT}px; padding: 8px 12px; font-size: 12px;
            font-family: {FONT_STACK};
        """)
        lay.addWidget(chip)

        cur_bal = get_tenant_balance(self.info["id"])
        color = ACCENT_RED if cur_bal > 0 else ACCENT_GREEN
        self.lbl_bal = QLabel(f"<b>Current Balance:</b> {fmt_money(cur_bal)}")
        self.lbl_bal.setStyleSheet(f"color: {color}; font-size: 12px; padding: 2px 2px; font-family: {FONT_STACK};")
        lay.addWidget(self.lbl_bal)

        lay.addWidget(self._lbl("Amount (₹) *"))
        self.in_amount = QLineEdit()
        self.in_amount.setPlaceholderText("0.00")
        self.in_amount.setStyleSheet(INPUT_STYLE)
        self.in_amount.setFixedHeight(INPUT_H)
        self.in_amount.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        lay.addWidget(self.in_amount)

        lay.addWidget(self._lbl("Paid Date *"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setStyleSheet(INPUT_STYLE)
        self.in_date.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_date)

        lay.addWidget(self._lbl("Payment Method"))
        self.combo_method = QComboBox()
        self.combo_method.addItems(["Cash", "UPI", "Bank Transfer", "Cheque", "Card", "Other"])
        self.combo_method.setStyleSheet(INPUT_STYLE)
        self.combo_method.setFixedHeight(INPUT_H)
        lay.addWidget(self.combo_method)

        lay.addWidget(self._lbl("Reference Number (Optional)"))
        self.in_ref = QLineEdit()
        self.in_ref.setPlaceholderText("UPI ref / cheque no / txn id")
        self.in_ref.setStyleSheet(INPUT_STYLE)
        self.in_ref.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_ref)

        lay.addWidget(self._lbl("Link to Bill"))
        self.combo_bill = QComboBox()
        self.combo_bill.setStyleSheet(INPUT_STYLE)
        self.combo_bill.setFixedHeight(INPUT_H)
        if self.is_advance:
            self.combo_bill.addItem("(Advance — no linked bill)", None)
            self.combo_bill.setEnabled(False)
        else:
            self.combo_bill.addItem("(General credit — auto-apply to oldest unpaid)", None)
            for b in self.unpaid_bills:
                bid, month, bdate, total, paid, rem = b
                self.combo_bill.addItem(f"{month} — Due {fmt_money(rem)}", bid)
        lay.addWidget(self.combo_bill)

        lay.addWidget(self._lbl("Notes (Optional)"))
        self.in_notes = QTextEdit()
        self.in_notes.setPlaceholderText("Optional note...")
        self.in_notes.setStyleSheet(INPUT_STYLE)
        self.in_notes.setMaximumHeight(56)
        lay.addWidget(self.in_notes)

        lay.addSpacing(6)
        row = QHBoxLayout()
        row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("RECORD ADVANCE" if self.is_advance else "RECORD PAYMENT")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        save.clicked.connect(self.save_payment)

        row.addWidget(cancel, 1)
        row.addWidget(save, 1)
        lay.addLayout(row)
        base.addWidget(env)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def save_payment(self):
        amt_raw = self.in_amount.text().strip()
        if not amt_raw:
            QMessageBox.warning(self, "Amount Required", "Please enter the amount."); return
        try: amt = float(amt_raw)
        except ValueError:
            QMessageBox.warning(self, "Invalid", "Amount must be a valid number."); return
        if amt <= 0:
            QMessageBox.warning(self, "Invalid", "Amount must be greater than zero."); return

        self.in_date.interpretText()
        qdate = self.in_date.date()
        raw_text = self.in_date.text().strip()
        for fmt in ["dd/MMM/yyyy", "dd-MMM-yyyy", "dd/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd", "d/M/yyyy", "d-M-yyyy"]:
            parsed = QDate.fromString(raw_text, fmt)
            if parsed.isValid():
                qdate = parsed
                break
        pdate = qdate.toString("yyyy-MM-dd")
        method = self.combo_method.currentText()
        ref = self.in_ref.text().strip()
        notes = self.in_notes.toPlainText().strip()
        linked_bill = self.combo_bill.currentData()
        txn_type = "advance" if self.is_advance else "payment"

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""INSERT INTO tenant_transactions
            (user_id, tenant_id, property_id, txn_date, txn_type, description,
             amount, payment_method, reference_no, linked_bill_id, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (self.user_id, self.info["id"], self.info["property_id"], pdate,
             txn_type, "Advance payment" if self.is_advance else "Payment",
             amt, method, ref, linked_bill, notes))
        txn_id = cur.lastrowid
        conn.commit()
        conn.close()

        try:
            recompute_tenant_from_transactions(self.info["id"])
        except Exception as e:
            log.error(f"Recompute after payment failed: {e}")

        try:
            addr = self.info["address"]
            y = self.in_date.date().year()
            b_month_target = self.in_date.date().toString("MMM yyyy")
            if linked_bill:
                conn_b = sqlite3.connect(get_db_path())
                cur_b = conn_b.cursor()
                cur_b.execute("SELECT billing_month, bill_date FROM tenant_bills WHERE id=?", (linked_bill,))
                brow = cur_b.fetchone()
                conn_b.close()
                if brow and brow[0]:
                    b_month_target = brow[0]
                    if brow[1]:
                        try: y = int(brow[1].split("-")[0])
                        except Exception: pass
            gen = get_bill_dir(self.user_id, addr, self.info["property_name"],
                               self.info["name"], self.info["tenant_uid"],
                               y, b_month_target, "Receipts")

            owner_name = "Property Administrator"; owner_phone = "--"
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("SELECT owner_name, phone FROM profiles WHERE user_id=?", (self.user_id,))
            row = cur.fetchone()
            if row:
                owner_name = row[0] or owner_name
                owner_phone = row[1] or owner_phone
            conn.close()

            bal_after = get_tenant_balance(self.info["id"])
            payload = {
                "txn_id": txn_id,
                "owner_name": owner_name, "owner_phone": owner_phone,
                "property_name": self.info["property_name"], "property_address": addr,
                "tenant_name": self.info["name"], "tenant_uid": self.info["tenant_uid"],
                "amount": amt, "payment_date": pdate, "paid_date": pdate,
                "payment_method": method, "reference_no": ref,
                "description": "Advance payment" if self.is_advance else "Payment",
                "balance_after": bal_after,
            }
            if linked_bill:
                conn = sqlite3.connect(get_db_path())
                cur = conn.cursor()
                cur.execute("SELECT billing_month, total_payable, balance_remaining FROM tenant_bills WHERE id=?", (linked_bill,))
                bb = cur.fetchone()
                conn.close()
                if bb:
                    payload["linked_bill_month"] = bb[0]
                    payload["linked_bill_amount"] = bb[1]
                    payload["linked_bill_remaining"] = bb[2]
            generate_receipt_pdfs(gen, payload)
        except Exception as e:
            log.error(f"Receipt PDF failed: {e}")

        log_edit(self.user_id, self.info["id"], txn_id,
                 "record_advance" if self.is_advance else "record_payment",
                 None, {"amount": amt, "linked_bill_id": linked_bill})

        self.accept()


class AdjustmentDialog(GlassDialog):
    def __init__(self, user_id, tenant_info, parent=None):
        super().__init__(parent, width=460)
        self.user_id = user_id
        self.info = tenant_info
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("settings", color=ACCENT_AMBER, size=18))
        t = QLabel("RECORD ADJUSTMENT")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        info_lbl = QLabel("Adjustments correct the running balance WITHOUT generating a bill or invoice.")
        info_lbl.setStyleSheet(f"""
            font-size: 11px; color: {T.c('text_soft')};
            background: rgba(245,158,11,0.08);
            border: 1px solid rgba(245,158,11,0.30);
            border-radius: 10px; padding: 8px 12px;
            font-family: {FONT_STACK};
        """)
        info_lbl.setWordWrap(True)
        lay.addWidget(info_lbl)

        chip = QLabel(f"<b>{self.info['name']}</b> — {self.info['property_name']}")
        chip.setStyleSheet(f"""
            background: rgba(245,158,11,0.10); color: {T.c('text_primary')};
            border: 1px solid rgba(245,158,11,0.30);
            border-radius: {RADIUS_INPUT}px; padding: 8px 12px; font-size: 12px;
            font-family: {FONT_STACK};
        """)
        lay.addWidget(chip)

        lay.addWidget(self._lbl("Adjustment Type *"))
        self.combo_type = QComboBox()
        self.combo_type.addItems(["Debit (increase balance — tenant owes more)",
                                  "Credit (reduce balance — tenant owes less)"])
        self.combo_type.setStyleSheet(INPUT_STYLE)
        self.combo_type.setFixedHeight(INPUT_H)
        lay.addWidget(self.combo_type)

        lay.addWidget(self._lbl("Amount (₹) *"))
        self.in_amount = QLineEdit()
        self.in_amount.setPlaceholderText("0.00")
        self.in_amount.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
        self.in_amount.setStyleSheet(INPUT_STYLE)
        self.in_amount.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_amount)

        lay.addWidget(self._lbl("Date *"))
        self.in_date = QDateEdit(QDate.currentDate())
        self.in_date.setCalendarPopup(True)
        setup_calendar_widget(self.in_date.calendarWidget())
        fix_stepper_margins(self.in_date)
        self.in_date.setDisplayFormat("dd/MMM/yyyy")
        self.in_date.setStyleSheet(INPUT_STYLE)
        self.in_date.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_date)

        lay.addWidget(self._lbl("Description *"))
        self.in_desc = QLineEdit()
        self.in_desc.setPlaceholderText("e.g. Round-off adjustment, meter calibration credit")
        self.in_desc.setStyleSheet(INPUT_STYLE)
        self.in_desc.setFixedHeight(INPUT_H)
        lay.addWidget(self.in_desc)

        lay.addSpacing(6)
        row = QHBoxLayout()
        row.setSpacing(10)
        cancel = QPushButton("CANCEL")
        cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel.setFixedHeight(BTN_PRIMARY)
        cancel.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
        cancel.clicked.connect(self.reject)

        save = QPushButton("SAVE ADJUSTMENT")
        save.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
        save.setCursor(Qt.CursorShape.PointingHandCursor)
        save.setFixedHeight(BTN_PRIMARY)
        save.setStyleSheet(primary_button_qss(ACCENT_AMBER, height=BTN_PRIMARY))
        save.clicked.connect(self.save_adjustment)

        row.addWidget(cancel, 1)
        row.addWidget(save, 1)
        lay.addLayout(row)
        base.addWidget(env)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def save_adjustment(self):
        amt_raw = self.in_amount.text().strip()
        desc = self.in_desc.text().strip()
        if not amt_raw or not desc:
            QMessageBox.warning(self, "Required", "Amount and description are required."); return
        try: amt = float(amt_raw)
        except ValueError:
            QMessageBox.warning(self, "Invalid", "Amount must be numeric."); return
        if amt <= 0:
            QMessageBox.warning(self, "Invalid", "Amount must be > 0."); return
        is_debit = self.combo_type.currentIndex() == 0
        ttype = "other_debit" if is_debit else "other_credit"
        tdate = self.in_date.date().toString("yyyy-MM-dd")
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""INSERT INTO tenant_transactions
            (user_id, tenant_id, property_id, txn_date, txn_type, description, amount, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (self.user_id, self.info["id"], self.info["property_id"],
             tdate, ttype, desc, amt, desc))
        conn.commit()
        conn.close()
        self.accept()


class SettleTenantWizard(GlassDialog):
    def __init__(self, user_id, tenant_info, parent=None):
        super().__init__(parent, width=540)
        self.user_id = user_id
        self.info = tenant_info
        self.step = 1
        self.damages = 0.0
        self.damage_desc = ""
        self.outstanding = get_tenant_balance(tenant_info["id"])
        self.deposit = float(tenant_info.get("deposit", 0.0))
        self.init_ui()
        self.render_step()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        self.env = GlassModal()
        self.lay = QVBoxLayout(self.env)
        self.lay.setContentsMargins(28, 24, 28, 24)
        self.lay.setSpacing(12)
        base.addWidget(self.env)

    def _clear(self):
        while self.lay.count():
            it = self.lay.takeAt(0)
            if it.widget(): it.widget().deleteLater()
            elif it.layout(): self._clear_layout(it.layout())

    def _clear_layout(self, layout):
        while layout.count():
            it = layout.takeAt(0)
            if it.widget(): it.widget().deleteLater()
            elif it.layout(): self._clear_layout(it.layout())

    def _header(self, step, title):
        head = QHBoxLayout()
        badge = GlassBadge(f"STEP {step} / 3", "warning")
        t = QLabel(title)
        t.setStyleSheet(f"font-size: 14px; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        head.addWidget(badge)
        head.addSpacing(10)
        head.addWidget(t)
        head.addStretch(1)
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.reject)
        head.addWidget(x)
        return head

    def render_step(self):
        self._clear()
        if self.step == 1:
            self.lay.addLayout(self._header(1, "Review Outstanding Dues"))
            intro = QLabel(f"Closing lease for <b>{self.info['name']}</b> ({self.info['tenant_uid']}) at <b>{self.info['property_name']}</b>.")
            intro.setStyleSheet(f"font-size: 12px; color: {T.c('text_soft')}; font-family: {FONT_STACK};")
            intro.setWordWrap(True)
            self.lay.addWidget(intro)

            card = GlassCard()
            cl = QVBoxLayout(card)
            cl.setContentsMargins(20, 16, 20, 16)
            cl.setSpacing(8)
            row = QHBoxLayout()
            l1 = QLabel("Current Running Balance:")
            l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_dim')}; font-family: {FONT_STACK};")
            l2 = QLabel(fmt_money(self.outstanding))
            l2.setStyleSheet(f"font-size: 18px; font-weight: 800; color: {ACCENT_RED if self.outstanding > 0 else ACCENT_GREEN}; font-family: {FONT_STACK};")
            row.addWidget(l1)
            row.addStretch(1)
            row.addWidget(l2)
            cl.addLayout(row)
            self.lay.addWidget(card)

        elif self.step == 2:
            self.lay.addLayout(self._header(2, "Damage & Repair Deductions"))
            self.lay.addWidget(self._lbl("Damage Charge (₹)"))
            self.in_dmg = QLineEdit()
            self.in_dmg.setPlaceholderText("0.00")
            self.in_dmg.setValidator(QDoubleValidator(0.0, 1e10, 2, self))
            self.in_dmg.setStyleSheet(INPUT_STYLE)
            self.in_dmg.setFixedHeight(INPUT_H)
            if self.damages: self.in_dmg.setText(f"{self.damages:.2f}")
            self.lay.addWidget(self.in_dmg)

            self.lay.addWidget(self._lbl("Damage Itemization & Description"))
            self.in_dmg_desc = QTextEdit()
            self.in_dmg_desc.setPlaceholderText("e.g. Broken fixtures, repainting expenses...")
            self.in_dmg_desc.setStyleSheet(INPUT_STYLE)
            self.in_dmg_desc.setMaximumHeight(70)
            if self.damage_desc: self.in_dmg_desc.setPlainText(self.damage_desc)
            self.lay.addWidget(self.in_dmg_desc)

        else:
            self.lay.addLayout(self._header(3, "Final Settlement Reconciliation"))
            refund = max(0.0, self.deposit - self.outstanding - self.damages)
            shortfall = max(0.0, (self.outstanding + self.damages) - self.deposit)

            card = GlassCard()
            cl = QVBoxLayout(card)
            cl.setContentsMargins(20, 18, 20, 18)
            cl.setSpacing(8)

            def add_summary_row(label, value, color=None):
                if color is None: color = T.c("text_primary")
                r = QHBoxLayout()
                l1 = QLabel(label)
                l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_dim')}; font-family: {FONT_STACK};")
                l2 = QLabel(value)
                l2.setStyleSheet(f"font-size: 13px; font-weight: 800; color: {color}; font-family: {FONT_STACK};")
                r.addWidget(l1)
                r.addStretch(1)
                r.addWidget(l2)
                cl.addLayout(r)

            add_summary_row("Security Deposit Held", fmt_money(self.deposit), ACCENT_INDIGO)
            add_summary_row("Minus: Outstanding dues", f"- {fmt_money(self.outstanding)}", ACCENT_RED)
            add_summary_row("Minus: Damage charges", f"- {fmt_money(self.damages)}", ACCENT_RED)

            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.HLine)
            sep.setStyleSheet(f"border: none; background-color: {T.c('divider')}; height: 1px;")
            cl.addWidget(sep)

            if shortfall > 0:
                add_summary_row("SHORTFALL — Tenant Owes", fmt_money(shortfall), ACCENT_RED)
            else:
                add_summary_row("NET REFUND TO TENANT", fmt_money(refund), ACCENT_GREEN)
            self.lay.addWidget(card)

        self.lay.addSpacing(4)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        if self.step > 1:
            back = QPushButton("BACK")
            back.setCursor(Qt.CursorShape.PointingHandCursor)
            back.setFixedHeight(BTN_PRIMARY)
            back.setStyleSheet(pill_button_qss("#ffffff", height=BTN_PRIMARY))
            back.clicked.connect(self.go_back)
            btn_row.addWidget(back)
        btn_row.addStretch(1)

        if self.step < 3:
            nxt = QPushButton("NEXT STEP")
            nxt.setCursor(Qt.CursorShape.PointingHandCursor)
            nxt.setFixedHeight(BTN_PRIMARY)
            nxt.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
            nxt.clicked.connect(self.go_next)
            btn_row.addWidget(nxt)
        else:
            confirm = QPushButton("CONFIRM SETTLEMENT")
            confirm.setIcon(SVGIconManager.icon("check", color="#04150E", size=16))
            confirm.setCursor(Qt.CursorShape.PointingHandCursor)
            confirm.setFixedHeight(BTN_PRIMARY)
            confirm.setStyleSheet(primary_button_qss(ACCENT_GREEN, height=BTN_PRIMARY))
            confirm.clicked.connect(self.finalize)
            btn_row.addWidget(confirm)
        self.lay.addLayout(btn_row)

    def _lbl(self, text):
        l = QLabel(text)
        l.setStyleSheet(f"font-size: 11px; color: {T.c('text_dim')}; font-weight: 600; letter-spacing: 0.4px; background: transparent; font-family: {FONT_STACK};")
        return l

    def go_next(self):
        if self.step == 2:
            try: self.damages = float(self.in_dmg.text() or 0)
            except ValueError: self.damages = 0.0
            self.damage_desc = self.in_dmg_desc.toPlainText().strip()
        self.step += 1
        self.render_step()

    def go_back(self):
        self.step -= 1
        self.render_step()

    def finalize(self):
        today = QDate.currentDate().toString("yyyy-MM-dd")
        pid = self.info["property_id"]
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()

        if self.damages > 0:
            cur.execute("""INSERT INTO tenant_transactions
                (user_id, tenant_id, property_id, txn_date, txn_type, description, amount, affects_deposit, notes)
                VALUES (?, ?, ?, ?, 'damage_charge', ?, ?, 1, ?)""",
                (self.user_id, self.info["id"], pid, today,
                 f"Damage charge: {self.damage_desc or 'Not specified'}", self.damages, self.damage_desc))

        total_owed = self.outstanding + self.damages
        deposit_applied = min(self.deposit, max(0.0, total_owed))
        if deposit_applied > 0:
            cur.execute("""INSERT INTO tenant_transactions
                (user_id, tenant_id, property_id, txn_date, txn_type, description, amount, affects_deposit, notes)
                VALUES (?, ?, ?, ?, 'deposit_adjust', ?, ?, 1, ?)""",
                (self.user_id, self.info["id"], pid, today,
                 f"Deposit applied against dues ({fmt_money(deposit_applied)})",
                 deposit_applied, "Settle: deposit coverage"))

        refund = max(0.0, self.deposit - total_owed)
        if refund > 0:
            cur.execute("""INSERT INTO tenant_transactions
                (user_id, tenant_id, property_id, txn_date, txn_type, description, amount, affects_deposit, notes)
                VALUES (?, ?, ?, ?, 'deposit_refund', ?, ?, 1, ?)""",
                (self.user_id, self.info["id"], pid, today,
                 f"Deposit refunded to tenant ({fmt_money(refund)})",
                 refund, "Settle: cash refund (no balance effect)"))

        cur.execute("SELECT id, balance_remaining, settled_status FROM tenant_bills WHERE tenant_id=? AND balance_remaining > 0.001", (self.info["id"],))
        for bid, rem, st in cur.fetchall():
            if st == "closed_manual": continue
            new_status = "settled_deposit" if self.deposit >= total_owed else "open"
            cur.execute("UPDATE tenant_bills SET settled_status=? WHERE id=?", (new_status, bid))
        cur.execute("UPDATE tenants SET is_active=0, exit_date=? WHERE id=?", (today, self.info["id"]))
        cur.execute("SELECT owner_name, phone FROM profiles WHERE user_id=?", (self.user_id,))
        pr = cur.fetchone()
        owner_name = pr[0] if pr else "Property Owner"
        owner_phone = pr[1] if pr else "—"
        conn.commit()
        conn.close()

        try:
            rebuild_tenant_ledger(self.user_id, self.info["id"])
        except Exception as e:
            log.error(f"Ledger rebuild after settle failed: {e}")

        try:
            record_tenant_exit_in_water_ledger(self.user_id, self.info["name"], today)
        except Exception as e:
            log.error(f"Water ledger exit record failed: {e}")

        try:
            year_str = today[:4]
            month_folder = f"{today[5:7]}_Exit_{today}"
            gen_dir = get_bill_dir(self.user_id, self.info["address"], self.info["property_name"],
                                   self.info["name"], self.info["tenant_uid"],
                                   year_str, month_folder, "Final")
            final_payload = {
                "owner_name": owner_name, "owner_phone": owner_phone,
                "property_name": self.info["property_name"], "property_address": self.info["address"],
                "tenant_name": self.info["name"], "tenant_uid": self.info["tenant_uid"],
                "bill_date": today,
                "prev_due": self.outstanding, "damages": self.damages, "deposit": self.deposit,
                "remarks": f"Damage: {self.damage_desc or 'None'} | Deposit: {fmt_money(self.deposit)} | Refund: {fmt_money(refund)}"
            }
            generate_final_settlement_pdfs(gen_dir, final_payload)
        except Exception as e:
            log.error(f"Final settlement PDF failed: {e}")

        QMessageBox.information(self, "Tenant Settled",
                                f"{self.info['name']} tenancy has been finalized.\n\n"
                                f"Deposit: {fmt_money(self.deposit)}\n"
                                f"Outstanding: {fmt_money(self.outstanding)}\n"
                                f"Damages: {fmt_money(self.damages)}\n"
                                f"Refund: {fmt_money(refund)}")
        self.accept()


class DepositsReportDialog(GlassDialog):
    def __init__(self, user_id, parent=None):
        super().__init__(parent, width=720, height=520)
        self.user_id = user_id
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(26, 22, 26, 22)
        lay.setSpacing(10)

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("dollar", color=ACCENT_AMBER, size=20))
        t = QLabel("SECURITY DEPOSITS HELD")
        t.setStyleSheet(f"font-size: 13px; font-weight: 800; letter-spacing: 1.2px; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Tenant", "Property", "Deposit Held", "Adjusted", "Net Held"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setStyleSheet(table_qss())
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        lay.addWidget(self.table)

        self.lbl_total = QLabel("Total net held: ₹0.00")
        self.lbl_total.setStyleSheet(f"color: {ACCENT_INDIGO}; font-size: 13px; font-weight: 700; font-family: {FONT_STACK};")
        lay.addWidget(self.lbl_total)
        base.addWidget(env)
        self.load_data()

    def load_data(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.id, t.name, p.name, t.security_deposit, t.is_active
                       FROM tenants t JOIN properties p ON t.property_id=p.id
                       WHERE t.user_id=? ORDER BY t.is_active DESC, t.name ASC""", (self.user_id,))
        tenants = cur.fetchall()
        rows = []; total_net = 0.0
        for tid, tname, pname, dep, active in tenants:
            cur.execute("SELECT COALESCE(SUM(amount),0) FROM tenant_transactions WHERE tenant_id=? AND txn_type='damage_charge'", (tid,))
            dmg = cur.fetchone()[0] or 0.0
            cur.execute("SELECT COALESCE(SUM(amount),0) FROM tenant_transactions WHERE tenant_id=? AND txn_type='deposit_adjust'", (tid,))
            adj = cur.fetchone()[0] or 0.0
            cur.execute("SELECT COALESCE(SUM(amount),0) FROM tenant_transactions WHERE tenant_id=? AND txn_type='deposit_refund'", (tid,))
            refund = cur.fetchone()[0] or 0.0
            net = float(dep or 0) - dmg - adj - refund
            rows.append((tname, pname, float(dep or 0), dmg + adj + refund, net, active))
            total_net += net
        conn.close()
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            vals = [r[0] + (" (Active)" if r[5] else " (Closed)"), r[1], fmt_money(r[2]), fmt_money(r[3]), fmt_money(r[4])]
            for j, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                if j >= 2: it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if j == 4: it.setForeground(QBrush(QColor(ACCENT_INDIGO)))
                self.table.setItem(i, j, it)
        self.lbl_total.setText(f"Total net held across portfolio: {fmt_money(total_net)}")


class BillDetailDialog(GlassDialog):
    def __init__(self, user_id, bill_id, parent=None):
        super().__init__(parent, width=600, height=560)
        self.user_id = user_id
        self.bill_id = bill_id
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT b.billing_month, b.bill_date, b.base_rent, b.elec_total, b.water_total,
                       b.discount_total, b.prev_due, b.total_payable, b.amount_paid,
                       b.balance_remaining, b.settled_status, b.remarks,
                       t.name, t.tenant_uid, p.name
                       FROM tenant_bills b JOIN tenants t ON b.tenant_id = t.id
                       JOIN properties p ON t.property_id = p.id
                       WHERE b.id=?""", (self.bill_id,))
        r = cur.fetchone()
        conn.close()
        if not r:
            lay.addWidget(QLabel("Bill record not found"))
            base.addWidget(env)
            return

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("file", color=ACCENT_INDIGO, size=20))
        t = QLabel(f"Bill Statement — {r[0]}")
        t.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        content = QWidget()
        content.setStyleSheet("background: transparent;")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(0,0,6,0)
        cl.setSpacing(4)
        sections = [
            ("BILL OVERVIEW", [("Tenant", f"{r[12]} ({r[13]})"), ("Property", r[14]),
                               ("Billing Cycle", r[0]), ("Bill Date", r[1]),
                               ("Status", (r[10] or "unpaid").replace("_", " ").title())]),
            ("LINE CHARGES", [("Base Rent", fmt_money(r[2])), ("Electricity", fmt_money(r[3])),
                             ("Water", fmt_money(r[4])), ("Discount", f"- {fmt_money(r[5])}"),
                             ("Previous Balance", fmt_money(r[6]))]),
            ("RECONCILIATION", [("Total Payable", fmt_money(r[7])), ("Amount Paid", fmt_money(r[8])),
                                ("Balance Remaining", fmt_money(r[9])), ("Remarks", r[11] or "—")]),
        ]
        for i, (sec, rows) in enumerate(sections):
            if i > 0:
                divider = QFrame()
                divider.setFrameShape(QFrame.Shape.HLine)
                divider.setStyleSheet(f"border: none; background-color: {T.c('divider')}; height: 1px;")
                cl.addWidget(divider)
            cl.addWidget(self._section(sec, rows))
        scroll.setWidget(content)
        lay.addWidget(scroll, 1)

        actions = QHBoxLayout()
        actions.addStretch(1)

        edit_btn = QPushButton("Edit Bill")
        edit_btn.setIcon(SVGIconManager.icon("edit", color="#F8FAFC", size=14))
        edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        edit_btn.setFixedHeight(BTN_SECONDARY)
        edit_btn.setStyleSheet(pill_button_qss("#ffffff", height=BTN_SECONDARY))
        edit_btn.clicked.connect(self.open_edit)
        actions.addWidget(edit_btn)

        del_btn = QPushButton("Delete Bill")
        del_btn.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=14))
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.setFixedHeight(BTN_SECONDARY)
        del_btn.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_SECONDARY))
        del_btn.clicked.connect(self.delete_bill)
        actions.addWidget(del_btn)

        close_btn = QPushButton("Mark Bill as Closed")
        close_btn.setIcon(SVGIconManager.icon("check", color=ACCENT_AMBER, size=14))
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setFixedHeight(BTN_SECONDARY)
        close_btn.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_SECONDARY))
        close_btn.clicked.connect(self.mark_closed)
        if (r[9] is not None and r[9] <= 0.001) or (r[10] in ("closed", "closed_manual", "settled", "settled_deposit")):
            close_btn.setEnabled(False)
            close_btn.setText("Bill Closed")
            close_btn.setStyleSheet(pill_button_qss(T.c("text_muted"), height=BTN_SECONDARY))
        actions.addWidget(close_btn)
        lay.addLayout(actions)
        base.addWidget(env)

    def _section(self, title, rows):
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 6, 0, 6)
        v.setSpacing(3)
        t = QLabel(title)
        t.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {ACCENT_INDIGO}; letter-spacing: 1.0px; background: transparent; padding: 0 0 6px 0; font-family: {FONT_STACK};")
        v.addWidget(t)
        for label, value in rows:
            row = QHBoxLayout()
            row.setSpacing(14)
            l1 = QLabel(label)
            l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            l1.setFixedWidth(200)
            l2 = QLabel(str(value))
            l2.setStyleSheet(f"font-size: 12px; color: {T.c('text_primary')}; font-weight: 500; background: transparent; font-family: {FONT_STACK};")
            l2.setWordWrap(True)
            row.addWidget(l1)
            row.addWidget(l2, 1)
            v.addLayout(row)
        return w

    def open_edit(self):
        dlg = EditBillDialog(self.user_id, self.bill_id, parent=self)
        if dlg.exec():
            self.accept()

    def delete_bill(self):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT billing_month, total_payable FROM tenant_bills WHERE id=?", (self.bill_id,))
        r = cur.fetchone()
        conn.close()
        if not r: return
        b_month, b_total = r

        if QMessageBox.question(
            self, "Delete Bill",
            f"Permanently delete the bill for {b_month} ({fmt_money(b_total)})?\n\n"
            "This will remove the bill, update the tenant's balance, rebuild the Excel ledger, "
            "and remove all generated bill PDFs.\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return

        ok, msg = delete_bill_record(self.user_id, self.bill_id)
        if ok:
            QMessageBox.information(self, "Deleted", msg)
            self.accept()
        else:
            QMessageBox.warning(self, "Delete Failed", msg)

    def mark_closed(self):
        if QMessageBox.question(self, "Mark Closed", "Waive the remaining balance and mark this bill as closed?",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("UPDATE tenant_bills SET settled_status='closed_manual', balance_remaining=0, payment_status='Paid' WHERE id=?", (self.bill_id,))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Closed", "Bill marked as closed.")
        self.accept()


class TransactionDetailDialog(GlassDialog):
    def __init__(self, user_id, txn_id, parent=None):
        super().__init__(parent, width=560, height=480)
        self.user_id = user_id
        self.txn_id = txn_id
        self.init_ui()

    def init_ui(self):
        base = QVBoxLayout(self)
        base.setContentsMargins(10, 10, 10, 10)
        env = GlassModal()
        lay = QVBoxLayout(env)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(12)

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT tt.txn_date, tt.txn_type, tt.description, tt.amount,
                       tt.payment_method, tt.reference_no, tt.linked_bill_id,
                       tt.affects_deposit, tt.notes, t.name, t.tenant_uid, p.name
                       FROM tenant_transactions tt
                       JOIN tenants t ON tt.tenant_id = t.id
                       JOIN properties p ON tt.property_id = p.id
                       WHERE tt.id=?""", (self.txn_id,))
        r = cur.fetchone()
        conn.close()
        if not r:
            lay.addWidget(QLabel("Transaction not found"))
            base.addWidget(env)
            return

        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(SVGIconManager.pixmap("records", color=ACCENT_INDIGO, size=20))
        t = QLabel(f"Transaction Audit — {(r[1] or '').replace('_',' ').title()}")
        t.setStyleSheet(f"font-size: {FS_TITLE}; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        x = QPushButton()
        x.setIcon(SVGIconManager.icon("close", color=ACCENT_RED, size=14))
        x.setFixedSize(BTN_ICON, BTN_ICON)
        x.setCursor(Qt.CursorShape.PointingHandCursor)
        x.setStyleSheet(round_button_qss(ACCENT_RED, size=BTN_ICON))
        x.clicked.connect(self.accept)
        head.addWidget(ic)
        head.addSpacing(6)
        head.addWidget(t)
        head.addStretch(1)
        head.addWidget(x)
        lay.addLayout(head)

        rows = [
            ("Tenant", f"{r[9]} ({r[10]})"), ("Property", r[11]), ("Date", r[0]),
            ("Type", (r[1] or "").replace("_"," ").title()), ("Amount", fmt_money(r[3])),
        ]
        if r[4]: rows.append(("Payment Method", r[4]))
        if r[5]: rows.append(("Reference", r[5]))
        if r[2]: rows.append(("Description", r[2]))
        if r[6]: rows.append(("Linked Bill ID", f"#{r[6]}"))
        else: rows.append(("Linked Bill", "— (unlinked)"))
        if r[7]: rows.append(("Affects Deposit", "Yes" if r[7] else "No"))
        if r[8]: rows.append(("Audit Notes", r[8]))

        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(6)
        for label, value in rows:
            row = QHBoxLayout()
            row.setSpacing(14)
            l1 = QLabel(label)
            l1.setStyleSheet(f"font-size: 12px; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
            l1.setFixedWidth(180)
            l2 = QLabel(str(value))
            l2.setStyleSheet(f"font-size: 12px; color: {T.c('text_primary')}; font-weight: 500; background: transparent; font-family: {FONT_STACK};")
            l2.setWordWrap(True)
            row.addWidget(l1)
            row.addWidget(l2, 1)
            v.addLayout(row)
        lay.addWidget(w)
        lay.addStretch(1)
        base.addWidget(env)


class RecordsSubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.tenant_cache = {}
        self.current_tid = None
        self.init_ui()

    def init_ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(14)

        top = QHBoxLayout()
        tb = QVBoxLayout()
        tb.setSpacing(2)
        t = QLabel("Records & Financial Ledger")
        t.setStyleSheet(f"font-size: {FS_SUB}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        s = QLabel("Track payments, advances, adjustments, and deposits")
        s.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        tb.addWidget(t)
        tb.addWidget(s)
        top.addLayout(tb)
        top.addStretch(1)

        files_btn = QPushButton("VIEW DOCUMENTS")
        files_btn.setIcon(SVGIconManager.icon("folder", color=ACCENT_INDIGO, size=16))
        files_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        files_btn.setFixedHeight(BTN_SECONDARY)
        files_btn.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_SECONDARY))
        files_btn.clicked.connect(self.open_files)
        top.addWidget(files_btn)

        dep_btn = QPushButton("DEPOSITS HELD")
        dep_btn.setIcon(SVGIconManager.icon("dollar", color=ACCENT_AMBER, size=16))
        dep_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        dep_btn.setFixedHeight(BTN_SECONDARY)
        dep_btn.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_SECONDARY))
        dep_btn.clicked.connect(self.open_deposits)
        top.addWidget(dep_btn)
        v.addLayout(top)

        # 1. Tenant Selector Card
        sel_card = GlassCard()
        sl = QVBoxLayout(sel_card)
        sl.setContentsMargins(20, 16, 20, 16)
        sl.setSpacing(8)
        lbl = QLabel("SELECT TENANT")
        lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.5px; font-family: {FONT_STACK};")
        sl.addWidget(lbl)
        self.combo_tenant = QComboBox()
        self.combo_tenant.setStyleSheet(INPUT_STYLE)
        self.combo_tenant.setFixedHeight(INPUT_H)
        self.combo_tenant.currentIndexChanged.connect(self.on_tenant_selected)
        sl.addWidget(self.combo_tenant)
        v.addWidget(sel_card)

        # 2. KPI Summary Cards
        summary = QHBoxLayout()
        summary.setSpacing(10)
        def mk_card(label, color=None):
            if color is None: color = T.c("text_primary")
            f = GlassCard()
            fl = QVBoxLayout(f)
            fl.setContentsMargins(16, 14, 16, 14)
            fl.setSpacing(5)
            lab = QLabel(label.upper())
            lab.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {T.c('text_faint')}; letter-spacing: 0.5px; font-family: {FONT_STACK};")
            val = QLabel("—")
            val.setStyleSheet(f"font-size: 20px; font-weight: 800; color: {color}; font-family: {FONT_STACK};")
            fl.addWidget(lab)
            fl.addWidget(val)
            return f, val
        self.card_bal, self.lbl_bal = mk_card("Outstanding Balance", ACCENT_RED)
        self.card_dep, self.lbl_dep = mk_card("Deposit Held", ACCENT_INDIGO)
        self.card_billed, self.lbl_billed = mk_card("Lifetime Billed")
        self.card_paid, self.lbl_paid = mk_card("Lifetime Paid", ACCENT_GREEN)
        for c in [self.card_bal, self.card_dep, self.card_billed, self.card_paid]:
            summary.addWidget(c)
        v.addLayout(summary)

        # 3. Action Buttons Row
        act_row = QHBoxLayout()
        act_row.setSpacing(8)

        self.btn_pay = QPushButton("RECORD PAYMENT")
        self.btn_pay.setIcon(SVGIconManager.icon("dollar", color="#04150E", size=16))
        self.btn_pay.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pay.setFixedHeight(BTN_SECONDARY)
        self.btn_pay.setStyleSheet(primary_button_qss(ACCENT_GREEN, height=BTN_SECONDARY))
        self.btn_pay.clicked.connect(self.record_payment)

        self.btn_advance = QPushButton("RECORD ADVANCE")
        self.btn_advance.setIcon(SVGIconManager.icon("dollar", color=ACCENT_INDIGO, size=16))
        self.btn_advance.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_advance.setFixedHeight(BTN_SECONDARY)
        self.btn_advance.setStyleSheet(pill_button_qss(ACCENT_INDIGO, height=BTN_SECONDARY))
        self.btn_advance.clicked.connect(self.record_advance)

        self.btn_adj = QPushButton("ADJUSTMENT")
        self.btn_adj.setIcon(SVGIconManager.icon("settings", color=ACCENT_AMBER, size=16))
        self.btn_adj.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_adj.setFixedHeight(BTN_SECONDARY)
        self.btn_adj.setStyleSheet(pill_button_qss(ACCENT_AMBER, height=BTN_SECONDARY))
        self.btn_adj.clicked.connect(self.record_adjustment)

        self.btn_edit_bill = QPushButton("EDIT BILL")
        self.btn_edit_bill.setIcon(SVGIconManager.icon("edit", color="#F8FAFC", size=16))
        self.btn_edit_bill.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_edit_bill.setFixedHeight(BTN_SECONDARY)
        self.btn_edit_bill.setStyleSheet(pill_button_qss("#ffffff", height=BTN_SECONDARY))
        self.btn_edit_bill.clicked.connect(self.edit_selected_bill)

        self.btn_delete_bill = QPushButton("DELETE BILL")
        self.btn_delete_bill.setIcon(SVGIconManager.icon("trash", color=ACCENT_RED, size=16))
        self.btn_delete_bill.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_delete_bill.setFixedHeight(BTN_SECONDARY)
        self.btn_delete_bill.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_SECONDARY))
        self.btn_delete_bill.clicked.connect(self.delete_selected_bill)

        self.btn_settle = QPushButton("SETTLE TENANT")
        self.btn_settle.setIcon(SVGIconManager.icon("logout", color=ACCENT_RED, size=16))
        self.btn_settle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_settle.setFixedHeight(BTN_SECONDARY)
        self.btn_settle.setStyleSheet(pill_button_qss(ACCENT_RED, height=BTN_SECONDARY))
        self.btn_settle.clicked.connect(self.settle_tenant)

        for b in [self.btn_pay, self.btn_advance, self.btn_adj, self.btn_edit_bill, self.btn_delete_bill, self.btn_settle]:
            act_row.addWidget(b)
        v.addLayout(act_row)

        # 4. Interactive Transaction Table
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Date", "Type", "Description", "Debit", "Credit", "Balance", "Linked"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.cellDoubleClicked.connect(self.on_row_doubleclicked)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_txn_context_menu)
        self.table.setStyleSheet(table_qss())

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for i in range(3, 7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        v.addWidget(self.table, 1)

        hint = QLabel("Double-click a bill or payment row to edit · Right-click for options (Edit / Relink / Delete)")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet(f"font-family: {FONT_STACK}; background: transparent; font-size: 11px; color: {T.c('text_muted')};")
        v.addWidget(hint)

    def load_data(self):
        uid = self.controller.current_session_id
        if uid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.id, t.name, p.name, t.security_deposit, p.id, p.address, t.tenant_uid, t.is_active
                       FROM tenants t JOIN properties p ON t.property_id=p.id
                       WHERE t.user_id=? ORDER BY t.is_active DESC, t.name ASC""", (uid,))
        rows = cur.fetchall()
        conn.close()
        self.tenant_cache = {}
        self.combo_tenant.blockSignals(True)
        self.combo_tenant.clear()
        self.combo_tenant.addItem("-- Select Tenant --", None)
        for tid, tn, pn, dep, pid, addr, tuid, active in rows:
            suffix = "" if active else " (Closed)"
            self.combo_tenant.addItem(f"{tn} — {pn}{suffix}", tid)
            self.tenant_cache[tid] = {
                "id": tid, "name": tn, "property_name": pn, "property_id": pid,
                "deposit": float(dep or 0), "address": addr,
                "tenant_uid": tuid, "is_active": bool(active),
            }
        self.combo_tenant.blockSignals(False)

    def on_tenant_selected(self):
        tid = self.combo_tenant.currentData()
        self.current_tid = tid
        if not tid or tid not in self.tenant_cache:
            self.table.setRowCount(0)
            self.lbl_bal.setText("—"); self.lbl_dep.setText("—")
            self.lbl_billed.setText("—"); self.lbl_paid.setText("—")
            return
        info = self.tenant_cache[tid]
        bal = get_tenant_balance(tid)
        self.lbl_bal.setText(fmt_money(bal))
        self.lbl_bal.setStyleSheet(f"font-size: 20px; font-weight: 800; color: {ACCENT_RED if bal > 0 else ACCENT_GREEN}; font-family: {FONT_STACK};")
        self.lbl_dep.setText(fmt_money(info["deposit"]))

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(SUM(amount),0) FROM tenant_transactions WHERE tenant_id=? AND txn_type='bill'", (tid,))
        billed = cur.fetchone()[0] or 0.0
        cur.execute("SELECT COALESCE(SUM(amount),0) FROM tenant_transactions WHERE tenant_id=? AND txn_type IN ('payment','advance')", (tid,))
        paid = cur.fetchone()[0] or 0.0
        conn.close()
        self.lbl_billed.setText(fmt_money(billed))
        self.lbl_paid.setText(fmt_money(paid))

        txns = fetch_tenant_ledger(tid)
        self.table.setRowCount(len(txns))
        running = 0.0
        for i, t in enumerate(txns):
            txn_id, txn_date, ttype, desc, amount, method, ref, linked, affects = t
            amount = amount or 0.0
            is_debit = ttype in DEBIT_TYPES
            is_credit = ttype in CREDIT_TYPES
            if is_debit:
                running += amount
                debit_s, credit_s = f"{amount:,.2f}", "—"
            elif is_credit:
                running -= amount
                debit_s, credit_s = "—", f"{amount:,.2f}"
            else:
                debit_s, credit_s = "—", "—"
            cols = [txn_date, ttype.replace("_", " ").title(), desc or "", debit_s, credit_s, f"{running:,.2f}"]
            for j, val in enumerate(cols):
                it = QTableWidgetItem(str(val))
                it.setData(Qt.ItemDataRole.UserRole, txn_id)
                if j in (3, 4, 5):
                    it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if j == 5:
                    if running > 0: it.setForeground(QBrush(QColor(ACCENT_RED)))
                    elif running < 0: it.setForeground(QBrush(QColor(ACCENT_GREEN)))
                if j == 1:
                    if ttype == "payment": it.setForeground(QBrush(QColor(ACCENT_GREEN)))
                    elif ttype == "bill": it.setForeground(QBrush(QColor(ACCENT_SKY)))
                    elif ttype in ("damage_charge", "deposit_adjust"): it.setForeground(QBrush(QColor(ACCENT_RED)))
                    elif ttype == "deposit_refund": it.setForeground(QBrush(QColor(ACCENT_AMBER)))
                self.table.setItem(i, j, it)

            if linked:
                linked_s = f"#{linked}"
                linked_color = QColor(ACCENT_GREEN) if ttype in CREDIT_TYPES else QColor("#94A3B8")
            else:
                linked_s = "—"
                linked_color = QColor("#64748B")
            it = QTableWidgetItem(linked_s)
            it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setForeground(QBrush(linked_color))
            it.setData(Qt.ItemDataRole.UserRole, txn_id)
            self.table.setItem(i, 6, it)

    def _refresh_after_bill_change(self):
        idx = self.combo_tenant.findData(self.current_tid)
        if idx >= 0: self.combo_tenant.setCurrentIndex(idx)
        self.on_tenant_selected()
        self.controller.dashboard_page.refresh_live_metrics()
        self.controller.dashboard_page.refresh_tab_states()

    def _get_selected_bill_id(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            row = self.table.currentRow()
            if row >= 0:
                selected_rows = [self.table.model().index(row, 0)]
        if selected_rows:
            row = selected_rows[0].row()
            type_item = self.table.item(row, 1)
            ttype_str = (type_item.text() if type_item else "").lower().replace(" ", "_")
            if ttype_str == "bill":
                item_0 = self.table.item(row, 0)
                txn_id = item_0.data(Qt.ItemDataRole.UserRole) if item_0 else None
                if txn_id:
                    conn = sqlite3.connect(get_db_path())
                    cur = conn.cursor()
                    cur.execute("SELECT linked_bill_id FROM tenant_transactions WHERE id=?", (txn_id,))
                    r = cur.fetchone()
                    conn.close()
                    if r and r[0]:
                        return r[0]
        return None

    def edit_selected_bill(self):
        info = self._cur_info()
        if not info: return
        bill_id = self._get_selected_bill_id()
        if not bill_id:
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("""SELECT id, billing_month, bill_date, total_payable, balance_remaining
                           FROM tenant_bills WHERE tenant_id=? ORDER BY bill_date DESC, id DESC""", (self.current_tid,))
            bills = cur.fetchall()
            conn.close()
            if not bills:
                QMessageBox.information(self, "No Bills", f"No bills found for {info['name']}.")
                return
            if len(bills) == 1:
                bill_id = bills[0][0]
            else:
                items = [f"Bill #{b[0]} — {b[1]} ({b[2]}) — Total: {fmt_money(b[3])}" for b in bills]
                chosen, ok = QInputDialog.getItem(self, "Select Bill to Edit", "Choose a bill:", items, 0, False)
                if not ok or not chosen:
                    return
                idx = items.index(chosen)
                bill_id = bills[idx][0]

        dlg = EditBillDialog(self.controller.current_session_id, bill_id, parent=self)
        if dlg.exec():
            self._refresh_after_bill_change()

    def delete_selected_bill(self):
        info = self._cur_info()
        if not info: return
        bill_id = self._get_selected_bill_id()
        if not bill_id:
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("""SELECT id, billing_month, bill_date, total_payable, balance_remaining
                           FROM tenant_bills WHERE tenant_id=? ORDER BY bill_date DESC, id DESC""", (self.current_tid,))
            bills = cur.fetchall()
            conn.close()
            if not bills:
                QMessageBox.information(self, "No Bills", f"No bills found for {info['name']}.")
                return
            if len(bills) == 1:
                bill_id = bills[0][0]
            else:
                items = [f"Bill #{b[0]} — {b[1]} ({b[2]}) — Total: {fmt_money(b[3])}" for b in bills]
                chosen, ok = QInputDialog.getItem(self, "Select Bill to Delete", "Choose a bill to delete:", items, 0, False)
                if not ok or not chosen:
                    return
                idx = items.index(chosen)
                bill_id = bills[idx][0]

        self._execute_bill_deletion(bill_id)

    def _execute_bill_deletion(self, bill_id):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT billing_month, total_payable FROM tenant_bills WHERE id=?", (bill_id,))
        r = cur.fetchone()
        conn.close()
        if not r: return
        b_month, b_total = r

        if QMessageBox.question(
            self, "Delete Bill",
            f"Permanently delete the bill for {b_month} ({fmt_money(b_total)})?\n\n"
            "This will remove the bill, update the tenant's balance, rebuild the Excel ledger, "
            "and remove all generated bill PDFs.\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return

        ok, msg = delete_bill_record(self.controller.current_session_id, bill_id)
        if ok:
            if hasattr(self.controller, "toast"):
                self.controller.toast(f"Bill for {b_month} deleted", "trash", ACCENT_AMBER)
            QMessageBox.information(self, "Deleted", msg)
            self._refresh_after_bill_change()
        else:
            QMessageBox.warning(self, "Delete Failed", msg)

    def on_row_doubleclicked(self, row, col):
        it = self.table.item(row, 0)
        if not it: return
        txn_id = it.data(Qt.ItemDataRole.UserRole)
        if not txn_id: return

        type_cell = self.table.item(row, 1)
        ttype_str = (type_cell.text() if type_cell else "").lower().replace(" ", "_")

        if ttype_str in ("payment", "advance", "other_credit"):
            dlg = EditPaymentDialog(self.controller.current_session_id, txn_id, parent=self)
            if dlg.exec():
                self._refresh_after_bill_change()
        elif ttype_str == "bill":
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("SELECT linked_bill_id FROM tenant_transactions WHERE id=?", (txn_id,))
            r = cur.fetchone()
            conn.close()
            bill_id = r[0] if (r and r[0]) else None
            if bill_id:
                dlg = EditBillDialog(self.controller.current_session_id, bill_id, parent=self)
                if dlg.exec():
                    self._refresh_after_bill_change()
        else:
            TransactionDetailDialog(self.controller.current_session_id, txn_id, parent=self).exec()

    def _show_txn_context_menu(self, pos):
        row = self.table.rowAt(pos.y())
        if row < 0: return
        it = self.table.item(row, 0)
        if not it: return
        txn_id = it.data(Qt.ItemDataRole.UserRole)
        if not txn_id: return

        type_cell = self.table.item(row, 1)
        ttype_str = (type_cell.text() if type_cell else "").lower().replace(" ", "_")
        is_payment = ttype_str in ("payment", "advance", "other_credit")
        is_bill = ttype_str == "bill"

        menu = QMenu(self)
        menu.setStyleSheet(f"""
            QMenu {{
                background-color: {T.c('menu_bg')}; color: {T.c('text_primary')};
                border: 1px solid {T.c('menu_border')};
                border-radius: 8px; padding: 6px;
                font-family: {FONT_STACK}; font-size: 12px;
            }}
            QMenu::item {{ padding: 6px 20px; border-radius: 6px; }}
            QMenu::item:selected {{ background-color: {T.c('menu_sel_bg')}; }}
            QMenu::item:disabled {{ color: {T.c('text_muted')}; }}
        """)

        a_edit = a_relink = a_delete = a_edit_bill = a_delete_bill = None
        if is_payment:
            a_edit = menu.addAction("Edit Payment…")
            a_relink = menu.addAction("Relink to Bill…")
            menu.addSeparator()
            a_delete = menu.addAction("Delete Payment")
            menu.addSeparator()
            a_view = menu.addAction("View Audit Details")
        elif is_bill:
            a_edit_bill = menu.addAction("Edit Bill…")
            menu.addSeparator()
            a_delete_bill = menu.addAction("Delete Bill")
            menu.addSeparator()
            a_view = menu.addAction("View Bill Details")
        else:
            info = menu.addAction("This transaction is locked")
            info.setEnabled(False)
            a_view = menu.addAction("View Audit Details")

        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action is None: return

        if is_bill:
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("SELECT linked_bill_id FROM tenant_transactions WHERE id=?", (txn_id,))
            r = cur.fetchone()
            conn.close()
            bill_id = r[0] if (r and r[0]) else None

            if action == a_view:
                if bill_id:
                    dlg = BillDetailDialog(self.controller.current_session_id, bill_id, parent=self)
                    if dlg.exec():
                        self._refresh_after_bill_change()
                else:
                    TransactionDetailDialog(self.controller.current_session_id, txn_id, parent=self).exec()
            elif a_edit_bill and action == a_edit_bill:
                if bill_id:
                    dlg = EditBillDialog(self.controller.current_session_id, bill_id, parent=self)
                    if dlg.exec():
                        self._refresh_after_bill_change()
            elif a_delete_bill and action == a_delete_bill:
                if bill_id:
                    self._execute_bill_deletion(bill_id)
        elif is_payment:
            if action == a_view:
                TransactionDetailDialog(self.controller.current_session_id, txn_id, parent=self).exec()
            elif a_edit and action == a_edit:
                dlg = EditPaymentDialog(self.controller.current_session_id, txn_id, parent=self)
                if dlg.exec():
                    self._refresh_after_bill_change()
            elif a_relink and action == a_relink:
                dlg = RelinkPaymentDialog(self.controller.current_session_id, txn_id, parent=self)
                if dlg.exec():
                    self._refresh_after_bill_change()
            elif a_delete and action == a_delete:
                self._delete_payment(txn_id)
        else:
            if action == a_view:
                TransactionDetailDialog(self.controller.current_session_id, txn_id, parent=self).exec()

    def _delete_payment(self, txn_id):
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT tenant_id, amount, txn_date, txn_type
                       FROM tenant_transactions WHERE id=?""", (txn_id,))
        row = cur.fetchone()
        conn.close()
        if not row: return
        tid, amt, tdate, ttype = row

        if QMessageBox.question(
            self, "Delete Payment",
            f"Delete {ttype} of {fmt_money(amt)} dated {tdate}?\n\n"
            "The tenant's Excel ledger and affected bill PDFs will be regenerated.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        ) != QMessageBox.StandardButton.Yes:
            return

        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("DELETE FROM tenant_transactions WHERE id=?", (txn_id,))
        conn.commit()
        conn.close()

        log_edit(self.controller.current_session_id, tid, txn_id, "delete_payment",
                 {"amount": amt, "txn_date": tdate}, None)
        try:
            delete_receipt_pdfs_for_txn(self.controller.current_session_id, tid, txn_id)
        except Exception as e:
            log.error(f"Receipt deletion error: {e}")

        try:
            recompute_tenant_from_transactions(tid)
        except Exception as e:
            log.error(f"Recompute after delete failed: {e}")
            QMessageBox.warning(self, "Partial Delete", str(e))

        self._refresh_after_bill_change()

    def _cur_info(self):
        tid = self.current_tid
        if not tid or tid not in self.tenant_cache:
            QMessageBox.warning(self, "No Tenant", "Select a tenant first."); return None
        return self.tenant_cache[tid]

    def record_payment(self):
        info = self._cur_info()
        if not info: return
        dlg = PaymentDialog(self.controller.current_session_id, info, parent=self, is_advance=False)
        if dlg.exec():
            idx = self.combo_tenant.findData(self.current_tid)
            if idx >= 0: self.combo_tenant.setCurrentIndex(idx)
            self.on_tenant_selected()
            self.controller.dashboard_page.refresh_live_metrics()

    def record_advance(self):
        info = self._cur_info()
        if not info: return
        dlg = PaymentDialog(self.controller.current_session_id, info, parent=self, is_advance=True)
        if dlg.exec():
            idx = self.combo_tenant.findData(self.current_tid)
            if idx >= 0: self.combo_tenant.setCurrentIndex(idx)
            self.on_tenant_selected()
            self.controller.dashboard_page.refresh_live_metrics()

    def record_adjustment(self):
        info = self._cur_info()
        if not info: return
        dlg = AdjustmentDialog(self.controller.current_session_id, info, parent=self)
        if dlg.exec():
            idx = self.combo_tenant.findData(self.current_tid)
            if idx >= 0: self.combo_tenant.setCurrentIndex(idx)
            self.on_tenant_selected()
            self.controller.dashboard_page.refresh_live_metrics()

    def settle_tenant(self):
        info = self._cur_info()
        if not info: return
        if not info["is_active"]:
            QMessageBox.information(self, "Already Closed", "This tenant lease is already terminated."); return
        if QMessageBox.question(self, "Settle Tenant", f"Begin exit settlement for {info['name']}?",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        dlg = SettleTenantWizard(self.controller.current_session_id, info, parent=self)
        if dlg.exec():
            self.load_data()
            idx = self.combo_tenant.findData(self.current_tid)
            if idx >= 0: self.combo_tenant.setCurrentIndex(idx)
            self.controller.dashboard_page.refresh_tab_states()
            self.controller.dashboard_page.refresh_live_metrics()

    def open_deposits(self):
        DepositsReportDialog(self.controller.current_session_id, parent=self).exec()

    def open_files(self):
        tid = self.current_tid
        info = self.tenant_cache.get(tid) if tid else None
        dlg = TenantFilesDialog(self.controller.current_session_id, tenant_info=info, parent=self)
        dlg.exec()


# ==============================================================================
# § 14  REPORTS & EXPORTS
# ==============================================================================
class ReportsSubPage(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.init_ui()

    def init_ui(self):
        v = QVBoxLayout(self)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(14)

        top = QHBoxLayout()
        tb = QVBoxLayout()
        tb.setSpacing(2)
        t = QLabel("Reports & Data Exports")
        t.setStyleSheet(f"font-size: {FS_SUB}; font-weight: 700; color: {T.c('text_primary')}; background: transparent; font-family: {FONT_STACK};")
        s = QLabel("Generate monthly summaries, tenant account statements, and accounting CSVs")
        s.setStyleSheet(f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; background: transparent; font-family: {FONT_STACK};")
        tb.addWidget(t)
        tb.addWidget(s)
        top.addLayout(tb)
        top.addStretch(1)
        v.addLayout(top)

        # 1. Monthly Summary Card
        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        cl.addWidget(self._hdr("Monthly Audit Summary", "Consolidated billing, collection, and outstanding report across all properties"))
        row = QHBoxLayout()
        row.setSpacing(10)
        self.combo_month = QComboBox()
        setup_combobox(self.combo_month)
        self.combo_month.addItems(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
        self.combo_month.setCurrentIndex(QDate.currentDate().month() - 1)
        self.combo_month.setStyleSheet(INPUT_STYLE)
        self.combo_month.setFixedHeight(INPUT_H)

        self.combo_year = QComboBox()
        setup_combobox(self.combo_year)
        cy = QDate.currentDate().year()
        for y in range(cy - 3, cy + 3): self.combo_year.addItem(str(y))
        self.combo_year.setCurrentText(str(cy))
        self.combo_year.setStyleSheet(INPUT_STYLE)
        self.combo_year.setFixedHeight(INPUT_H)

        gen = QPushButton("GENERATE PDF")
        gen.setIcon(SVGIconManager.icon("file", color="#04150E", size=16))
        gen.setCursor(Qt.CursorShape.PointingHandCursor)
        gen.setFixedHeight(BTN_PRIMARY)
        gen.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        gen.clicked.connect(self.gen_monthly)

        row.addWidget(self.combo_month, 1)
        row.addWidget(self.combo_year, 1)
        row.addWidget(gen)
        cl.addLayout(row)
        v.addWidget(card)

        # 2. Tenant Statement Card
        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        cl.addWidget(self._hdr("Tenant Account Statement", "Detailed historical financial ledger statement for individual tenant audit"))
        row = QHBoxLayout()
        row.setSpacing(10)
        self.combo_tenant = QComboBox()
        self.combo_tenant.setStyleSheet(INPUT_STYLE)
        self.combo_tenant.setFixedHeight(INPUT_H)

        gen2 = QPushButton("GENERATE PDF")
        gen2.setIcon(SVGIconManager.icon("file", color="#04150E", size=16))
        gen2.setCursor(Qt.CursorShape.PointingHandCursor)
        gen2.setFixedHeight(BTN_PRIMARY)
        gen2.setStyleSheet(primary_button_qss(ACCENT_INDIGO, height=BTN_PRIMARY))
        gen2.clicked.connect(self.gen_statement)

        row.addWidget(self.combo_tenant, 1)
        row.addWidget(gen2)
        cl.addLayout(row)
        v.addWidget(card)

        # 3. CSV Transactions Export
        card = GlassCard()
        cl = QVBoxLayout(card)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(10)
        cl.addWidget(self._hdr("CSV Raw Data Export", "Export all database transaction rows for spreadsheet analysis or accounting software"))

        exp = QPushButton("EXPORT TRANSACTIONS CSV")
        exp.setIcon(SVGIconManager.icon("download", color="#04150E", size=16))
        exp.setCursor(Qt.CursorShape.PointingHandCursor)
        exp.setFixedHeight(BTN_PRIMARY)
        exp.setStyleSheet(primary_button_qss(ACCENT_AMBER, height=BTN_PRIMARY))
        exp.clicked.connect(self.export_csv)
        cl.addWidget(exp)
        v.addWidget(card)

        v.addStretch(1)

    def _hdr(self, title, sub):
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        l = QVBoxLayout(w)
        l.setContentsMargins(0,0,0,0)
        l.setSpacing(2)
        t = QLabel(title)
        t.setStyleSheet(f"font-size: 13px; font-weight: 700; color: {T.c('text_primary')}; font-family: {FONT_STACK};")
        s = QLabel(sub)
        s.setStyleSheet(f"font-size: 11px; color: {T.c('text_secondary')}; font-family: {FONT_STACK};")
        l.addWidget(t)
        l.addWidget(s)
        return w

    def load_reports_data(self):
        uid = self.controller.current_session_id
        if uid is None: return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.id, t.name, p.name FROM tenants t
                       JOIN properties p ON t.property_id=p.id
                       WHERE t.user_id=? ORDER BY t.name""", (uid,))
        rows = cur.fetchall()
        conn.close()
        self.combo_tenant.clear()
        for tid, tn, pn in rows:
            self.combo_tenant.addItem(f"{tn} — {pn}", tid)

    def gen_monthly(self):
        uid = self.controller.current_session_id
        if uid is None: return
        month_str = f"{self.combo_month.currentText()} {self.combo_year.currentText()}"
        out_dir = os.path.join(get_output_root(uid), "Reports", "Monthly_Summaries")
        ensure_dir(out_dir)
        path = os.path.join(out_dir, f"Monthly_Summary_{sanitize_fs_name(month_str)}.pdf")
        try:
            generate_monthly_summary_pdf(path, uid, month_str)
            QMessageBox.information(self, "Generated", f"Monthly summary PDF saved:\n\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def gen_statement(self):
        uid = self.controller.current_session_id
        tid = self.combo_tenant.currentData()
        if not tid:
            QMessageBox.warning(self, "Select", "Select a tenant first."); return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("""SELECT t.name, t.tenant_uid, p.name, p.address
                       FROM tenants t JOIN properties p ON t.property_id=p.id
                       WHERE t.id=?""", (tid,))
        row = cur.fetchone()
        if not row: conn.close(); return
        info = {"name": row[0], "tenant_uid": row[1], "property_name": row[2], "address": row[3]}
        cur.execute("""SELECT txn_date, txn_type, description, amount FROM tenant_transactions
                       WHERE tenant_id=? ORDER BY txn_date, id""", (tid,))
        txns = [{"txn_date": r[0], "txn_type": r[1], "description": r[2], "amount": r[3]} for r in cur.fetchall()]
        conn.close()
        tdir = get_tenant_dir(uid, info["address"], info["property_name"], info["name"], info["tenant_uid"])
        path = os.path.join(tdir, f"Statement_{datetime.now().strftime('%Y%m%d')}.pdf")
        try:
            generate_tenant_statement_pdf(path, info, txns)
            QMessageBox.information(self, "Generated", f"Statement saved:\n\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def export_csv(self):
        uid = self.controller.current_session_id
        if uid is None: return
        out_dir = os.path.join(get_output_root(uid), "Exports")
        ensure_dir(out_dir)
        path = os.path.join(out_dir, f"transactions_{datetime.now().strftime('%Y%m%d_%H%M')}.csv")
        try:
            n = export_transactions_csv(uid, path)
            QMessageBox.information(self, "Exported", f"{n} transactions exported to:\n\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))


# ==============================================================================
# § 15  DASHBOARD & MAIN VIEWPORT
# ==============================================================================
class DashboardCard(QWidget):
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self._current_reel_index = 0
        self._nav_items = []
        self.init_ui()

    def init_ui(self):
        master = QHBoxLayout(self)
        master.setContentsMargins(16, 16, 16, 16)
        master.setSpacing(16)

        # 1. Glass Navigation Sidebar
        self.nav = QFrame()
        self.nav.setObjectName("NavSidebar")
        self.nav.setFixedWidth(240)
        self.nav.setStyleSheet(f"""
            QFrame#NavSidebar {{
                background-color: {T.c('nav_bg')};
                border: 1px solid {T.c('border')};
                border-top: 1px solid {T.c('border_specular')};
                border-radius: {RADIUS_CARD}px;
            }}
            QLabel {{
                background: transparent;
                border: none;
                color: {T.c('text_primary')};
                font-family: {FONT_STACK};
            }}
        """)

        sb = QVBoxLayout(self.nav)
        sb.setContentsMargins(16, 24, 16, 20)
        sb.setSpacing(6)

        # macOS Traffic Lights window header
        lights_row = QHBoxLayout()
        lights_row.setContentsMargins(4, 0, 4, 14)
        lights_row.setSpacing(8)

        red_light = QFrame()
        red_light.setFixedSize(12, 12)
        red_light.setCursor(Qt.CursorShape.PointingHandCursor)
        red_light.setToolTip("Close Window")
        red_light.setStyleSheet("""
            QFrame { background-color: #FF5F56; border: 1px solid #E0443E; border-radius: 6px; }
            QFrame:hover { background-color: #FF3B30; }
        """)
        red_light.mousePressEvent = lambda e: self.controller.close()

        yellow_light = QFrame()
        yellow_light.setFixedSize(12, 12)
        yellow_light.setCursor(Qt.CursorShape.PointingHandCursor)
        yellow_light.setToolTip("Minimize Window")
        yellow_light.setStyleSheet("""
            QFrame { background-color: #FFBD2E; border: 1px solid #DEA123; border-radius: 6px; }
            QFrame:hover { background-color: #FF9500; }
        """)
        yellow_light.mousePressEvent = lambda e: self.controller.showMinimized()

        green_light = QFrame()
        green_light.setFixedSize(12, 12)
        green_light.setCursor(Qt.CursorShape.PointingHandCursor)
        green_light.setToolTip("Toggle Fullscreen / Maximize")
        green_light.setStyleSheet("""
            QFrame { background-color: #27C93F; border: 1px solid #1AAB29; border-radius: 6px; }
            QFrame:hover { background-color: #34C759; }
        """)
        def _toggle_max(e):
            if self.controller.isMaximized():
                self.controller.showNormal()
            else:
                self.controller.showMaximized()
        green_light.mousePressEvent = _toggle_max

        lights_row.addWidget(red_light)
        lights_row.addWidget(yellow_light)
        lights_row.addWidget(green_light)
        lights_row.addStretch(1)
        sb.addLayout(lights_row)

        # Brand header
        brand_row = QHBoxLayout()
        brand_row.setContentsMargins(4, 0, 4, 0)
        brand_row.setSpacing(10)
        brand_dot = QFrame()
        brand_dot.setFixedSize(10, 10)
        brand_dot.setStyleSheet("background-color: #FF6A00; border-radius: 5px;")
        brand_lbl = QLabel("<span style='color: #FFFFFF;'>APEX</span> <span style='color: #FF6A00;'>LEDGER</span> <span style='color: #888888; font-size: 10px;'>PRO</span>")
        brand_lbl.setStyleSheet(
            f"font-size: 13px; font-weight: 800; letter-spacing: 1.6px; "
            f"background: transparent; font-family: {FONT_STACK};"
        )
        brand_row.addWidget(brand_dot)
        brand_row.addWidget(brand_lbl)
        brand_row.addStretch(1)
        sb.addLayout(brand_row)
        sb.addSpacing(20)

        nav_lbl = QLabel("PORTFOLIO NAVIGATION")
        nav_lbl.setStyleSheet(
            f"color: {T.c('text_faint')}; font-size: 9px; "
            f"font-weight: 700; letter-spacing: 1.4px; "
            f"padding: 0 8px 6px 8px; background: transparent; "
            f"font-family: {FONT_STACK};"
        )
        sb.addWidget(nav_lbl)

        def nav_btn(text, icon_name, route):
            b = QPushButton(f"  {text}")
            b.setIcon(SVGIconManager.icon(icon_name, color="#8A96AC", size=16))
            b.setIconSize(QSize(16, 16))
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setStyleSheet(self._nav_btn_qss(active=False))
            b.clicked.connect(route)
            return b

        self.btn_dash    = nav_btn("Dashboard",      "dashboard",  lambda: self.animate_reel_slide(0))
        self.btn_owner   = nav_btn("Owner Details",  "owner",      self.route_owner)
        self.btn_props   = nav_btn("Properties",     "properties", self.route_props)
        self.btn_tenants = nav_btn("Tenants",        "tenants",    self.route_tenants)
        self.btn_water   = nav_btn("Water Record",   "water",      self.route_water_record)
        self.btn_billing = nav_btn("Billing",        "billing",    self.route_billing)
        self.btn_records = nav_btn("Records",        "records",    self.route_records)
        self.btn_reports = nav_btn("Reports",        "reports",    self.route_reports)

        self._nav_items = [
            (self.btn_dash, 0, "dashboard"),
            (self.btn_owner, 1, "owner"),
            (self.btn_props, 2, "properties"),
            (self.btn_tenants, 3, "tenants"),
            (self.btn_water, 4, "water"),
            (self.btn_billing, 5, "billing"),
            (self.btn_records, 6, "records"),
            (self.btn_reports, 7, "reports"),
        ]
        for b, _, _ in self._nav_items:
            sb.addWidget(b)

        sb.addStretch(1)

        self.version_lbl = QLabel(f"v{APP_VERSION} · Apex Ledger Pro")
        self.version_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.version_lbl.setStyleSheet(
            f"color: {T.c('text_faint')}; font-size: 10px; "
            f"padding-top: 8px; background: transparent; "
            f"font-family: {FONT_STACK};"
        )
        sb.addWidget(self.version_lbl)

        # 2. Main Viewport & Sliding Reel
        self.viewport = QWidget()
        self.viewport.setObjectName("ViewportWindow")
        self.viewport.setStyleSheet("QWidget#ViewportWindow { background: transparent; }")
        self.reel = QWidget(self.viewport)
        self.reel.setStyleSheet("background: transparent;")

        # Home Dashboard Subpage with responsive Linux scroll area
        self.home = QWidget(self.reel)
        self.home.setStyleSheet("background: transparent;")
        home_outer_lay = QVBoxLayout(self.home)
        home_outer_lay.setContentsMargins(0, 0, 0, 0)

        self.home_scroll = QScrollArea()
        self.home_scroll.setWidgetResizable(True)
        self.home_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.home_scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 24px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.home_content = QWidget()
        self.home_content.setStyleSheet("background: transparent;")
        hv = QVBoxLayout(self.home_content)
        hv.setContentsMargins(4, 0, 10, 8)
        hv.setSpacing(14)

        hdr = QHBoxLayout()
        tb = QVBoxLayout()
        tb.setSpacing(2)
        main_t = QLabel("<span style='color: #FFFFFF;'>PORTFOLIO</span> <span style='color: #FF6A00;'>OVERVIEW</span>")
        main_t.setStyleSheet(
            f"font-size: {FS_SUB}; font-weight: 800; letter-spacing: 1px; "
            f"background: transparent; font-family: {FONT_STACK};"
        )
        self.owner_lbl = QLabel("Property Owner Node: Verification Required")
        self.owner_lbl.setStyleSheet(
            f"font-size: {FS_HELPER}; color: {T.c('text_secondary')}; "
            f"background: transparent; font-family: {FONT_STACK};"
        )
        tb.addWidget(main_t)
        tb.addWidget(self.owner_lbl)
        hdr.addLayout(tb)
        hdr.addStretch(1)

        logout = QPushButton("LOG OUT")
        logout.setIcon(SVGIconManager.icon("logout", color=ACCENT_RED, size=14))
        logout.setCursor(Qt.CursorShape.PointingHandCursor)
        logout.setFixedHeight(BTN_INLINE)
        logout.setStyleSheet(btn_qss(ACCENT_RED, solid=False, height=BTN_INLINE))
        logout.clicked.connect(self.handle_logout)
        hdr.addWidget(logout, alignment=Qt.AlignmentFlag.AlignVCenter)
        hv.addLayout(hdr)

        # Metric Cards Row 1
        r1 = QHBoxLayout()
        r1.setSpacing(10)
        self.c_billed, self.l_billed = self._metric("TOTAL BILLED",    T.c("text_primary"), "dollar")
        self.c_coll,   self.l_coll   = self._metric("TOTAL COLLECTED", ACCENT_GREEN,        "check")
        self.c_out,    self.l_out    = self._metric("OUTSTANDING DUES", ACCENT_RED,         "alert")
        self.c_adv,    self.l_adv    = self._metric("ADVANCE CREDIT",  ACCENT_SKY,         "download")
        for c in [self.c_billed, self.c_coll, self.c_out, self.c_adv]:
            r1.addWidget(c)
        hv.addLayout(r1)

        # Metric Cards Row 2
        r2 = QHBoxLayout()
        r2.setSpacing(10)
        self.c_dep,     self.l_dep     = self._metric("DEPOSITS HELD",  ACCENT_VIOLET,      "folder")
        self.c_unpaid,  self.l_unpaid  = self._metric("UNPAID BILLS",   ACCENT_AMBER,       "billing")
        self.c_tenants, self.l_tenants = self._metric("ACTIVE TENANTS", T.c("text_primary"), "tenants")
        self.c_props,   self.l_props   = self._metric("PROPERTIES",     T.c("text_primary"), "properties")
        for c in [self.c_dep, self.c_unpaid, self.c_tenants, self.c_props]:
            r2.addWidget(c)
        hv.addLayout(r2)

        chart_head = QHBoxLayout()
        chart_head.addWidget(section_label("Billed vs Collected (Last 6 Months)"))
        chart_head.addStretch(1)
        hv.addLayout(chart_head)

        self.chart = UtilityMetricsBarChart()
        hv.addWidget(self.chart, 1)

        unpaid_head = QHBoxLayout()
        unpaid_head.addWidget(section_label("Unpaid Bills — Oldest First (Click to View)"))
        unpaid_head.addStretch(1)
        hv.addLayout(unpaid_head)

        self.unpaid_scroll = QScrollArea()
        self.unpaid_scroll.setWidgetResizable(True)
        self.unpaid_scroll.setMinimumHeight(140)
        self.unpaid_scroll.setMaximumHeight(240)
        self.unpaid_scroll.setStyleSheet(f"""
            QScrollArea {{ background: transparent; border: none; outline: none; }}
            QScrollBar:vertical {{ background: transparent; width: 6px; }}
            QScrollBar::handle:vertical {{ background: {T.c('scroll_handle')}; border-radius: 3px; min-height: 20px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.unpaid_container = QWidget()
        self.unpaid_container.setStyleSheet("background: transparent;")
        self.unpaid_layout = QVBoxLayout(self.unpaid_container)
        self.unpaid_layout.setContentsMargins(0, 0, 0, 0)
        self.unpaid_layout.setSpacing(6)
        self.unpaid_scroll.setWidget(self.unpaid_container)
        hv.addWidget(self.unpaid_scroll)

        self.home_scroll.setWidget(self.home_content)
        home_outer_lay.addWidget(self.home_scroll)

        # Reel subpages
        self.owner_page   = OwnerDetailsSubPage(self.controller);       self.owner_page.setParent(self.reel)
        self.props_page   = PropertyManagementSubPage(self.controller); self.props_page.setParent(self.reel)
        self.tenants_page = TenantManagementSubPage(self.controller);   self.tenants_page.setParent(self.reel)
        self.water_page   = WaterRecordSubPage(self.controller);        self.water_page.setParent(self.reel)
        self.billing_page = BillingEntrySubPage(self.controller);       self.billing_page.setParent(self.reel)
        self.records_page = RecordsSubPage(self.controller);            self.records_page.setParent(self.reel)
        self.reports_page = ReportsSubPage(self.controller);            self.reports_page.setParent(self.reel)

        master.addWidget(self.nav)
        master.addWidget(self.viewport, 1)

        self._sync_nav_active()

    def _nav_btn_qss(self, active=False):
        if active:
            # Active navigation item: Solid white rounded pill with bold black text
            return f"""
                QPushButton {{
                    background-color: #FFFFFF;
                    color: #000000;
                    border: none;
                    border-radius: 12px;
                    padding: 10px 16px;
                    text-align: left;
                    font-size: 13px;
                    font-weight: 700;
                    font-family: {FONT_STACK};
                    outline: none;
                }}
                QPushButton:hover {{
                    background-color: #EAEAEA;
                    color: #000000;
                }}
                QPushButton:focus {{ outline: none; }}
            """
        return f"""
            QPushButton {{
                background-color: transparent;
                color: #A0A0A0;
                border: 1px solid transparent;
                border-radius: 12px;
                padding: 10px 16px;
                text-align: left;
                font-size: 13px;
                font-weight: 500;
                font-family: {FONT_STACK};
                outline: none;
            }}
            QPushButton:hover {{
                background-color: rgba(255, 255, 255, 0.08);
                border: 1px solid rgba(255, 255, 255, 0.12);
                color: #FFFFFF;
            }}
            QPushButton:focus {{ outline: none; }}
            QPushButton:disabled {{
                color: {T.c('text_faint')};
                background-color: transparent;
            }}
        """

    def _sync_nav_active(self):
        if not hasattr(self, "_nav_items"):
            return
        for btn, idx, icon_name in self._nav_items:
            try:
                is_active = (idx == self._current_reel_index)
                btn.setStyleSheet(self._nav_btn_qss(active=is_active))
                btn.setIcon(SVGIconManager.icon(icon_name, color="#000000" if is_active else "#A0A0A0", size=16))
            except RuntimeError:
                return

    def _metric(self, label, color, icon_name="dollar"):
        f = QFrame()
        f.setObjectName("MetricCard")
        f.setStyleSheet(f"""
            QFrame#MetricCard {{
                background-color: {T.c('card')};
                border: 1px solid {T.c('border')};
                border-top: 1px solid {T.c('border_specular')};
                border-radius: {RADIUS_CARD}px;
            }}
            QFrame#MetricCard:hover {{
                background-color: {T.c('card_hover')};
                border: 1px solid {T.c('border_hover')};
                border-top: 1px solid rgba(255, 255, 255, 0.40);
            }}
            QLabel {{
                background: transparent;
                border: none;
            }}
        """)
        fl = QVBoxLayout(f)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(6)

        # Small white line icon on top
        icon_lbl = QLabel()
        icon_lbl.setPixmap(SVGIconManager.pixmap(icon_name, color="#FFFFFF", size=20))
        icon_lbl.setFixedSize(22, 22)
        fl.addWidget(icon_lbl)

        # macOS-style animated metric value label (30px)
        is_curr = (icon_name != "tenants" and icon_name != "properties" and label != "UNPAID BILLS")
        val = AnimatedNumberLabel("—", prefix="₹", is_currency=is_curr)
        val.setStyleSheet(
            f"font-size: 30px; font-weight: 600; "
            f"letter-spacing: -0.02em; color: {color if color else '#FFFFFF'}; "
            f"background: transparent; font-family: {FONT_STACK};"
        )
        fl.addWidget(val)

        # Small bold label under it
        lab = QLabel(label)
        lab.setStyleSheet(
            f"font-size: 11px; font-weight: 700; color: #A0A0A0; "
            f"letter-spacing: 1.2px; background: transparent; "
            f"font-family: {FONT_STACK};"
        )
        fl.addWidget(lab)
        return f, val

    def populate_verified_session(self, name):
        self.owner_lbl.setText(f"Property Owner Node: {name.upper()} (Verified)")

    def handle_logout(self):
        self.controller.current_session_id = None
        for b in [self.btn_props, self.btn_tenants, self.btn_water,
                  self.btn_billing, self.btn_records, self.btn_reports]:
            b.setEnabled(False)
        self.owner_lbl.setText("Property Owner Node: Verification Required")
        self.controller.fade_to_page(0)

    def refresh_tab_states(self):
        uid = self.controller.current_session_id
        if uid is None:
            for b in [self.btn_props, self.btn_tenants, self.btn_water,
                      self.btn_billing, self.btn_records, self.btn_reports]:
                b.setEnabled(False)
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()
        cur.execute("SELECT user_id FROM profiles WHERE user_id=?", (uid,))
        has_profile = cur.fetchone() is not None
        cur.execute("SELECT id FROM properties WHERE user_id=?", (uid,))
        has_props = cur.fetchone() is not None
        cur.execute("SELECT id FROM properties WHERE user_id=? AND water_meter_type='Common'", (uid,))
        has_common = cur.fetchone() is not None
        cur.execute("SELECT id FROM tenants WHERE user_id=? AND is_active=1", (uid,))
        has_tenants = cur.fetchone() is not None
        conn.close()
        self.btn_props.setEnabled(has_profile)
        self.btn_tenants.setEnabled(has_profile and has_props)
        self.btn_water.setEnabled(has_profile and has_common)
        self.btn_billing.setEnabled(has_profile and has_props and has_tenants)
        self.btn_records.setEnabled(has_profile and has_props)
        self.btn_reports.setEnabled(has_profile and has_props)

    def refresh_live_metrics(self):
        uid = self.controller.current_session_id
        if uid is None:
            return
        conn = sqlite3.connect(get_db_path())
        cur = conn.cursor()

        cur.execute("SELECT COALESCE(SUM(total_payable),0) FROM tenant_bills WHERE user_id=?", (uid,))
        total_billed = cur.fetchone()[0] or 0.0
        cur.execute("SELECT COALESCE(SUM(amount_paid),0) FROM tenant_bills WHERE user_id=?", (uid,))
        total_paid = cur.fetchone()[0] or 0.0
        cur.execute("SELECT COALESCE(SUM(balance_remaining),0) FROM tenant_bills WHERE user_id=? AND balance_remaining > 0.001 AND payment_status != 'Paid' AND settled_status NOT IN ('closed', 'closed_manual', 'settled', 'settled_deposit')", (uid,))
        total_out = cur.fetchone()[0] or 0.0
        cur.execute("SELECT COALESCE(SUM(security_deposit),0) FROM tenants WHERE user_id=? AND is_active=1", (uid,))
        deposits = cur.fetchone()[0] or 0.0
        cur.execute("SELECT COUNT(*) FROM tenant_bills WHERE user_id=? AND balance_remaining > 0.001 AND payment_status != 'Paid' AND settled_status NOT IN ('closed', 'closed_manual', 'settled', 'settled_deposit')", (uid,))
        unpaid_count = cur.fetchone()[0] or 0
        cur.execute("SELECT COUNT(*) FROM tenants WHERE user_id=? AND is_active=1", (uid,))
        active_tenants = cur.fetchone()[0] or 0
        cur.execute("SELECT COUNT(*) FROM properties WHERE user_id=?", (uid,))
        prop_count = cur.fetchone()[0] or 0

        advance_credit = 0.0
        cur.execute("SELECT id FROM tenants WHERE user_id=? AND is_active=1", (uid,))
        for (tid,) in cur.fetchall():
            b = get_tenant_balance(tid)
            if b < 0:
                advance_credit += -b

        months, billed_series, collected_series = [], [], []
        now = QDate.currentDate()
        for i in range(5, -1, -1):
            d = now.addMonths(-i)
            months.append(d.toString("MMM").upper())
            m_str = d.toString("MMM yyyy")
            cur.execute("SELECT COALESCE(SUM(total_payable),0) FROM tenant_bills WHERE user_id=? AND billing_month=?", (uid, m_str))
            billed_series.append(cur.fetchone()[0] or 0.0)
            cur.execute("SELECT COALESCE(SUM(amount_paid),0) FROM tenant_bills WHERE user_id=? AND billing_month=?", (uid, m_str))
            collected_series.append(cur.fetchone()[0] or 0.0)

        cur.execute("""SELECT b.id, t.name, p.name, b.billing_month, b.balance_remaining, b.bill_date
                       FROM tenant_bills b JOIN tenants t ON b.tenant_id = t.id
                       JOIN properties p ON t.property_id = p.id
                       WHERE b.user_id=? AND b.balance_remaining > 0.001
                         AND b.payment_status != 'Paid'
                         AND b.settled_status NOT IN ('closed', 'closed_manual', 'settled', 'settled_deposit')
                       ORDER BY b.bill_date ASC LIMIT 50""", (uid,))
        unpaid_rows = cur.fetchall()
        conn.close()

        if hasattr(self.l_billed, "animate_to_value"):
            self.l_billed.animate_to_value(total_billed)
            self.l_coll.animate_to_value(total_paid)
            self.l_out.animate_to_value(total_out)
            self.l_adv.animate_to_value(advance_credit)
            self.l_dep.animate_to_value(deposits)
            self.l_unpaid.animate_to_value(unpaid_count)
            self.l_tenants.animate_to_value(active_tenants)
            self.l_props.animate_to_value(prop_count)
        else:
            self.l_billed.setText(fmt_money(total_billed))
            self.l_coll.setText(fmt_money(total_paid))
            self.l_out.setText(fmt_money(total_out))
            self.l_adv.setText(fmt_money(advance_credit))
            self.l_dep.setText(fmt_money(deposits))
            self.l_unpaid.setText(str(unpaid_count))
            self.l_tenants.setText(str(active_tenants))
            self.l_props.setText(str(prop_count))
        self.chart.set_data(months, billed_series, collected_series)

        while self.unpaid_layout.count():
            it = self.unpaid_layout.takeAt(0)
            if it.widget():
                it.widget().deleteLater()

        if not unpaid_rows:
            good = QLabel("All rent bills paid. Zero outstanding balances.")
            good.setStyleSheet(
                f"color: {ACCENT_GREEN}; font-size: 12px; padding: 12px; "
                f"background: transparent; font-family: {FONT_STACK};"
            )
            good.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.unpaid_layout.addWidget(good)
        else:
            for bid, tname, pname, month, bal, bdate in unpaid_rows:
                row = QFrame()
                row.setStyleSheet(
                    "QFrame { background-color: rgba(239,68,68,0.08); "
                    "border: 1px solid rgba(239,68,68,0.22); border-radius: 10px; }"
                    "QFrame:hover { background-color: rgba(239,68,68,0.16); "
                    "border: 1px solid rgba(239,68,68,0.45); }"
                )
                row.setCursor(Qt.CursorShape.PointingHandCursor)
                rl = QHBoxLayout(row)
                rl.setContentsMargins(14, 9, 14, 9)
                rl.setSpacing(12)
                l1 = QLabel(f"<b style='color:{T.c('text_primary')};'>{tname}</b> "
                            f"<span style='color:{T.c('text_secondary')};'>· {pname}</span>")
                l1.setStyleSheet(f"font-size: 12px; font-family: {FONT_STACK};")
                l2 = QLabel(f"<span style='color:{T.c('text_dim')};'>{month} · {bdate}</span>")
                l2.setStyleSheet(f"font-size: 11px; font-family: {FONT_STACK};")
                l3 = QLabel(f"<b style='color:{ACCENT_RED}; font-size:12px;'>{fmt_money(bal)}</b>")
                l3.setStyleSheet(f"font-size: 12px; font-family: {FONT_STACK};")
                rl.addWidget(l1)
                rl.addStretch(1)
                rl.addWidget(l2)
                rl.addWidget(l3)
                row.mousePressEvent = (lambda e, _bid=bid: self._open_bill(_bid))
                self.unpaid_layout.addWidget(row)
        self.unpaid_layout.addStretch(1)

    def _open_bill(self, bill_id):
        dlg = BillDetailDialog(self.controller.current_session_id, bill_id, parent=self)
        if dlg.exec():
            self.refresh_live_metrics()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        w = self.viewport.width()
        h = self.viewport.height()
        if w <= 0 or h <= 0:
            return
        self.reel.resize(w, h * 8)
        self.home.setGeometry(0, 0, w, h)
        self.owner_page.setGeometry(0, h, w, h)
        self.props_page.setGeometry(0, h * 2, w, h)
        self.tenants_page.setGeometry(0, h * 3, w, h)
        self.water_page.setGeometry(0, h * 4, w, h)
        self.billing_page.setGeometry(0, h * 5, w, h)
        self.records_page.setGeometry(0, h * 6, w, h)
        self.reports_page.setGeometry(0, h * 7, w, h)
        self.reel.move(0, -self._current_reel_index * h)

    def animate_reel_slide(self, idx):
        self._current_reel_index = idx
        target_y = -idx * self.viewport.height()
        anim = QPropertyAnimation(self.reel, b"pos")
        anim.setDuration(400)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.setStartValue(self.reel.pos())
        anim.setEndValue(QPoint(0, target_y))
        anim.start()
        self._slide_anim = anim
        self._sync_nav_active()

        if idx == 0 and self.controller.current_session_id is not None:
            QTimer.singleShot(50, self.refresh_live_metrics)

    def route_owner(self):
        self.owner_page.load_existing_profile()
        self.animate_reel_slide(1)
    def route_props(self):
        self.props_page.load_properties()
        self.animate_reel_slide(2)
    def route_tenants(self):
        self.tenants_page.load_tenants()
        self.animate_reel_slide(3)
    def route_water_record(self):
        self.water_page.refresh_page_data()
        self.animate_reel_slide(4)
    def route_billing(self):
        self.billing_page.refresh_page_data()
        self.animate_reel_slide(5)
    def route_records(self):
        self.records_page.load_data()
        self.animate_reel_slide(6)
    def route_reports(self):
        self.reports_page.load_reports_data()
        self.animate_reel_slide(7)


# ==============================================================================
# § 16  MAIN WINDOW & ENTRY POINT
# ==============================================================================
class PremiumMainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION} · {APP_SUBTITLE}")
        self.resize(1240, 800)
        self.setMinimumSize(980, 680)
        self.current_session_id = None
        self._bg_cache = None

        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet("PremiumMainWindow { background-color: #050505; }")

        # Glass Container Envelope
        self.card = QFrame(self)
        self.card.setObjectName("MainEnvelopeFrame")
        self.card.setStyleSheet(f"""
            QFrame#MainEnvelopeFrame {{
                background-color: {T.c('surface')};
                border: 1px solid {T.c('border')};
                border-top: 1px solid {T.c('border_specular')};
                border-radius: {RADIUS_MODAL}px;
            }}
        """)

        cl = QVBoxLayout(self.card)
        cl.setContentsMargins(0, 0, 0, 0)

        self.stack = QStackedWidget(self.card)
        self.login_card    = LoginCard(controller=self)
        self.signup_card   = SignUpCard(controller=self)
        self.dashboard_page = DashboardCard(controller=self)

        self.stack.addWidget(self.login_card)
        self.stack.addWidget(self.signup_card)
        self.stack.addWidget(self.dashboard_page)
        cl.addWidget(self.stack)

        try:
            conn = sqlite3.connect(get_db_path())
            cur = conn.cursor()
            cur.execute("SELECT id FROM users LIMIT 1")
            exists = cur.fetchone() is not None
            conn.close()
        except Exception:
            exists = False
        self.stack.setCurrentIndex(0 if exists else 1)

        self.opacity = QGraphicsOpacityEffect(self.card)
        self.opacity.setOpacity(1.0)

        # macOS Floating Dynamic Island Toast Notification
        self.toast_widget = MacOSToast(self)

        QTimer.singleShot(0, self.position_layers)

    def toast(self, message, icon_name="check", color="#10B981"):
        if hasattr(self, "toast_widget"):
            self.toast_widget.show_message(message, icon_name=icon_name, color=color)

    def fade_to_page(self, target_index):
        if hasattr(self, "morph_group") and self.morph_group.state() == self.morph_group.State.Running:
            return

        if self.card.graphicsEffect() is None:
            self.card.setGraphicsEffect(self.opacity)

        self.morph_group = QSequentialAnimationGroup()

        fo = QPropertyAnimation(self.opacity, b"opacity")
        fo.setDuration(180)
        fo.setStartValue(1.0)
        fo.setEndValue(0.0)
        fo.setEasingCurve(QEasingCurve.Type.InCubic)

        def morph():
            self.card.setUpdatesEnabled(False)
            self.stack.setCurrentIndex(target_index)

            if target_index == 2:
                pad = 12
                self.card.setFixedSize(self.width() - pad * 2, self.height() - pad * 2)
                self.card.move(pad, pad)
                self.stack.setGeometry(self.card.rect())
                self.dashboard_page._current_reel_index = 0
                self.dashboard_page.reel.move(0, 0)
                self.dashboard_page.refresh_tab_states()
                self.dashboard_page.refresh_live_metrics()
                self.dashboard_page._sync_nav_active()
            else:
                w, h = 420, 580
                cx = (self.width() - w) // 2
                cy = (self.height() - h) // 2
                self.card.setFixedSize(w, h)
                self.card.move(cx, cy)
                self.stack.setGeometry(self.card.rect())

            self.card.setUpdatesEnabled(True)

        hold = QPauseAnimation(80)
        hold.finished.connect(morph)

        fi = QPropertyAnimation(self.opacity, b"opacity")
        fi.setDuration(260)
        fi.setStartValue(0.0)
        fi.setEndValue(1.0)
        fi.setEasingCurve(QEasingCurve.Type.OutCubic)

        def on_finish():
            self.position_layers()
            if target_index == 2:
                # Remove opacity effect so child widgets and subpixel text render with full fidelity
                self.card.setGraphicsEffect(None)

        fi.finished.connect(on_finish)

        self.morph_group.addAnimation(fo)
        self.morph_group.addAnimation(hold)
        self.morph_group.addAnimation(fi)
        self.morph_group.start()

    def showEvent(self, event):
        super().showEvent(event)
        self.position_layers()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._bg_cache = None
        self.position_layers()

    def position_layers(self):
        if hasattr(self, "morph_group") and self.morph_group.state() == self.morph_group.State.Running:
            return
        if self.stack.currentIndex() == 2:
            pad = 12
            self.card.setFixedSize(self.width() - pad * 2, self.height() - pad * 2)
            self.card.move(pad, pad)
        else:
            w, h = 420, 580
            cx = (self.width() - w) // 2
            cy = (self.height() - h) // 2
            self.card.setFixedSize(w, h)
            self.card.move(cx, cy)
        self.stack.setGeometry(self.card.rect())

    def _render_bg_cache(self):
        w, h = max(1, self.width()), max(1, self.height())
        pix = QPixmap(w, h)
        p = QPainter(pix)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.fillRect(QRect(0, 0, w, h), QColor("#050505"))
        diag = math.hypot(float(w), float(h))

        g1 = QRadialGradient(QPointF(w * 0.28, h * 0.40), diag * 0.52)
        g1.setColorAt(0.00, QColor(123, 47, 247, 110))
        g1.setColorAt(0.38, QColor(194, 24, 91, 75))
        g1.setColorAt(0.72, QColor(255, 59, 31, 35))
        g1.setColorAt(1.00, QColor(255, 59, 31, 0))
        p.setBrush(g1)
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRect(QRect(0, 0, w, h))

        g2 = QRadialGradient(QPointF(w * 0.65, h * 0.75), diag * 0.48)
        g2.setColorAt(0.00, QColor(255, 59, 31, 95))
        g2.setColorAt(0.45, QColor(194, 24, 91, 55))
        g2.setColorAt(1.00, QColor(123, 47, 247, 0))
        p.setBrush(g2)
        p.drawRect(QRect(0, 0, w, h))

        g3 = QRadialGradient(QPointF(w * 0.88, h * 0.12), diag * 0.38)
        g3.setColorAt(0.00, QColor(212, 230, 0, 85))
        g3.setColorAt(0.40, QColor(255, 106, 0, 30))
        g3.setColorAt(1.00, QColor(212, 230, 0, 0))
        p.setBrush(g3)
        p.drawRect(QRect(0, 0, w, h))
        p.end()
        return pix

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._bg_cache is None or self._bg_cache.size() != self.size():
            self._bg_cache = self._render_bg_cache()
        p = QPainter(self)
        p.drawPixmap(0, 0, self._bg_cache)


if __name__ == "__main__":
    log.info(f"=== {APP_NAME} v{APP_VERSION} starting ===")

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)

    # Universal stepper filter to ensure all QSpinBox, QDateEdit, and QDoubleSpinBox
    # widgets have a comfortable 14px left padding away from the left edge across the entire app
    _stepper_filter = UniversalSpinBoxFilter(app)
    app.installEventFilter(_stepper_filter)

    # Set dark global palette so native popup windows and frames never render white borders
    dark_pal = QPalette()
    dark_pal.setColor(QPalette.ColorRole.Window, QColor("#141416"))
    dark_pal.setColor(QPalette.ColorRole.WindowText, QColor("#FFFFFF"))
    dark_pal.setColor(QPalette.ColorRole.Base, QColor("#141416"))
    dark_pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#1E1E22"))
    dark_pal.setColor(QPalette.ColorRole.ToolTipBase, QColor("#1E1E22"))
    dark_pal.setColor(QPalette.ColorRole.ToolTipText, QColor("#FFFFFF"))
    dark_pal.setColor(QPalette.ColorRole.Text, QColor("#FFFFFF"))
    dark_pal.setColor(QPalette.ColorRole.Button, QColor("#1E1E22"))
    dark_pal.setColor(QPalette.ColorRole.ButtonText, QColor("#FFFFFF"))
    dark_pal.setColor(QPalette.ColorRole.Highlight, QColor("#2D3142"))
    dark_pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    dark_pal.setColor(QPalette.ColorRole.Light, QColor("#1E1E22"))
    dark_pal.setColor(QPalette.ColorRole.Midlight, QColor("#18181A"))
    dark_pal.setColor(QPalette.ColorRole.Dark, QColor("#0A0A0C"))
    dark_pal.setColor(QPalette.ColorRole.Mid, QColor("#141416"))
    dark_pal.setColor(QPalette.ColorRole.Shadow, QColor("#000000"))
    app.setPalette(dark_pal)

    available_fonts = QFontDatabase.families()
    chosen_family = "Inter"
    for candidate in ["Inter", "Ubuntu", "Cantarell", "Noto Sans", "DejaVu Sans", "SF Pro Display", "Segoe UI"]:
        if candidate in available_fonts:
            chosen_family = candidate
            break
    base_font = QFont(chosen_family, 10)
    app.setFont(base_font)

    if load_data_folder() is None:
        legacy = os.path.join(os.path.expanduser("~"), "Desktop", "APEX")
        if os.path.isdir(legacy) and os.path.exists(os.path.join(legacy, "users.db")):
            save_data_folder(legacy)
        else:
            picker = DataFolderPickerDialog(parent=None)
            if picker.exec() and picker.chosen:
                save_data_folder(picker.chosen)
            else:
                save_data_folder(_default_data_folder())

    init_database()
    auto_backup_database()
    register_pdf_fonts()
    print(f"Repository Root: {get_apex_root()}")
    print("PDF Fonts:", _PDF_FONT_REGISTERED)
    print("Design System: Glassmorphic Edition")

    window = PremiumMainWindow()
    window.show()
    code = app.exec()
    log.info(f"=== exit code {code} ===")
    sys.exit(code)