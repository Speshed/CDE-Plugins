from __future__ import annotations

import os
import sys
from pathlib import Path
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPixmap
from typing import Union


def _load_shared_theme():
    starts = []
    raw_assets = os.environ.get("SOD_MANAGER_ASSETS", "").strip()
    if raw_assets:
        starts.append(Path(raw_assets).expanduser().resolve().parent)
    if getattr(sys, "frozen", False):
        starts.append(Path(sys.executable).resolve().parent)
    starts.append(Path(__file__).resolve().parent)
    for start in starts:
        for base in (start, *start.parents):
            if (base / "shared" / "theme_core.py").is_file():
                if str(base) not in sys.path:
                    sys.path.insert(0, str(base))
                return
    raise ImportError("Larix CDE shared/theme_core.py not found")

_load_shared_theme()
from shared.theme_core import (
    ACCENT, ACCENT_HOVER, ACCENT_PRESSED, STATUS_SUCCESS, STATUS_DANGER,
    ThemeToggle, asset_path, themed_icon, shared_asset_dir, initial_dark_theme,
    persist_dark_theme, _tint_pixmap, apply_windows_titlebar_theme,
)

def build_qss(dark: bool, asset_dir: Union[Path, str]) -> str:
    asset_dir = Path(asset_dir)
    suffix = "-white" if dark else ""
    left = asset_path(asset_dir, f"arrow-left{suffix}.png")
    right = asset_path(asset_dir, f"arrow-right{suffix}.png")
    up = asset_path(asset_dir, f"arrow-up{suffix}.png")
    down = asset_path(asset_dir, f"arrow-down{suffix}.png")
    combo_down = asset_path(asset_dir, f"free-icon-down-arrow-3889508{suffix}.png")
    cb_off = asset_path(asset_dir, f"check{suffix}.png")
    cb_on = asset_path(asset_dir, f"select{suffix}.png")
    cb_mid = asset_path(asset_dir, f"poloska{suffix}.png")

    bg = "#121212" if dark else "#FFFFFF"
    text = "#e0e0e0" if dark else "#222222"
    muted = "#909090" if dark else "#777777"
    border = "#404040" if dark else "#dcdcdc"
    soft_border = "#2f2f2f" if dark else "#eeeeee"
    disabled_bg = "#202020" if dark else "#f0f0f0"
    disabled_text = "#666666" if dark else "#9b9b9b"
    field_bg = "#151515" if dark else "#FFFFFF"
    hover = "rgba(247, 146, 30, 0.10)"
    pressed = "rgba(247, 146, 30, 0.20)"
    popup_selected = "rgba(247, 146, 30, 0.20)"
    hover_solid = "#3a2b1a" if dark else "#FFE3C2"
    row_alt = "#181818" if dark else "#FAFAFA"
    nav_bg = "#191919" if dark else "#F6F6F6"
    status_bg = "#171717" if dark else "#FAFAFA"
    log_bg = "#0d1117" if dark else "#111827"

    return f'''
        * {{ font-family: "Segoe UI", "Arial", sans-serif; color: {text}; font-size: 10pt; outline: none; }}
        QMainWindow, QWidget#centralRoot, QWidget#page, QDialog, QMessageBox {{ background: {bg}; }}
        QScrollArea#mainScroll, QScrollArea#mainScroll > QWidget > QWidget {{ background: {bg}; border: none; }}

        QLabel#pageTitle {{ font-size: 18pt; font-weight: 700; color: {text}; background: transparent; }}
        QLabel#pageSubtitle {{ font-size: 9.5pt; color: {muted}; background: transparent; }}
        QLabel#cardTitle {{ font-size: 10.5pt; font-weight: 700; color: {text}; background: transparent; }}
        QLabel#fieldLabel {{ font-size: 8.6pt; color: {muted}; background: transparent; }}
        QLabel#rowSubtitle, QLabel#hintLabel {{ font-size: 8.6pt; color: {muted}; background: transparent; }}
        QLabel#authStatus, QLabel#goodStatus, QLabel#badStatus {{ font-size: 8.6pt; font-weight: 600; background: transparent; }}
        QLabel#goodStatus {{ color: {STATUS_SUCCESS}; }}
        QLabel#badStatus {{ color: {STATUS_DANGER}; }}
        QLabel#inlineIcon {{ background: transparent; border: none; }}

        QFrame#card {{ background: {bg}; border: 1px solid {border}; border-radius: 12px; }}
        QFrame#servicePanel {{ background: {status_bg}; border: 1px solid {soft_border}; border-radius: 9px; }}
        QFrame#separator {{ border: none; border-top: 1px solid {soft_border}; max-height: 1px; }}
        QFrame#infoBanner {{ background: {hover}; border: 1px solid {ACCENT_HOVER}; border-radius: 8px; }}
        QFrame#rightsLegendBox {{ background: {status_bg}; border: 1px solid {soft_border}; border-radius: 8px; }}
        QLabel#infoText, QLabel#rightsLegend {{ font-size: 8.5pt; color: {text}; background: transparent; }}

        QLineEdit, QTextEdit, QComboBox, QPlainTextEdit {{
            background: {field_bg}; color: {text}; border: 1px solid {border}; border-radius: 8px;
            selection-background-color: {ACCENT}; selection-color: #000000;
        }}
        QLineEdit {{ min-height: 34px; padding: 0 9px; }}
        QTextEdit, QPlainTextEdit {{ padding: 6px 8px; }}
        QComboBox {{ min-height: 34px; padding: 0 28px 0 10px; }}
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border-color: {ACCENT_HOVER}; }}
        QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled {{ background: {disabled_bg}; color: {disabled_text}; }}
        QComboBox::drop-down {{ width: 20px; border: none; background: transparent; }}
        QComboBox::down-arrow {{ image: url("{combo_down}"); width: 12px; height: 12px; }}
        QComboBox QAbstractItemView {{
            background: {bg}; color: {text}; border: 1px solid {border}; border-radius: 0px;
            padding: 4px; selection-background-color: transparent; selection-color: {text};
        }}
        QComboBox QAbstractItemView::item {{ padding: 6px 10px; margin: 2px; border: none; border-radius: 6px; background: {bg}; color: {text}; }}
        QComboBox QAbstractItemView::item:hover {{ background: {hover_solid}; }}
        QComboBox QAbstractItemView::item:selected {{ background: {popup_selected}; }}

        QPushButton, QToolButton {{
            background: {bg}; color: {text}; border: 1px solid {border}; border-radius: 14px;
            padding: 6px 12px; font-weight: 600; min-height: 22px;
        }}
        QPushButton:hover, QToolButton:hover {{ background: {hover}; border-color: {ACCENT_HOVER}; }}
        QPushButton:pressed, QToolButton:pressed {{ background: {pressed}; border-color: {ACCENT_PRESSED}; }}
        QPushButton:disabled, QToolButton:disabled {{ background: {disabled_bg}; color: {disabled_text}; border-color: {soft_border}; }}
        QPushButton#loginButton {{
            background: {"#342719" if dark else "#FFF3E8"};
            color: {"#F5B46B" if dark else "#A85C08"};
            border: 1px solid {"#7A4A18" if dark else "#F3B66E"};
            font-weight: 700;
        }}
        QPushButton#loginButton:hover {{
            background: {"#402D19" if dark else "#FFE8D1"};
            color: {"#FFC47E" if dark else "#8F4B00"};
            border-color: {ACCENT};
        }}
        QPushButton#loginButton:pressed {{
            background: {"#2D2116" if dark else "#FFD9B5"};
            border-color: {ACCENT_PRESSED};
        }}
        QPushButton#loginButton:disabled {{ background: {disabled_bg}; color: {disabled_text}; border-color: {soft_border}; }}
        /* Primary actions follow the Larix rule: neutral at rest; orange is only an interaction accent. */
        QPushButton#orangeAction {{
            background: {bg}; color: {text}; border: 1px solid {border}; font-weight: 700;
        }}
        QPushButton#orangeAction:hover {{ background: {hover}; border-color: {ACCENT_HOVER}; color: {text}; }}
        QPushButton#orangeAction:pressed {{ background: {pressed}; border-color: {ACCENT_PRESSED}; color: {text}; }}
        QPushButton#orangeAction:disabled {{ background: {disabled_bg}; color: {disabled_text}; border-color: {soft_border}; }}
        QPushButton#secondaryAction {{ background: {bg}; border-color: {border}; }}
        QPushButton#downloadButton {{ background: {bg}; border-color: {border}; }}
        QPushButton#logButton {{ border: none; background: transparent; color: {muted}; padding: 4px 7px; font-weight: 500; }}
        QPushButton#logButton:hover {{ color: {text}; background: {hover}; border: none; }}
        QPushButton#serviceToggle {{ border: none; background: transparent; color: {muted}; padding: 3px 4px; font-weight: 600; text-align: left; }}
        QPushButton#serviceToggle:hover {{ color: {text}; background: transparent; border: none; }}
        QToolButton#passwordEye, QToolButton#eyeButton {{ background: transparent; border: none; border-radius: 8px; padding: 4px; }}
        QToolButton#passwordEye:hover, QToolButton#eyeButton:hover {{ background: {hover}; border: none; }}

        QTabWidget#workspaceTabs {{ background: transparent; border: none; }}
        QTabWidget#workspaceTabs::pane {{ border: none; background: transparent; top: 10px; }}
        QTabWidget#workspaceTabs QTabBar {{ background: transparent; border: none; padding: 0; }}
        QTabBar::base {{ background: transparent; border: none; height: 0px; }}
        QTabBar::tab {{
            background: {nav_bg}; color: {muted}; border: 1px solid {soft_border}; border-radius: 12px;
            min-height: 28px; padding: 10px 18px; margin: 0 6px; font-weight: 600;
        }}
        QTabBar::tab:first {{ margin-left: 0; }}
        QTabBar::tab:last {{ margin-right: 0; }}
        QTabBar::tab:hover {{ background: {hover}; border-color: {ACCENT_HOVER}; color: {text}; }}
        QTabBar::tab:selected {{ background: {hover}; color: {text}; border: 1px solid {ACCENT_HOVER}; }}
        QWidget#tabPage {{ background: transparent; }}

        QTableWidget#matrixTable {{
            background: {bg}; alternate-background-color: {row_alt}; color: {text}; border: 1px solid {border};
            border-radius: 8px; gridline-color: transparent; selection-background-color: {popup_selected}; selection-color: {text};
        }}
        QTableWidget#matrixTable::item {{ padding: 5px 7px; border: none; border-bottom: 1px solid {soft_border}; }}
        QTableWidget#matrixTable::item:hover {{ background: {hover}; }}
        QTableWidget#matrixTable QHeaderView::section {{
            background: {bg}; color: {muted}; border: none; border-bottom: 1px solid {border};
            padding: 7px; font-weight: 600;
        }}

        QCheckBox {{ background: transparent; padding: 2px; spacing: 7px; }}
        QCheckBox::indicator {{ width: 18px; height: 18px; }}
        QCheckBox::indicator:unchecked {{ image: url("{cb_off}"); }}
        QCheckBox::indicator:checked {{ image: url("{cb_on}"); }}
        QCheckBox::indicator:indeterminate {{ image: url("{cb_mid}"); }}

        QTextEdit#logArea {{ background: {log_bg}; color: #e5e7eb; border: 1px solid {border}; border-radius: 10px; padding: 9px; }}
        QStatusBar {{ background: {bg}; color: {muted}; border-top: 1px solid {soft_border}; min-height: 24px; }}
        QStatusBar::item {{ border: none; }}
        QFrame#authOverlay {{ background: {bg}; border: none; }}
        QLabel#authTitle {{ font-size: 18pt; font-weight: 700; color: {text}; }}

        QScrollBar:vertical {{ background: {bg}; width: 12px; margin: 16px 0 16px 0; border: none; }}
        QScrollBar::handle:vertical {{ background: rgba(247,146,30,0.12); min-height: 24px; border-radius: 6px; border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ background: {bg}; height: 16px; subcontrol-origin: margin; border: none; }}
        QScrollBar::add-line:vertical {{ subcontrol-position: bottom; }}
        QScrollBar::sub-line:vertical {{ subcontrol-position: top; }}
        QScrollBar::up-arrow:vertical {{ image: url("{up}"); width: 12px; height: 12px; }}
        QScrollBar::down-arrow:vertical {{ image: url("{down}"); width: 12px; height: 12px; }}
        QScrollBar:horizontal {{ background: {bg}; height: 12px; margin: 0 16px 0 16px; border: none; }}
        QScrollBar::handle:horizontal {{ background: rgba(247,146,30,0.12); min-width: 24px; border-radius: 6px; border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ background: {bg}; width: 16px; subcontrol-origin: margin; border: none; }}
        QScrollBar::add-line:horizontal {{ subcontrol-position: right; }}
        QScrollBar::sub-line:horizontal {{ subcontrol-position: left; }}
        QScrollBar::left-arrow:horizontal {{ image: url("{left}"); width: 12px; height: 12px; }}
        QScrollBar::right-arrow:horizontal {{ image: url("{right}"); width: 12px; height: 12px; }}

        QToolTip {{
            background-color: {'rgba(32,32,32,0.96)' if dark else 'rgba(255,255,255,0.97)'};
            color: {text}; border: 1px solid rgba(247,146,30,0.45); border-radius: 8px;
            padding: 7px 10px; font-size: 9pt;
        }}
    '''
