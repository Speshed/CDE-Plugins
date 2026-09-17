from __future__ import annotations

import os
import sys
from pathlib import Path
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPixmap


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

UI_ASSET_DIR = shared_asset_dir(Path(__file__).resolve().parent)

def build_qss(dark: bool, asset_dir: Path | str = UI_ASSET_DIR) -> str:
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
    field_bg = "#121212" if dark else "#FFFFFF"
    popup_selected = "rgba(247, 146, 30, 0.20)"
    hover = "rgba(247, 146, 30, 0.10)"
    pressed = "rgba(247, 146, 30, 0.20)"
    hover_solid = "#3a2b1a" if dark else "#FFE3C2"
    row_alt = "#181818" if dark else "#FAFAFA"
    nav_bg = "#191919" if dark else "#F6F6F6"
    selected_tab_bg = "#252525" if dark else "#FFFFFF"
    success_bg = "rgba(60, 170, 105, 0.10)"
    success_border = STATUS_SUCCESS
    error_bg = "rgba(217, 74, 74, 0.10)"
    error_border = STATUS_DANGER

    return f'''
        * {{ font-family: "Segoe UI", "Arial", sans-serif; color: {text}; font-size: 10pt; outline: none; }}
        QMainWindow, QWidget#centralRoot, QWidget#page, QWidget#tabPage, QDialog, QMessageBox {{ background: {bg}; }}
        QScrollArea, QScrollArea > QWidget > QWidget {{ background: {bg}; border: none; }}

        QLabel#pageTitle {{ font-size: 18pt; font-weight: 700; color: {text}; background: transparent; }}
        QLabel#pageSubtitle {{ font-size: 9.5pt; color: {muted}; background: transparent; }}
        QLabel#brandLabel {{ font-size: 10pt; font-weight: 700; color: {ACCENT}; background: transparent; }}
        QLabel#cardTitle {{ font-size: 10.5pt; font-weight: 700; color: {text}; background: transparent; }}
        QLabel#rowTitle {{ font-size: 9.5pt; font-weight: 700; background: transparent; }}
        QLabel#rowSubtitle, QLabel#fieldLabel {{ font-size: 8.6pt; color: {muted}; background: transparent; }}
        QLabel#authStatus {{ font-size: 8.6pt; font-weight: 600; background: transparent; }}
        QLabel#authStatus[tone="pending"] {{ color: {ACCENT}; }}
        QLabel#authStatus[tone="success"] {{ color: {STATUS_SUCCESS}; }}
        QLabel#authStatus[tone="danger"] {{ color: {STATUS_DANGER}; }}
        QLabel#authStatus[tone="neutral"] {{ color: {muted}; }}

        QFrame#card {{ background: {bg}; border: 1px solid {border}; border-radius: 12px; }}
        QFrame#separator {{ border: none; border-top: 1px solid {soft_border}; max-height: 1px; margin-top: 8px; }}
        QWidget#fileRow, QWidget#fileRowLast {{ background: transparent; border: none; }}
        QWidget#fileRow {{ border-bottom: 1px solid {soft_border}; }}
        QLabel#iconBadge {{ background: transparent; color: {muted}; border: 1px solid {border}; border-radius: 8px; font-weight: 700; }}
        QLabel#iconBadge[tone="orange"], QLabel#iconBadge[tone="purple"], QLabel#iconBadge[tone="green"] {{ color: {ACCENT}; }}
        QLabel#iconBadge[tone="info"] {{ color: {ACCENT}; }}

        QGroupBox {{
            background: {bg}; color: {text}; border: 1px solid {border}; border-radius: 12px;
            margin-top: 13px; padding: 13px 11px 10px 11px; font-weight: 700;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; subcontrol-position: top left; left: 11px;
            padding: 0 6px; color: {text}; background: {bg};
        }}

        QLineEdit, QTextEdit, QComboBox, QPlainTextEdit, QListWidget {{
            background: {field_bg}; color: {text}; border: 1px solid {border}; border-radius: 8px;
            selection-background-color: {ACCENT}; selection-color: #000000;
        }}
        QLineEdit {{ min-height: 34px; padding: 0 8px; }}
        QTextEdit, QPlainTextEdit {{ padding: 6px 8px; }}
        QListWidget {{ padding: 4px; }}
        QListWidget::item {{ padding: 5px 8px; border-radius: 6px; }}
        QListWidget::item:hover {{ background: {hover}; }}
        QListWidget::item:selected {{ background: {popup_selected}; color: {text}; }}
        QComboBox {{ min-height: 34px; padding: 0 28px 0 10px; }}
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QPlainTextEdit:focus, QListWidget:focus {{ border-color: {'#606060' if dark else '#C9C9C9'}; }}
        QLineEdit:disabled, QTextEdit:disabled, QComboBox:disabled, QListWidget:disabled {{ background: {disabled_bg}; color: {disabled_text}; }}
        QComboBox::drop-down {{ width: 20px; border: none; background: transparent; }}
        QComboBox::down-arrow {{ image: url("{combo_down}"); width: 12px; height: 12px; }}
        /* Popup без внешнего border-radius: у стандартного Qt popup за
           скруглёнными углами мог просвечивать белый native-container. */
        QComboBox QAbstractItemView, QListView#comboPopupView {{
            background: {bg}; color: {text}; border: 1px solid {border}; border-radius: 0px;
            padding: 4px; margin: 0px; selection-background-color: transparent; selection-color: {text};
        }}
        QListView#comboPopupView QWidget#comboPopupViewport {{ background: {bg}; color: {text}; }}
        QComboBoxPrivateContainer {{ background: {bg}; border: none; padding: 0px; margin: 0px; }}
        QComboBox QAbstractItemView::item, QListView#comboPopupView::item {{
            padding: 6px 10px; margin: 2px; border: none; border-radius: 6px; background: {bg}; color: {text};
        }}
        QComboBox QAbstractItemView::item:hover, QListView#comboPopupView::item:hover {{
            background: {hover_solid}; color: {text};
        }}
        QComboBox QAbstractItemView::item:selected, QListView#comboPopupView::item:selected {{
            background: {popup_selected}; color: {text};
        }}
        QComboBox QAbstractItemView::item:selected:hover, QListView#comboPopupView::item:selected:hover {{
            background: rgba(247, 146, 30, 0.28); color: {text};
        }}

        QPushButton, QToolButton {{
            background: {bg}; color: {text}; border: 1px solid {border}; border-radius: 12px;
            min-height: 32px; padding: 4px 12px; font-weight: 600;
        }}
        QPushButton:hover, QToolButton:hover {{ background: {hover}; border-color: {ACCENT_HOVER}; }}
        QPushButton:pressed, QToolButton:pressed {{ background: {pressed}; border-color: {ACCENT_PRESSED}; }}
        QPushButton:disabled, QToolButton:disabled {{ background: {disabled_bg}; color: {disabled_text}; border-color: {soft_border}; }}
        QPushButton#modeSwitch:checked {{ background: {pressed}; border-color: {ACCENT_HOVER}; color: {text}; }}
        QPushButton#secondaryBtn {{ background: {bg}; color: {text}; }}
        /* Main actions stay neutral at rest. Orange is an interaction/selection accent,
           not a permanent border around every primary button. */
        QPushButton#greenBtn, QPushButton#greenAction {{
            background: {bg}; border-color: {border}; color: {text}; font-weight: 700;
        }}
        QPushButton#greenBtn:hover, QPushButton#greenAction:hover {{
            background: {hover}; border-color: {ACCENT_HOVER};
        }}
        QPushButton#greenBtn:pressed, QPushButton#greenAction:pressed {{
            background: {pressed}; border-color: {ACCENT_PRESSED};
        }}
        QPushButton#greenBtn:disabled, QPushButton#greenAction:disabled {{
            background: {disabled_bg}; color: {disabled_text}; border-color: {soft_border};
        }}
        QPushButton#downloadButton {{ min-width: 148px; }}
        QPushButton#uploadButton {{ min-width: 148px; }}
        QPushButton#loginButton {{ min-height: 34px; }}
        QPushButton#logBtn {{ border: none; background: transparent; color: {muted}; padding: 4px 8px; font-weight: 500; }}
        QPushButton#logBtn:hover {{ color: {text}; background: {hover}; border: none; }}
        QToolButton#passwordEye {{ background: transparent; border: none; border-radius: 8px; padding: 4px; }}
        QToolButton#passwordEye:hover {{ background: {hover}; border: none; }}
        QToolButton#passwordEye:pressed {{ background: {pressed}; border: none; }}
        QLabel#inlineIcon {{ background: transparent; border: none; }}
        QPushButton#projectModeSwitch {{
            background: {bg}; color: {muted}; border: 1px solid {border}; border-radius: 14px;
            padding: 8px 14px; font-weight: 600; text-align: center;
        }}
        QPushButton#projectModeSwitch:hover {{
            background: {hover}; color: {text}; border-color: {ACCENT_HOVER};
        }}
        QPushButton#projectModeSwitch:checked {{
            background: {hover}; color: {text}; border: 1px solid {ACCENT_HOVER};
        }}
        QPushButton#logButton {{ border: none; background: transparent; padding: 0 4px; color: {muted}; font-weight: 500; }}
        QPushButton#logButton:hover {{ color: {text}; background: transparent; border: none; }}

        QFrame#infoBanner {{ background: {hover}; border: 1px solid {ACCENT_HOVER}; border-radius: 8px; }}
        QLabel#infoText {{ font-size: 8.5pt; color: {text}; background: transparent; }}
        QFrame#logBar {{ min-height: 38px; background: {bg}; border: 1px solid {border}; border-radius: 10px; }}
        QLabel#chevron {{ color: {muted}; background: transparent; }}

        QFrame#rightsLegendBox {{
            background: {'#181818' if dark else '#FAFAFA'}; border: 1px solid {soft_border}; border-radius: 8px;
        }}
        QLabel#rightsLegend {{ background: transparent; color: {text}; border: none; padding: 0; font-size: 8.5pt; }}
        QLabel#statusChip {{ padding: 4px 10px; border-radius: 10px; font-size: 8.4pt; font-weight: 600; }}
        QLabel#statusChip[tone="pending"] {{ background: {hover}; border: 1px solid {ACCENT_HOVER}; color: {text}; }}
        QLabel#statusChip[tone="success"] {{ background: {success_bg}; border: 1px solid {success_border}; color: {text}; }}
        QLabel#statusChip[tone="danger"] {{ background: {error_bg}; border: 1px solid {error_border}; color: {text}; }}
        QLabel#statusChip[tone="neutral"] {{ background: {row_alt}; border: 1px solid {border}; color: {muted}; }}

        QScrollArea#workspaceScrollArea {{ background: transparent; border: none; }}
        QScrollArea#workspaceScrollArea > QWidget > QWidget {{ background: {bg}; }}
        QTabWidget#workspaceTabs {{ background: transparent; border: none; }}
        QTabWidget#workspaceTabs::pane {{ border: none; background: transparent; top: 6px; }}
        QTabWidget#workspaceTabs QTabBar {{
            background: transparent; border: none; padding: 0;
        }}
        QTabBar::base {{ background: transparent; border: none; height: 0px; }}
        QTabBar::tab {{
            background: {nav_bg}; color: {muted}; border: 1px solid {soft_border}; border-radius: 12px;
            min-height: 26px; padding: 8px 13px; margin: 0 4px; font-weight: 600;
        }}
        QTabBar::tab:first {{ margin-left: 0; }}
        QTabBar::tab:last {{ margin-right: 0; }}
        QTabBar::tab:hover {{ background: {hover}; border-color: {ACCENT_HOVER}; color: {text}; }}
        QTabBar::tab:selected {{
            background: {hover}; color: {text}; border: 1px solid {ACCENT_HOVER};
        }}
        QWidget#tabPage {{ background: transparent; }}

        QTableWidget#matrixTable {{
            background: {bg}; alternate-background-color: {row_alt}; color: {text}; border: 1px solid {border};
            border-radius: 8px; gridline-color: transparent; selection-background-color: {popup_selected}; selection-color: {text};
        }}
        QTableWidget#matrixTable::item {{ padding: 5px 7px; border: none; border-bottom: 1px solid {border}; }}
        QTableWidget#matrixTable::item:hover {{ background: {hover}; }}
        QTableWidget#matrixTable::item:selected {{ background: {popup_selected}; }}
        QTableWidget#matrixTable QHeaderView::section {{
            background: {bg}; color: {muted}; border: none; border-bottom: 1px solid {border};
            padding: 6px; font-weight: 600;
        }}


        QTabWidget#previewTabs {{ background: transparent; border: none; }}
        QTabWidget#previewTabs::pane {{ border: 1px solid {border}; border-radius: 8px; background: {bg}; top: 4px; }}
        QTabWidget#previewTabs QTabBar::tab {{ min-height: 22px; padding: 6px 12px; margin: 0 3px; border-radius: 9px; }}

        QTextEdit#logArea {{
            background: {field_bg}; color: {text}; border: 1px solid {border}; border-radius: 10px;
            padding: 10px; selection-background-color: {popup_selected}; selection-color: {text};
        }}

        QCheckBox {{ background: transparent; padding: 2px; spacing: 7px; }}
        QCheckBox::indicator {{ width: 18px; height: 18px; }}
        QCheckBox::indicator:unchecked {{ image: url("{cb_off}"); }}
        QCheckBox::indicator:checked {{ image: url("{cb_on}"); }}
        QCheckBox::indicator:indeterminate {{ image: url("{cb_mid}"); }}

        QScrollBar:vertical {{ background: {bg}; width: 12px; margin: 16px 0 16px 0; border: none; }}
        QScrollBar::handle:vertical {{ background: rgba(247, 146, 30, 0.12); min-height: 24px; border-radius: 6px; border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::handle:vertical:hover {{ background: rgba(247, 146, 30, 0.15); border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::handle:vertical:pressed {{ background: rgba(247, 146, 30, 0.25); border: 1px solid {ACCENT_PRESSED}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ background: {bg}; height: 16px; subcontrol-origin: margin; border: none; image: none; }}
        QScrollBar::add-line:vertical {{ subcontrol-position: bottom; }}
        QScrollBar::sub-line:vertical {{ subcontrol-position: top; }}
        QScrollBar::add-line:vertical:hover, QScrollBar::sub-line:vertical:hover {{ background: rgba(247, 146, 30, 0.15); }}
        QScrollBar::add-line:vertical:pressed, QScrollBar::sub-line:vertical:pressed {{ background: rgba(247, 146, 30, 0.25); }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: {bg}; }}
        QScrollBar:horizontal {{ background: {bg}; height: 12px; margin: 0 16px 0 16px; border: none; }}
        QScrollBar::handle:horizontal {{ background: rgba(247, 146, 30, 0.12); min-width: 24px; border-radius: 6px; border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::handle:horizontal:hover {{ background: rgba(247, 146, 30, 0.15); border: 1px solid {ACCENT_HOVER}; }}
        QScrollBar::handle:horizontal:pressed {{ background: rgba(247, 146, 30, 0.25); border: 1px solid {ACCENT_PRESSED}; }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ background: {bg}; width: 16px; subcontrol-origin: margin; border: none; image: none; }}
        QScrollBar::add-line:horizontal {{ subcontrol-position: right; }}
        QScrollBar::sub-line:horizontal {{ subcontrol-position: left; }}
        QScrollBar::add-line:horizontal:hover, QScrollBar::sub-line:horizontal:hover {{ background: rgba(247, 146, 30, 0.15); }}
        QScrollBar::add-line:horizontal:pressed, QScrollBar::sub-line:horizontal:pressed {{ background: rgba(247, 146, 30, 0.25); }}
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: {bg}; }}
        QScrollBar::left-arrow:horizontal {{ image: url("{left}"); width: 12px; height: 12px; }}
        QScrollBar::right-arrow:horizontal {{ image: url("{right}"); width: 12px; height: 12px; }}
        QScrollBar::up-arrow:vertical {{ image: url("{up}"); width: 12px; height: 12px; }}
        QScrollBar::down-arrow:vertical {{ image: url("{down}"); width: 12px; height: 12px; }}

        QToolTip {{
            background-color: {'rgba(32,32,32,0.96)' if dark else 'rgba(255,255,255,0.95)'};
            color: {text}; border: 1px solid rgba(247,146,30,0.4); border-radius: 10px;
            padding: 8px 12px; font-size: 12px;
        }}
    '''


STYLESHEET = build_qss(False)

__all__ = [
    "ACCENT", "ACCENT_HOVER", "ACCENT_PRESSED", "STATUS_SUCCESS", "STATUS_DANGER",
    "UI_ASSET_DIR", "ThemeToggle", "asset_path", "themed_icon", "build_qss", "STYLESHEET",
    "initial_dark_theme", "persist_dark_theme",
]
