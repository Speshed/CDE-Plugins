"""Shared Larix CDE theme primitives used by every SOD application."""
from __future__ import annotations

import os
import sys
from pathlib import Path
from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import QEasingCurve, Property, QPropertyAnimation, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap

ACCENT = "#F7921E"
ACCENT_HOVER = "#FFA74B"
ACCENT_PRESSED = "#E07E12"
STATUS_SUCCESS = "#4CAF73"
STATUS_DANGER = "#D94A4A"

def shared_asset_dir(anchor: Path | str | None = None) -> Path:
    """Locate the single shared ``assets`` directory used by all SOD tools."""
    override = os.environ.get("SOD_MANAGER_ASSETS", "").strip()
    if override:
        return Path(override).expanduser().resolve()

    starts = []
    if getattr(sys, "frozen", False):
        starts.append(Path(sys.executable).resolve().parent)
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            starts.append(Path(meipass).resolve())
    if anchor is not None:
        starts.append(Path(anchor).resolve())
    else:
        starts.append(Path(__file__).resolve().parent)

    seen = set()
    for start in starts:
        for base in (start, *start.parents):
            candidate = base / "assets"
            key = str(candidate)
            if key in seen:
                continue
            seen.add(key)
            if candidate.is_dir():
                return candidate

    # Stable fallback for diagnostics even when the folder is missing.
    return starts[0] / "assets"



def asset_path(asset_dir: Path | str, name: str) -> str:
    p = Path(asset_dir) / name
    return str(p.resolve()).replace("\\", "/") if p.exists() else ""


def _tint_pixmap(pm: QPixmap, color: QColor) -> QPixmap:
    if pm.isNull():
        return pm
    tinted = QPixmap(pm.size())
    tinted.fill(Qt.transparent)
    painter = QPainter(tinted)
    painter.drawPixmap(0, 0, pm)
    painter.setCompositionMode(QPainter.CompositionMode_SourceIn)
    painter.fillRect(tinted.rect(), color)
    painter.end()
    return tinted


def themed_icon(asset_dir: Path | str, name: str, dark: bool = False) -> QIcon:
    """Return one consistent icon palette across all CDE tools."""
    path = asset_path(asset_dir, name)
    if not path:
        return QIcon()
    preserve_original = {"excel.png", "krug_galka.png"}
    base_name = Path(str(name)).name.lower()
    if base_name in preserve_original:
        return QIcon(path)
    pm = QPixmap(path)
    if pm.isNull():
        return QIcon(path)
    neutral = QColor("#E7E7E7") if dark else QColor("#303238")
    return QIcon(_tint_pixmap(pm, neutral))



def initial_dark_theme(default: bool = False) -> bool:
    """Initial theme shared with the Larix CDE launcher.

    The launcher passes SOD_MANAGER_THEME to child tools.  When a tool is
    started directly, fall back to the same QSettings value used by the
    launcher so all applications still open in one theme.
    """
    raw = os.environ.get("SOD_MANAGER_THEME", "").strip().lower()
    if raw in {"dark", "1", "true", "yes", "on"}:
        return True
    if raw in {"light", "0", "false", "no", "off"}:
        return False
    try:
        settings = QtCore.QSettings("Larix", "SODManager")
        return settings.value("dark_theme", bool(default), type=bool)
    except Exception:
        return bool(default)


def persist_dark_theme(dark: bool) -> None:
    """Persist a theme change so launcher and other SOD tools stay in sync."""
    value = bool(dark)
    os.environ["SOD_MANAGER_THEME"] = "dark" if value else "light"
    try:
        settings = QtCore.QSettings("Larix", "SODManager")
        settings.setValue("dark_theme", value)
        settings.sync()
    except Exception:
        pass

def apply_windows_titlebar_theme(widget: QtWidgets.QWidget, dark: bool) -> None:
    """Apply the same native Windows caption colors to every Larix CDE window."""
    if sys.platform != "win32" or widget is None:
        return
    try:
        import ctypes
        from ctypes import byref, c_int, sizeof

        hwnd = int(widget.winId())
        value = c_int(1 if dark else 0)
        for attr in (20, 19):
            try:
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, byref(value), sizeof(value)) == 0:
                    break
            except Exception:
                continue

        caption_color = 0x000000 if dark else 0xFFFFFF
        text_color = 0xFFFFFF if dark else 0x000000
        for attr, color in ((35, caption_color), (36, text_color)):
            try:
                cval = c_int(color)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, byref(cval), sizeof(cval))
            except Exception:
                pass
    except Exception:
        pass


class ThemeToggle(QtWidgets.QWidget):
    """Animated Larix CDE theme switch used identically in every SOD tool."""

    toggled = QtCore.Signal(bool) if hasattr(QtCore, "Signal") else QtCore.pyqtSignal(bool)

    def __init__(self, asset_dir, parent=None):
        super().__init__(parent)
        self._asset_dir = Path(asset_dir)
        self._checked = False
        self._handle_progress = 0.0
        self._hovered = False
        self._pressed = False
        self._icon_cache = {}
        self._sun_source = QtGui.QPixmap(asset_path(self._asset_dir, "sun.png"))
        self._moon_source = QtGui.QPixmap(asset_path(self._asset_dir, "moon.png"))

        self._anim = QtCore.QPropertyAnimation(self, b"handleProgress", self)
        self._anim.setDuration(190)
        try:
            self._anim.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        except AttributeError:  # PyQt5 fallback
            self._anim.setEasingCurve(QtCore.QEasingCurve.Type.OutCubic)

        self.setObjectName("themeToggle")
        self.setCursor(Qt.PointingHandCursor)
        self.setFocusPolicy(Qt.NoFocus)
        self.setAttribute(Qt.WA_Hover, True)
        self.setAttribute(Qt.WA_StyledBackground, False)
        self.setAutoFillBackground(False)
        self.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        self.setFixedSize(66, 28)
        # Do not let a parent stylesheet paint a rectangular widget background.
        self.setStyleSheet("QWidget#themeToggle { background: transparent; border: none; }")

    def sizeHint(self):
        return QtCore.QSize(66, 28)

    def minimumSizeHint(self):
        return QtCore.QSize(66, 28)

    def isChecked(self):
        return self._checked

    def setChecked(self, checked: bool, animate: bool = True):
        checked = bool(checked)
        target = 1.0 if checked else 0.0
        if checked == self._checked:
            if not animate:
                self._handle_progress = target
                self.update()
            return

        self._checked = checked
        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(self._handle_progress)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._handle_progress = target
            self.update()
        self.toggled.emit(self._checked)

    def _get_handle_progress(self):
        return self._handle_progress

    def _set_handle_progress(self, value):
        self._handle_progress = max(0.0, min(1.0, float(value)))
        self.update()

    if hasattr(QtCore, "Property"):
        handleProgress = QtCore.Property(float, _get_handle_progress, _set_handle_progress)
    else:  # PyQt5 fallback
        handleProgress = QtCore.pyqtProperty(float, _get_handle_progress, _set_handle_progress)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._pressed = True
            self.update()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            was_pressed = self._pressed
            self._pressed = False
            if was_pressed and self.rect().contains(event.position().toPoint()):
                self.setChecked(not self._checked)
            self.update()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Space, Qt.Key_Return, Qt.Key_Enter):
            self.setChecked(not self._checked)
            event.accept()
            return
        super().keyPressEvent(event)

    def enterEvent(self, event):
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hovered = False
        self._pressed = False
        self.update()
        super().leaveEvent(event)

    def _scaled_icon(self, key, source, size):
        if source.isNull() or size <= 0:
            return QtGui.QPixmap()
        dpr = max(1.0, float(self.devicePixelRatioF()))
        cache_key = (key, int(size), round(dpr, 3))
        cached = self._icon_cache.get(cache_key)
        if cached is not None:
            return cached
        px = max(1, int(round(size * dpr)))
        scaled = source.scaled(px, px, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        scaled.setDevicePixelRatio(dpr)
        self._icon_cache[cache_key] = scaled
        return scaled

    def paintEvent(self, event):
        del event
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)

        track = QtCore.QRectF(self.rect().adjusted(1, 1, -1, -1))
        dark = bool(self._checked)
        radius = track.height() * 0.5

        if dark:
            bg_start = QtGui.QColor("#2b2b2d")
            bg_end = QtGui.QColor("#1d1d1f")
            border_col = QtGui.QColor(255, 255, 255, 28)
        else:
            bg_start = QtGui.QColor("#f2f2f2")
            bg_end = QtGui.QColor("#e7e7e7")
            border_col = QtGui.QColor(0, 0, 0, 24)

        if self._hovered:
            bg_start = bg_start.lighter(105)
            bg_end = bg_end.lighter(105)
        if self._pressed:
            bg_start = bg_start.darker(103)
            bg_end = bg_end.darker(103)

        grad = QtGui.QLinearGradient(track.topLeft(), track.bottomLeft())
        grad.setColorAt(0.0, bg_start)
        grad.setColorAt(1.0, bg_end)
        p.setPen(Qt.NoPen)
        p.setBrush(grad)
        p.drawRoundedRect(track, radius, radius)

        p.setPen(QtGui.QPen(border_col, 1.0))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(track, radius, radius)

        icon_size = int(track.height() * 0.48)
        center_y = track.center().y()
        pad_x = max(6.0, track.height() * 0.24)
        left_x = track.left() + pad_x
        right_x = track.right() - icon_size - pad_x

        sun_pm = self._scaled_icon("sun", self._sun_source, icon_size)
        moon_pm = self._scaled_icon("moon", self._moon_source, icon_size)
        tint = QtGui.QColor(255, 255, 255) if dark else QtGui.QColor(0, 0, 0)
        sun_pm = _tint_pixmap(sun_pm, tint)
        moon_pm = _tint_pixmap(moon_pm, tint)

        highlight_size = icon_size + 9
        highlight_center_x = right_x + icon_size * 0.5 - (right_x - left_x) * self._handle_progress
        p.setPen(Qt.NoPen)
        p.setBrush(QtGui.QColor(ACCENT))
        p.drawEllipse(QtCore.QRectF(
            highlight_center_x - highlight_size * 0.5,
            center_y - highlight_size * 0.5,
            highlight_size,
            highlight_size,
        ))

        if not sun_pm.isNull():
            p.setOpacity(1.0 - 0.62 * self._handle_progress)
            p.drawPixmap(int(right_x), int(center_y - sun_pm.height() / 2), sun_pm)
        if not moon_pm.isNull():
            p.setOpacity(0.38 + 0.62 * self._handle_progress)
            p.drawPixmap(int(left_x), int(center_y - moon_pm.height() / 2), moon_pm)
        p.setOpacity(1.0)
        p.end()


# ---------------------------------------------------------------------------
# Shared UX layer
# ---------------------------------------------------------------------------
def launched_from_manager() -> bool:
    return os.environ.get("SOD_MANAGER_PARENT", "").strip().lower() in {"1", "true", "yes", "on"}


def shared_style_overrides(dark: bool) -> str:
    """Final QSS overrides shared by every CDE tool."""
    bg = "#121212" if dark else "#FFFFFF"
    text = "#E5E5E5" if dark else "#242424"
    muted = "#969696" if dark else "#777777"
    border = "#3A3A3A" if dark else "#DCDCDC"
    soft = "#2A2A2A" if dark else "#EFEFEF"
    hover = "#1B1B1B" if dark else "#F7F7F7"
    pressed = "#24201B" if dark else "#FFF6EB"
    disabled_bg = "#1B1B1B" if dark else "#F2F2F2"
    disabled_text = "#666666" if dark else "#A0A0A0"
    scroll = "rgba(180,180,180,0.26)" if dark else "rgba(80,80,80,0.20)"
    scroll_hover = "rgba(247,146,30,0.72)"
    info_bg = "#171717" if dark else "#FAFAFA"
    selected = "rgba(247,146,30,0.10)"
    selected_border = "rgba(247,146,30,0.58)"
    return f"""
        QPushButton#backToManagerButton {{
            background: transparent; color: {muted}; border: 1px solid {border};
            border-radius: 12px; padding: 6px 11px; font-weight: 600;
        }}
        QPushButton#backToManagerButton:hover {{
            background: {hover}; color: {text}; border-color: {selected_border};
        }}
        QPushButton#backToManagerButton:pressed {{ background: {pressed}; }}

        QTabBar::tab:hover {{ background: {hover}; color: {text}; border-color: {border}; }}
        QTabBar::tab:selected {{
            background: {selected}; color: {text}; border: 1px solid {selected_border};
        }}

        QPushButton#orangeAction, QPushButton#greenAction, QPushButton#greenBtn,
        QPushButton#primaryButton, QPushButton#outlineOrange, QPushButton#outlineGreen,
        QPushButton#purpleAction {{
            background: {bg}; color: {text}; border: 1px solid {border};
        }}
        QPushButton#orangeAction:hover, QPushButton#greenAction:hover, QPushButton#greenBtn:hover,
        QPushButton#primaryButton:hover, QPushButton#outlineOrange:hover, QPushButton#outlineGreen:hover,
        QPushButton#purpleAction:hover {{
            background: {hover}; color: {text}; border-color: {selected_border};
        }}
        QPushButton#orangeAction:pressed, QPushButton#greenAction:pressed, QPushButton#greenBtn:pressed,
        QPushButton#primaryButton:pressed, QPushButton#outlineOrange:pressed, QPushButton#outlineGreen:pressed,
        QPushButton#purpleAction:pressed {{ background: {pressed}; border-color: {selected_border}; }}
        QPushButton#orangeAction:disabled, QPushButton#greenAction:disabled, QPushButton#greenBtn:disabled,
        QPushButton#primaryButton:disabled, QPushButton#outlineOrange:disabled, QPushButton#outlineGreen:disabled,
        QPushButton#purpleAction:disabled {{
            background: {disabled_bg}; color: {disabled_text}; border-color: {soft};
        }}

        QPushButton#destructiveButton {{ background: {bg}; color: {muted}; border: 1px solid {border}; }}
        QPushButton#destructiveButton:hover {{
            background: rgba(217,74,74,0.10); color: #D94A4A; border-color: rgba(217,74,74,0.55);
        }}

        QFrame#infoBanner {{ background: {info_bg}; border: 1px solid {soft}; border-radius: 8px; }}
        QLabel#infoText {{ color: {muted}; background: transparent; }}

        QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; border: none; }}
        QScrollBar::handle:vertical {{ background: {scroll}; min-height: 28px; border: none; border-radius: 4px; }}
        QScrollBar::handle:vertical:hover, QScrollBar::handle:vertical:pressed {{ background: {scroll_hover}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; background: transparent; border: none; }}
        QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
        QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; border: none; }}
        QScrollBar::handle:horizontal {{ background: {scroll}; min-width: 28px; border: none; border-radius: 4px; }}
        QScrollBar::handle:horizontal:hover, QScrollBar::handle:horizontal:pressed {{ background: {scroll_hover}; }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0px; background: transparent; border: none; }}
        QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{ background: transparent; }}

        QStatusBar#unifiedStatusBar {{
            background: {bg}; color: {muted}; border-top: 1px solid {soft}; padding: 0 6px;
        }}
        QStatusBar#unifiedStatusBar::item {{ border: none; }}
        QStatusBar#unifiedStatusBar[tone="busy"] {{ color: #C88935; }}
        QStatusBar#unifiedStatusBar[tone="success"] {{ color: #4CAF73; }}
        QStatusBar#unifiedStatusBar[tone="danger"] {{ color: #D94A4A; }}
        QLabel#unifiedStatusText {{ color: {muted}; background: transparent; }}
        QMessageBox QPushButton {{ min-width: 92px; }}
        QToolTip {{ border: 1px solid {border}; }}
    """


class _WindowStateBinder(QtCore.QObject):
    def __init__(self, window: QtWidgets.QWidget, key: str):
        super().__init__(window)
        self.window = window
        self.key = str(key)
        window.installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj is self.window and event.type() == QtCore.QEvent.Close:
            try:
                settings = QtCore.QSettings("Larix", "SODManager")
                settings.setValue(f"window_geometry_v2/{self.key}", self.window.saveGeometry())
                settings.sync()
            except Exception:
                pass
        return super().eventFilter(obj, event)


def install_window_state(window: QtWidgets.QWidget, key: str) -> None:
    """Restore and persist the last user-chosen size/position for a main window."""
    try:
        settings = QtCore.QSettings("Larix", "SODManager")
        geometry = settings.value(f"window_geometry_v2/{key}")
        if geometry:
            window.restoreGeometry(geometry)
    except Exception:
        pass
    window._larix_window_state_binder = _WindowStateBinder(window, key)
