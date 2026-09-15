# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from app_registry import APPS, AppSpec
from shared.theme_core import install_window_state

APP_TITLE = "Larix CDE"
APP_VERSION = "1.0.0"


def application_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = application_root()
ASSET_DIR = ROOT / "assets"


LIGHT = {
    "window": "#F6F5F4",
    "surface": "#FEFEFE",
    "text": "#21252E",
    "muted": "#73819B",
    "muted2": "#98A2B3",
    "border": "#D9DDE5",
    "card": "rgba(255, 255, 255, 230)",
    "card_border": "rgba(212, 216, 224, 235)",
    "card_hover": "rgba(255, 255, 255, 245)",
    "card_pressed": "rgba(255, 247, 238, 250)",
    "accent": "#F7921E",
    "accent_soft": "rgba(247, 146, 30, 36)",
    "arrow_bg": "rgba(247, 146, 30, 0.12)",
    "arrow_fg": "#F7921E",
    "arrow_border": "rgba(247, 146, 30, 0.18)",
    "switch_track": "#ECEEF3",
    "switch_border": "#D6D9E1",
    "switch_icon": "#7A7F89",
    "hint": "#7B879C",
    "card_title": "#1F2633",
    "card_subtitle": "#8A93A3",
}

DARK = {
    "window": "#050505",
    "surface": "#0A0A0A",
    "text": "#F4F4F4",
    "muted": "#B0B0B0",
    "muted2": "#898989",
    "border": "rgba(255, 255, 255, 0.12)",
    "card": "rgba(17, 17, 17, 220)",
    "card_border": "rgba(255, 255, 255, 0.16)",
    "card_hover": "rgba(25, 25, 25, 235)",
    "card_pressed": "rgba(30, 30, 30, 245)",
    "accent": "#F7921E",
    "accent_soft": "rgba(247, 146, 30, 52)",
    "arrow_bg": "rgba(255, 255, 255, 0.06)",
    "arrow_fg": "#EEEEEE",
    "arrow_border": "rgba(255, 255, 255, 0.12)",
    "switch_track": "rgba(255, 255, 255, 0.05)",
    "switch_border": "rgba(255, 255, 255, 0.10)",
    "switch_icon": "#EEEEEE",
    "hint": "#B5B5B5",
    "card_title": "#F4F4F4",
    "card_subtitle": "#A6A6A6",
}


def theme(dark: bool) -> dict:
    return DARK if dark else LIGHT


class BackgroundWidget(QtWidgets.QWidget):
    def __init__(self, dark: bool = False, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self._dark = dark

    def set_dark(self, dark: bool) -> None:
        self._dark = dark
        self.update()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        c = theme(self._dark)
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = self.rect()

        p.fillRect(rect, QtGui.QColor(c["window"]))

        # Main vertical wash
        grad = QtGui.QLinearGradient(0, 0, 0, rect.height())
        if self._dark:
            grad.setColorAt(0.0, QtGui.QColor("#090909"))
            grad.setColorAt(0.55, QtGui.QColor("#030303"))
            grad.setColorAt(1.0, QtGui.QColor("#080808"))
        else:
            grad.setColorAt(0.0, QtGui.QColor("#F7F6F5"))
            grad.setColorAt(1.0, QtGui.QColor("#F5F5F4"))
        p.fillRect(rect, grad)

        # Soft glows
        glow = QtGui.QRadialGradient(QtCore.QPointF(70, rect.height() - 60), rect.width() * 0.42)
        if self._dark:
            glow.setColorAt(0.0, QtGui.QColor(247, 146, 30, 70))
            glow.setColorAt(0.55, QtGui.QColor(247, 146, 30, 16))
            glow.setColorAt(1.0, QtGui.QColor(247, 146, 30, 0))
        else:
            glow.setColorAt(0.0, QtGui.QColor(247, 146, 30, 34))
            glow.setColorAt(0.60, QtGui.QColor(247, 146, 30, 8))
            glow.setColorAt(1.0, QtGui.QColor(247, 146, 30, 0))
        p.fillRect(rect, glow)

        glow2 = QtGui.QRadialGradient(QtCore.QPointF(rect.width() * 0.88, rect.height() * 0.18), rect.width() * 0.35)
        if self._dark:
            glow2.setColorAt(0.0, QtGui.QColor(255, 255, 255, 12))
            glow2.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
        else:
            glow2.setColorAt(0.0, QtGui.QColor(255, 255, 255, 0))
            glow2.setColorAt(1.0, QtGui.QColor(255, 255, 255, 0))
        p.fillRect(rect, glow2)

        # Decorative wave lines
        pen_color = QtGui.QColor(255, 255, 255, 30) if self._dark else QtGui.QColor(237, 221, 204, 110)
        pen = QtGui.QPen(pen_color, 1.2)
        p.setPen(pen)
        for i in range(7):
            offset = i * 16
            path = QtGui.QPainterPath()
            path.moveTo(0, 118 + offset)
            path.cubicTo(rect.width() * 0.08, 25 + offset, rect.width() * 0.18, 22 + offset, rect.width() * 0.31, 0 + offset * 0.14)
            p.drawPath(path)

            path2 = QtGui.QPainterPath()
            path2.moveTo(rect.width(), rect.height() - (118 + offset))
            path2.cubicTo(rect.width() * 0.92, rect.height() - (25 + offset), rect.width() * 0.82, rect.height() - (22 + offset), rect.width() * 0.69, rect.height() - offset * 0.14)
            p.drawPath(path2)


class ThemeSwitch(QtWidgets.QWidget):
    """Animated Larix-style theme toggle. checked=True means dark theme."""

    toggled = QtCore.Signal(bool)

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self._checked = False
        self._handle_progress = 0.0
        self._hovered = False
        self._pressed = False
        self._icon_cache: dict[tuple[str, int, float], QtGui.QPixmap] = {}
        self._sun_source = QtGui.QPixmap(str(ASSET_DIR / "sun.png"))
        self._moon_source = QtGui.QPixmap(str(ASSET_DIR / "moon.png"))

        self._anim = QtCore.QPropertyAnimation(self, b"handleProgress", self)
        self._anim.setDuration(190)
        self._anim.setEasingCurve(QtCore.QEasingCurve.OutCubic)

        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFocusPolicy(QtCore.Qt.NoFocus)
        self.setAttribute(QtCore.Qt.WA_Hover, True)
        self.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        self.setFixedSize(66, 28)

    def sizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(66, 28)

    def minimumSizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(66, 28)

    def isChecked(self) -> bool:
        return self._checked

    def setChecked(self, checked: bool, animate: bool = True, emit_signal: bool = True) -> None:
        checked = bool(checked)
        if self._checked == checked:
            target = 1.0 if checked else 0.0
            if not animate:
                self._handle_progress = target
                self.update()
            return

        self._checked = checked
        target = 1.0 if checked else 0.0
        if animate and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(self._handle_progress)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._handle_progress = target
            self.update()

        if emit_signal:
            self.toggled.emit(self._checked)

    def _get_handle_progress(self) -> float:
        return self._handle_progress

    def _set_handle_progress(self, value: float) -> None:
        self._handle_progress = max(0.0, min(1.0, float(value)))
        self.update()

    handleProgress = QtCore.Property(float, _get_handle_progress, _set_handle_progress)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.LeftButton:
            self._pressed = True
            self.update()
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.LeftButton:
            was_pressed = self._pressed
            self._pressed = False
            if was_pressed and self.rect().contains(event.position().toPoint()):
                self.setChecked(not self._checked)
            self.update()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if event.key() in (QtCore.Qt.Key_Space, QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
            self.setChecked(not self._checked)
            event.accept()
            return
        super().keyPressEvent(event)

    def enterEvent(self, event: QtCore.QEvent) -> None:
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event: QtCore.QEvent) -> None:
        self._hovered = False
        self._pressed = False
        self.update()
        super().leaveEvent(event)

    def _scaled_icon(self, key: str, source: QtGui.QPixmap, size: int) -> QtGui.QPixmap:
        if source.isNull() or size <= 0:
            return QtGui.QPixmap()
        dpr = max(1.0, self.devicePixelRatioF())
        cache_key = (key, size, dpr)
        cached = self._icon_cache.get(cache_key)
        if cached is not None:
            return cached
        px = max(1, int(round(size * dpr)))
        scaled = source.scaled(px, px, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        scaled.setDevicePixelRatio(dpr)
        self._icon_cache[cache_key] = scaled
        return scaled

    @staticmethod
    def _tint_pixmap(pm: QtGui.QPixmap, color: QtGui.QColor) -> QtGui.QPixmap:
        if pm.isNull():
            return pm
        tinted = QtGui.QPixmap(pm.size())
        tinted.setDevicePixelRatio(pm.devicePixelRatio())
        tinted.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(tinted)
        painter.drawPixmap(0, 0, pm)
        painter.setCompositionMode(QtGui.QPainter.CompositionMode_SourceIn)
        painter.fillRect(tinted.rect(), color)
        painter.end()
        return tinted

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
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
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(grad)
        p.drawRoundedRect(track, radius, radius)

        p.setPen(QtGui.QPen(border_col, 1.0))
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawRoundedRect(track, radius, radius)

        icon_size = int(track.height() * 0.48)
        center_y = track.center().y()
        pad_x = max(6.0, track.height() * 0.24)
        left_x = track.left() + pad_x
        right_x = track.right() - icon_size - pad_x

        sun_pm = self._scaled_icon("sun", self._sun_source, icon_size)
        moon_pm = self._scaled_icon("moon", self._moon_source, icon_size)
        tint = QtGui.QColor(255, 255, 255) if dark else QtGui.QColor(0, 0, 0)
        sun_pm = self._tint_pixmap(sun_pm, tint)
        moon_pm = self._tint_pixmap(moon_pm, tint)

        highlight_size = icon_size + 9
        highlight_center_x = right_x + icon_size * 0.5 - (right_x - left_x) * self._handle_progress
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(QtGui.QColor("#F7921E"))
        p.drawEllipse(
            QtCore.QRectF(
                highlight_center_x - highlight_size * 0.5,
                center_y - highlight_size * 0.5,
                highlight_size,
                highlight_size,
            )
        )

        if not sun_pm.isNull():
            p.setOpacity(1.0 - 0.62 * self._handle_progress)
            p.drawPixmap(int(right_x), int(center_y - sun_pm.height() / 2), sun_pm)
        if not moon_pm.isNull():
            p.setOpacity(0.38 + 0.62 * self._handle_progress)
            p.drawPixmap(int(left_x), int(center_y - moon_pm.height() / 2), moon_pm)
        p.setOpacity(1.0)


class ArrowButton(QtWidgets.QPushButton):
    def __init__(self, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.setText("")
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedSize(36, 36)
        self.setIconSize(QtCore.QSize(18, 18))
        self.setFocusPolicy(QtCore.Qt.NoFocus)
        self._dark = False
        self._accent = False
        self._source_pixmap = QtGui.QPixmap(str(ASSET_DIR / "right-arrow.png"))

    def set_theme(self, dark: bool, accent: bool = False) -> None:
        self._dark = dark
        self._accent = accent
        if self._source_pixmap.isNull():
            self.setIcon(QtGui.QIcon())
            return

        if accent:
            color = QtGui.QColor("#F7921E")
        elif dark:
            color = QtGui.QColor("#E8EEF7")
        else:
            color = QtGui.QColor("#303641")

        source = self._source_pixmap.scaled(
            18,
            18,
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation,
        )
        tinted = QtGui.QPixmap(source.size())
        tinted.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(tinted)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.drawPixmap(0, 0, source)
        painter.setCompositionMode(QtGui.QPainter.CompositionMode_SourceIn)
        painter.fillRect(tinted.rect(), color)
        painter.end()
        self.setIcon(QtGui.QIcon(tinted))


class SodCard(QtWidgets.QWidget):
    clicked = QtCore.Signal(str)

    NORMAL_RECT = QtCore.QRect(6, 8, 168, 280)
    HOVER_RECT = QtCore.QRect(2, 1, 176, 288)

    def __init__(self, spec: AppSpec, dark: bool = False, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self.spec = spec
        self._dark = dark
        self._hovered = False
        self._pressed = False
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedSize(180, 292)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, False)

        # The outer widget keeps a fixed slot in the row. Only this inner frame grows,
        # so neighboring cards do not jump when the hover animation runs.
        self.card_frame = QtWidgets.QFrame(self)
        self.card_frame.setObjectName("sodCardRoot")
        self.card_frame.setGeometry(self.NORMAL_RECT)
        self.card_frame.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)

        self.shadow = QtWidgets.QGraphicsDropShadowEffect(self.card_frame)
        self.shadow.setBlurRadius(24)
        self.shadow.setOffset(0, 8)
        self.card_frame.setGraphicsEffect(self.shadow)

        self._hover_anim = QtCore.QPropertyAnimation(self.card_frame, b"geometry", self)
        self._hover_anim.setDuration(155)
        self._hover_anim.setEasingCurve(QtCore.QEasingCurve.OutCubic)

        layout = QtWidgets.QVBoxLayout(self.card_frame)
        layout.setContentsMargins(16, 18, 16, 18)
        layout.setSpacing(0)

        layout.addSpacing(8)

        self.logo_box = QtWidgets.QWidget()
        self.logo_box.setObjectName("logoBox")
        self.logo_box.setFixedSize(126, 94)
        logo_layout = QtWidgets.QVBoxLayout(self.logo_box)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo_layout.setSpacing(0)

        self.image_label = QtWidgets.QLabel()
        self.image_label.setAlignment(QtCore.Qt.AlignCenter)
        self.image_label.setFixedSize(126, 94)
        logo_layout.addWidget(self.image_label, 0, QtCore.Qt.AlignCenter)
        layout.addWidget(self.logo_box, 0, QtCore.Qt.AlignCenter)

        layout.addSpacing(14)

        self.title_label = QtWidgets.QLabel(spec.card_title)
        self.title_label.setObjectName("titleLabel")
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        self.title_label.setWordWrap(True)
        self.title_label.setFixedHeight(42)
        layout.addWidget(self.title_label)

        self.subtitle_label = QtWidgets.QLabel(spec.card_subtitle)
        self.subtitle_label.setObjectName("subtitleLabel")
        self.subtitle_label.setAlignment(QtCore.Qt.AlignCenter)
        self.subtitle_label.setWordWrap(True)
        self.subtitle_label.setFixedHeight(42)
        layout.addWidget(self.subtitle_label)

        layout.addStretch(1)

        self.arrow_button = ArrowButton()
        self.arrow_button.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, True)
        self.arrow_button.set_theme(dark, accent=False)
        layout.addWidget(self.arrow_button, 0, QtCore.Qt.AlignCenter)

        self.apply_theme(dark)
        self.reload_image(dark)

    def _animate_hover(self, hovering: bool) -> None:
        self._hover_anim.stop()
        self._hover_anim.setStartValue(self.card_frame.geometry())
        self._hover_anim.setEndValue(self.HOVER_RECT if hovering else self.NORMAL_RECT)
        self._hover_anim.setDuration(155 if hovering else 125)
        self._hover_anim.setEasingCurve(
            QtCore.QEasingCurve.OutBack if hovering else QtCore.QEasingCurve.InOutCubic
        )
        self._hover_anim.start()

    def apply_theme(self, dark: bool) -> None:
        self._dark = dark
        c = theme(dark)
        bg = c["card_pressed"] if self._pressed else c["card_hover"] if self._hovered else c["card"]
        border = c["accent"] if self._hovered else c["card_border"]
        title_size = 13 if self.spec.key != "projectpoint" else 12
        subtitle_size = 9 if self.spec.key == "projectpoint" else 1

        # SIGNAL's source badge already contains the blue rounded rectangle and
        # proper white glyphs/text. Keep the holder itself transparent so there
        # is no oversized blue slab around the logo.
        self.card_frame.setStyleSheet(
            f"""
            QFrame#sodCardRoot {{
                background: {bg};
                border: 1px solid {border};
                border-radius: 22px;
            }}
            QWidget#logoBox {{
                background: transparent;
                border: none;
            }}
            QLabel {{
                background: transparent;
                border: none;
            }}
            QLabel#titleLabel {{
                color: {c['card_title']};
                font-size: {title_size}px;
                font-weight: 600;
            }}
            QLabel#subtitleLabel {{
                color: {c['card_subtitle']};
                font-size: {subtitle_size}px;
                font-weight: 500;
            }}
            QPushButton {{
                background: {c['arrow_bg']};
                color: {c['arrow_fg']};
                border: 1px solid {c['arrow_border']};
                border-radius: 18px;
                font-size: 18px;
                font-weight: 700;
                padding: 0px;
            }}
            """
        )
        self.arrow_button.set_theme(dark, accent=False)
        self.shadow.setColor(QtGui.QColor(8, 10, 18, 76 if dark else 34))
        self.shadow.setBlurRadius(32 if self._hovered else 24)
        self.shadow.setOffset(0, 11 if self._hovered else 8)

    def reload_image(self, dark: Optional[bool] = None) -> None:
        if dark is not None:
            self._dark = dark
        image_name = self.spec.image_dark_name if self._dark else self.spec.image_light_name
        image_path = ASSET_DIR / image_name
        if image_path.exists():
            pixmap = QtGui.QPixmap(str(image_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    self.spec.logo_max_width,
                    self.spec.logo_max_height,
                    QtCore.Qt.KeepAspectRatio,
                    QtCore.Qt.SmoothTransformation,
                )
                self.image_label.setPixmap(scaled)
                self.image_label.setText("")
                return
        self.image_label.setPixmap(QtGui.QPixmap())
        self.image_label.setText(self.spec.title)

    def enterEvent(self, event: QtCore.QEvent) -> None:
        self._hovered = True
        self.raise_()
        self.apply_theme(self._dark)
        self._animate_hover(True)
        super().enterEvent(event)

    def leaveEvent(self, event: QtCore.QEvent) -> None:
        self._hovered = False
        self._pressed = False
        self.apply_theme(self._dark)
        self._animate_hover(False)
        super().leaveEvent(event)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.LeftButton:
            self._pressed = True
            self.apply_theme(self._dark)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        was_pressed = self._pressed
        self._pressed = False
        self.apply_theme(self._dark)
        if was_pressed and event.button() == QtCore.Qt.LeftButton and self.rect().contains(event.position().toPoint()):
            self.clicked.emit(self.spec.key)
            event.accept()
            return
        super().mouseReleaseEvent(event)


class LauncherWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1200, 700)
        self.setMinimumSize(1020, 620)
        self._processes: Dict[str, subprocess.Popen] = {}
        self._cards: Dict[str, SodCard] = {}
        self._active_process_key: Optional[str] = None
        self._process_watch_timer = QtCore.QTimer(self)
        self._process_watch_timer.setInterval(300)
        self._process_watch_timer.timeout.connect(self._check_active_process)

        icon_path = ASSET_DIR / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QtGui.QIcon(str(icon_path)))

        settings = QtCore.QSettings("Larix", "SODManager")
        self.is_dark = settings.value("dark_theme", False, type=bool)
        self._settings = settings

        self._setup_ui()
        self._apply_theme()
        install_window_state(self, "launcher")

    def _setup_ui(self) -> None:
        central = BackgroundWidget(self.is_dark)
        self._background = central
        self.setCentralWidget(central)

        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(46, 24, 46, 26)
        root.setSpacing(0)

        top = QtWidgets.QHBoxLayout()
        top.addStretch(1)
        self.theme_switch = ThemeSwitch()
        self.theme_switch.setChecked(self.is_dark, animate=False, emit_signal=False)
        self.theme_switch.toggled.connect(self._toggle_theme)
        top.addWidget(self.theme_switch)
        root.addLayout(top)

        root.addSpacing(26)

        brand_row = QtWidgets.QHBoxLayout()
        brand_row.addStretch(1)
        self.brand_logo = QtWidgets.QLabel()
        self.brand_logo.setAlignment(QtCore.Qt.AlignCenter)
        self.brand_logo.setFixedSize(600, 92)
        brand_row.addWidget(self.brand_logo)
        brand_row.addStretch(1)
        root.addLayout(brand_row)

        root.addSpacing(22)

        self.section_title = QtWidgets.QLabel("Выберите СОД")
        self.section_title.setAlignment(QtCore.Qt.AlignCenter)
        root.addWidget(self.section_title)

        root.addStretch(1)
        root.addSpacing(12)

        cards_row = QtWidgets.QHBoxLayout()
        cards_row.setSpacing(14)
        cards_row.addStretch(1)
        for spec in APPS:
            card = SodCard(spec, dark=self.is_dark)
            card.clicked.connect(self._launch_by_key)
            cards_row.addWidget(card)
            self._cards[spec.key] = card
        cards_row.addStretch(1)
        root.addLayout(cards_row)

        root.addStretch(2)

        bottom = QtWidgets.QHBoxLayout()
        bottom.addStretch(1)
        self.hint_label = QtWidgets.QLabel("ⓘ  Нажмите на карточку, чтобы открыть инструмент выбранной СОД")
        bottom.addWidget(self.hint_label)
        bottom.addStretch(1)
        root.addLayout(bottom)

        root.addSpacing(18)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        self.version_label = QtWidgets.QLabel(f"Версия {APP_VERSION}")
        footer.addWidget(self.version_label)
        root.addLayout(footer)

    def _reload_brand_logo(self) -> None:
        image_name = "brand_larix_cde_dark.png" if self.is_dark else "brand_larix_cde_light.png"
        image_path = ASSET_DIR / image_name
        if image_path.exists():
            pixmap = QtGui.QPixmap(str(image_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(520, 64, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                self.brand_logo.setPixmap(scaled)
                self.brand_logo.setText("")
                return
        self.brand_logo.setPixmap(QtGui.QPixmap())
        self.brand_logo.setText("Larix CDE")

    def _apply_theme(self) -> None:
        c = theme(self.is_dark)
        self._background.set_dark(self.is_dark)
        self.theme_switch.setChecked(self.is_dark, animate=False, emit_signal=False)
        self.theme_switch.setToolTip("Светлая тема" if self.is_dark else "Тёмная тема")

        self.section_title.setStyleSheet(
            f"font-size: 17px; font-weight: 500; color: {c['muted']}; background: transparent;"
        )
        self.hint_label.setStyleSheet(
            f"font-size: 11px; font-weight: 500; color: {c['hint']}; background: transparent;"
        )
        self.version_label.setStyleSheet(
            f"font-size: 11px; color: {c['hint']}; background: transparent;"
        )

        self._reload_brand_logo()
        for card in self._cards.values():
            card.apply_theme(self.is_dark)
            card.reload_image(self.is_dark)

        self._apply_windows_titlebar()

    def _apply_windows_titlebar(self) -> None:
        """Use a black native Windows title bar in dark mode."""
        if os.name != "nt":
            return
        try:
            hwnd = int(self.winId())
            dwm = ctypes.windll.dwmapi

            # Windows 10/11: tell DWM that this window uses a dark caption.
            dark_value = ctypes.c_int(1 if self.is_dark else 0)
            for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE, older fallback
                try:
                    result = dwm.DwmSetWindowAttribute(
                        ctypes.c_void_p(hwnd),
                        ctypes.c_uint(attr),
                        ctypes.byref(dark_value),
                        ctypes.sizeof(dark_value),
                    )
                    if result == 0:
                        break
                except Exception:
                    pass

            # Windows 11: explicitly make the caption black and the caption text white.
            # 0xFFFFFFFF restores the system default for the light theme.
            caption_color = ctypes.c_uint(0x00000000 if self.is_dark else 0xFFFFFFFF)
            text_color = ctypes.c_uint(0x00FFFFFF if self.is_dark else 0xFFFFFFFF)
            try:
                dwm.DwmSetWindowAttribute(
                    ctypes.c_void_p(hwnd), ctypes.c_uint(35),
                    ctypes.byref(caption_color), ctypes.sizeof(caption_color)
                )
                dwm.DwmSetWindowAttribute(
                    ctypes.c_void_p(hwnd), ctypes.c_uint(36),
                    ctypes.byref(text_color), ctypes.sizeof(text_color)
                )
            except Exception:
                pass
        except Exception:
            pass

    def _toggle_theme(self, checked: bool) -> None:
        self.is_dark = checked
        self._settings.setValue("dark_theme", self.is_dark)
        self._apply_theme()

    def _launch_by_key(self, key: str) -> None:
        spec = next((item for item in APPS if item.key == key), None)
        if spec is None:
            return

        running = self._processes.get(key)
        if running is not None and running.poll() is None:
            # The launcher should normally be hidden while a child is active,
            # but keep this guard for rapid double clicks.
            return

        try:
            command, cwd = self._command_for(spec)
            creationflags = 0
            if os.name == "nt" and not getattr(sys, "frozen", False):
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

            child_env = os.environ.copy()
            child_env["SOD_MANAGER_ASSETS"] = str(ASSET_DIR)
            child_env["SOD_MANAGER_THEME"] = "dark" if self.is_dark else "light"
            child_env["SOD_MANAGER_PARENT"] = "1"
            process = subprocess.Popen(command, cwd=str(cwd), creationflags=creationflags, env=child_env)
            self._processes[key] = process
            self._active_process_key = key

            # Visually replace the launcher with the selected CDE tool instead
            # of leaving two application windows visible at the same time.
            self.hide()
            self._process_watch_timer.start()
        except FileNotFoundError as exc:
            QtWidgets.QMessageBox.critical(self, "Файл не найден", f"Не удалось открыть {spec.title}.\n\n{exc}")
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Ошибка запуска", f"Не удалось открыть {spec.title}.\n\n{type(exc).__name__}: {exc}")

    def _check_active_process(self) -> None:
        key = self._active_process_key
        if not key:
            self._process_watch_timer.stop()
            return

        process = self._processes.get(key)
        if process is not None and process.poll() is None:
            return

        self._process_watch_timer.stop()
        self._active_process_key = None
        self._processes.pop(key, None)

        # A child SOD can change the shared theme. Read it back before the
        # launcher becomes visible again so the whole application stays in sync.
        shared_dark = self._settings.value("dark_theme", self.is_dark, type=bool)
        if bool(shared_dark) != self.is_dark:
            self.is_dark = bool(shared_dark)
            self._apply_theme()

        # Return to CDE selection when the launched tool is closed.
        self.show()
        self.raise_()
        self.activateWindow()
        QtCore.QTimer.singleShot(0, self._apply_windows_titlebar)

    def _command_for(self, spec: AppSpec) -> tuple[list[str], Path]:
        if getattr(sys, "frozen", False):
            entry = ROOT / spec.frozen_entry
            if not entry.exists():
                raise FileNotFoundError(f"Ожидался файл:\n{entry}")
            return [str(entry)], entry.parent

        entry = ROOT / spec.source_entry
        if not entry.exists():
            raise FileNotFoundError(f"Ожидался файл:\n{entry}")
        return [sys.executable, str(entry)], entry.parent


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("Larix")
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Segoe UI", 10))

    window = LauncherWindow()
    window.show()
    QtCore.QTimer.singleShot(0, window._apply_windows_titlebar)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
