# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from core.paths import application_root
from core.plugin import PluginSpec
from core.plugin_loader import PluginLoadError, discover_plugins
from core.plugin_dispatcher import (
    PluginDispatchError,
    clean_plugin_argv,
    dispatch_plugin,
    parse_plugin_key,
)
from shared.theme_core import (
    ThemeToggle, apply_windows_titlebar_theme, classify_plugin_exit,
    install_window_state, persist_dark_theme
)

APP_TITLE = "Larix CDE"
APP_VERSION = "1.0.0"


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

    def __init__(self, spec: PluginSpec, dark: bool = False, parent: Optional[QtWidgets.QWidget] = None):
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
        try:
            self.plugins = discover_plugins(ROOT)
        except PluginLoadError as exc:
            self.plugins = tuple()
            QtCore.QTimer.singleShot(0, lambda message=str(exc): QtWidgets.QMessageBox.critical(self, "Ошибка плагинов", message))

        self._processes: Dict[str, subprocess.Popen] = {}
        self._process_logs: Dict[str, tuple[object, Path]] = {}
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
        self.theme_switch = ThemeToggle(ASSET_DIR)
        self.theme_switch.setChecked(self.is_dark, animate=False)
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
        for spec in self.plugins:
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

        self.recommendation_label = QtWidgets.QLabel(
            "Перед каждым заполнением рекомендуем скачивать актуальные шаблоны."
        )
        self.recommendation_label.setAlignment(QtCore.Qt.AlignCenter)
        root.addWidget(self.recommendation_label)

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
        self.theme_switch.setChecked(self.is_dark, animate=False)
        self.theme_switch.setToolTip("Светлая тема" if self.is_dark else "Тёмная тема")

        self.section_title.setStyleSheet(
            f"font-size: 17px; font-weight: 500; color: {c['muted']}; background: transparent;"
        )
        self.hint_label.setStyleSheet(
            f"font-size: 11px; font-weight: 500; color: {c['hint']}; background: transparent;"
        )
        self.recommendation_label.setStyleSheet(
            f"font-size: 11px; font-weight: 500; color: {c['accent']}; background: transparent;"
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
        apply_windows_titlebar_theme(self, self.is_dark)

    def _toggle_theme(self, checked: bool) -> None:
        self.is_dark = bool(checked)
        persist_dark_theme(self.is_dark)
        self._apply_theme()

    def _launch_by_key(self, key: str) -> None:
        spec = next((item for item in self.plugins if item.key == key), None)
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
            child_env["SOD_MANAGER_ROOT"] = str(ROOT)

            # Each plugin is still started as its own Python script.  Python puts
            # the plugin folder (for example plugins/signal) into sys.path, but
            # the refactored shared/ package lives in the project root.  Add the
            # project root explicitly so imports such as ``from shared...`` keep
            # working after apps/ was moved to plugins/.
            existing_pythonpath = child_env.get("PYTHONPATH", "").strip()
            child_env["PYTHONPATH"] = (
                str(ROOT)
                if not existing_pythonpath
                else str(ROOT) + os.pathsep + existing_pythonpath
            )

            # Keep child startup errors visible. Previously CREATE_NO_WINDOW hid
            # import/runtime failures and a broken plugin looked like a dead card.
            log_dir = ROOT / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            log_path = log_dir / f"launch_{key}.log"
            log_handle = open(log_path, "w", encoding="utf-8", errors="replace")
            log_handle.write(f"Command: {command!r}\n")
            log_handle.write(f"Working directory: {cwd}\n")
            log_handle.write(f"PYTHONPATH: {child_env.get('PYTHONPATH', '')}\n\n")
            log_handle.flush()

            process = subprocess.Popen(
                command,
                cwd=str(cwd),
                creationflags=creationflags,
                env=child_env,
                stdout=log_handle,
                stderr=subprocess.STDOUT,
            )
            self._processes[key] = process
            self._process_logs[key] = (log_handle, log_path)
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
        process = self._processes.pop(key, None)
        return_code = process.returncode if process is not None else None

        log_info = self._process_logs.pop(key, None)
        log_path = None
        if log_info is not None:
            log_handle, log_path = log_info
            try:
                log_handle.flush()
                log_handle.close()
            except Exception:
                pass

        # A child SOD can change the shared theme. Read it back before the
        # launcher becomes visible again so the whole application stays in sync.
        shared_dark = self._settings.value("dark_theme", self.is_dark, type=bool)
        if bool(shared_dark) != self.is_dark:
            self.is_dark = bool(shared_dark)
            self._apply_theme()

        # If a plugin crashed during startup, show the actual traceback instead
        # of silently returning to the launcher. Keep the full log on disk.
        exit_kind = classify_plugin_exit(return_code)
        if exit_kind == "return_to_menu":
            self.show()
            self.raise_()
            self.activateWindow()
            QtCore.QTimer.singleShot(0, self._apply_windows_titlebar)
            return

        if exit_kind == "crashed" and log_path is not None:
            try:
                raw = log_path.read_text(encoding="utf-8", errors="replace").strip()
            except Exception:
                raw = ""
            lines = raw.splitlines()
            tail = "\n".join(lines[-18:]) if lines else "Ошибка запуска без текста."
            spec = next((item for item in self.plugins if item.key == key), None)
            title = spec.title if spec is not None else key
            QtWidgets.QMessageBox.critical(
                self,
                "Ошибка запуска",
                f"{title} завершился с ошибкой (код {return_code}).\n\n"
                f"{tail}\n\nПолный лог: {log_path}",
            )

        # A system close (or a failed child) ends the complete CDE session;
        # the hidden menu must not reappear.
        self.close()
        QtWidgets.QApplication.quit()

    def _command_for(self, spec: PluginSpec) -> tuple[list[str], Path]:
        # In a one-file build every CDE tool is an internal mode of this same
        # executable.  Source runs retain the convenient ``python launcher.py``
        # form, so no separate plugin executable is needed in either mode.
        if getattr(sys, "frozen", False):
            return [sys.executable, "--plugin", spec.key], ROOT

        entry = spec.source_path
        if not entry.exists():
            raise FileNotFoundError(
                f"Не найден entry point плагина {spec.title}:\n{entry}"
            )
        launcher_path = ROOT / "launcher.py"
        return [sys.executable, str(launcher_path), "--plugin", spec.key], ROOT


def main() -> int:
    try:
        plugin_key = parse_plugin_key(sys.argv[1:])
    except PluginDispatchError as exc:
        print(f"Ошибка запуска: {exc}", file=sys.stderr)
        return 2
    if plugin_key is not None:
        # Qt parses sys.argv itself; hide the internal launcher protocol before
        # creating the selected tool's QApplication.
        sys.argv[:] = clean_plugin_argv(sys.argv)
        try:
            return dispatch_plugin(plugin_key)
        except PluginDispatchError as exc:
            print(f"Ошибка запуска: {exc}", file=sys.stderr)
            return 2

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("Larix")
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Segoe UI", 10))

    window = LauncherWindow()
    window.show()
    QtCore.QTimer.singleShot(0, window._apply_windows_titlebar)
    if os.environ.get("CDE_SMOKE_TEST", "").strip() == "1":
        QtCore.QTimer.singleShot(250, app.quit)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
