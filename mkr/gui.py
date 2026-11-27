"""PySide6 sidecar window — floating, always-on-top, borderless with custom title bar.

Three screens stacked in a QStackedWidget: Home → Wizard → Preview.
Language button toggles zh/en by mutating Config.lang and calling refresh_texts().
"""
from __future__ import annotations

import colorsys
import sys
from typing import Any

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QColor, QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from mkr import __version__
from mkr.config import DEFAULT_THEME, Config
from mkr.excel_bridge import ExcelComBridge
from mkr.formulas import base as fbase  # noqa: F401 — triggers registry
from mkr.formulas.base import FormulaBase, all_formula_names, get_formula
from mkr.i18n import tr
from mkr.intent import detect

# ---------- theme palette ------------------------------------------------
# One base colour drives the whole theme: every surface / text colour is
# derived from it so contrast stays readable for any user-picked colour.
_BG = "#5C2D91"
_BG_CARD = "#4E2578"
_BG_DEEP = "#3A1A5C"
_BORDER = "#C9B6E8"
_ACCENT = "#431F66"
_ACCENT_HOVER = "#744DA9"
_TEXT = "#F6F1FD"
_TEXT_DIM = "#DCCEF2"
_WARN = "#FFD27A"
_OK = "#E9DFFB"

DEFAULT_BASE = DEFAULT_THEME


def _left_button_down() -> bool:
    """True while the physical left mouse button is held anywhere on screen.

    Pick mode must never confirm mid-drag: a column-header drag from B to D
    pauses on B:B long enough for a poll to fire, and popping the window back
    to the front during a drag also aborts the drag itself.
    """
    try:
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000)
    except Exception:
        return False


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    v = value.lstrip("#")
    if len(v) == 3:
        v = "".join(ch * 2 for ch in v)
    return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)


def _rgb_to_hex(rgb) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, round(c))) for c in rgb))


def _mix(color_a: str, color_b: str, t: float) -> str:
    """Linear blend: t=0 → a, t=1 → b."""
    a, b = _hex_to_rgb(color_a), _hex_to_rgb(color_b)
    return _rgb_to_hex(tuple(x + (y - x) * t for x, y in zip(a, b)))


def _luminance(value: str) -> float:
    def _ch(raw: int) -> float:
        c = raw / 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (_ch(c) for c in _hex_to_rgb(value))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _hls(value: str) -> tuple[float, float, float]:
    r, g, b = _hex_to_rgb(value)
    return colorsys.rgb_to_hls(r / 255, g / 255, b / 255)


def _from_hls(h: float, light: float, sat: float) -> str:
    return _rgb_to_hex(tuple(c * 255 for c in colorsys.hls_to_rgb(h, light, sat)))


def _shift(value: str, lightness: float, saturation: float | None = None) -> str:
    """Keep the hue, set a new HLS lightness (and optionally saturation)."""
    h, _, sat = _hls(value)
    return _from_hls(h, lightness, sat if saturation is None else saturation)


def normalize_hex(value: str | None) -> str | None:
    if not value:
        return None
    v = value.strip().lstrip("#")
    if len(v) == 3:
        v = "".join(ch * 2 for ch in v)
    if len(v) != 6:
        return None
    try:
        int(v, 16)
    except ValueError:
        return None
    return "#" + v.upper()


def windows_accent_hex() -> str | None:
    """Windows 10/11 accent colour.  AccentColorMenu stores ABGR."""
    try:
        import winreg

        key = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            val = int(winreg.QueryValueEx(k, "AccentColorMenu")[0])
        return _rgb_to_hex((val & 0xFF, (val >> 8) & 0xFF, (val >> 16) & 0xFF))
    except Exception:
        return None


def derive_palette(base: str) -> dict[str, str]:
    """Derive every theme colour from one base colour.

    Text colour follows WCAG AA: use light text whenever white reaches 4.5:1
    on the base, otherwise fall back to dark ink on light surfaces.
    """
    h, light, sat = _hls(base)
    lum = _luminance(base)
    use_light_text = (1.05 / (lum + 0.05)) >= 4.5
    if use_light_text:  # dark base → light text
        return {
            "card": _from_hls(h, max(light - 0.07, 0.05), sat),
            "deep": _from_hls(h, max(light * 0.62, 0.04), sat),
            "accent": _from_hls(h, max(light * 0.72, 0.05), min(sat + 0.05, 1)),
            "accent_hover": _from_hls(h, min(max(light * 1.2, light + 0.06), 0.88), sat),
            "border": _mix(base, "#FFFFFF", 0.62),
            "text": _mix(base, "#FFFFFF", 0.92),
            "text_dim": _mix(base, "#FFFFFF", 0.76),
            "warn": "#FFD27A",
        }
    # light base → dark ink text
    return {
        "card": _from_hls(h, 0.94, max(sat * 0.4, 0.05)),
        "deep": _from_hls(h, 0.97, max(sat * 0.3, 0.04)),
        "accent": _from_hls(h, min(light * 0.45, 0.30), min(sat * 0.9, 1)),
        "accent_hover": _from_hls(h, min(light * 0.6, 0.40), min(sat * 0.95, 1)),
        "border": _from_hls(h, 0.32, min(sat * 0.8, 1)),
        "text": _mix(base, "#1A1200", 0.90),
        "text_dim": _mix(base, "#1A1200", 0.74),
        "warn": "#7A1300",
    }


class Style:
    """QSS strings; values are (re)generated by _rebuild_style()."""
    WINDOW = ""
    CARD = ""
    BTN = ""
    BTN_GHOST = ""
    BTN_TITLE = ""
    INPUT = ""
    LABEL = ""
    LABEL_TITLE = ""
    LABEL_STEP = ""
    LABEL_TIP = ""
    LABEL_WARN = ""
    LIST = ""
    PREVIEW_BOX = ""
    STATUS = ""


def _rebuild_style() -> None:
    Style.WINDOW = f"""
        QWidget {{ background: {_BG}; color: {_TEXT}; font-family: "Microsoft YaHei", "Segoe UI"; font-size: 13px; }}
    """
    Style.CARD = f"""
        QFrame {{ background: {_BG_CARD}; border-radius: 10px; padding: 12px; }}
        QFrame:hover {{ background: {_ACCENT}; }}
    """
    Style.BTN = f"""
        QPushButton {{ background: {_ACCENT}; color: white; border: none; border-radius: 6px;
                       padding: 6px 14px; font-weight: bold; }}
        QPushButton:hover {{ background: {_ACCENT_HOVER}; }}
        QPushButton:disabled {{ background: {_BORDER}; color: {_TEXT_DIM}; }}
    """
    Style.BTN_GHOST = f"""
        QPushButton {{ background: transparent; color: {_TEXT}; border: 1px solid {_BORDER};
                       border-radius: 6px; padding: 5px 12px; }}
        QPushButton:hover {{ background: {_ACCENT}; color: white; }}
    """
    Style.BTN_TITLE = f"""
        QPushButton {{ background: transparent; color: {_TEXT_DIM}; border: none; font-size: 16px;
                       padding: 2px 8px; }}
        QPushButton:hover {{ color: {_TEXT}; background: {_ACCENT_HOVER}; border-radius: 4px; }}
    """
    Style.INPUT = f"""
        QLineEdit, QComboBox, QPlainTextEdit {{
            background: {_BG_DEEP}; border: 1px solid {_BORDER}; border-radius: 6px;
            color: {_TEXT}; padding: 6px; selection-background-color: {_ACCENT_HOVER};
        }}
        QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border: 1px solid {_ACCENT_HOVER}; }}
    """
    Style.LABEL = f"color: {_TEXT_DIM}; font-size: 12px;"
    Style.LABEL_TITLE = f"color: {_TEXT}; font-size: 16px; font-weight: bold;"
    Style.LABEL_STEP = f"color: {_TEXT}; font-size: 14px; font-weight: bold;"
    Style.LABEL_TIP = f"color: {_TEXT_DIM}; font-size: 11px; padding: 2px 4px;"
    Style.LABEL_WARN = f"color: {_WARN}; font-size: 11px; padding: 2px 4px;"
    Style.LIST = f"""
        QListWidget {{ background: {_BG}; border: none; outline: none; }}
        QListWidget::item {{ background: {_BG_CARD}; border-radius: 8px; padding: 10px; margin: 3px 0; }}
        QListWidget::item:hover {{ background: {_ACCENT}; color: white; }}
        QListWidget::item:selected {{ background: {_ACCENT_HOVER}; color: white; }}
    """
    Style.PREVIEW_BOX = f"""
        QFrame {{ background: {_BG_DEEP}; border: 1px solid {_BORDER}; border-radius: 8px;
                  padding: 10px; font-family: "Consolas", "Courier New"; font-size: 13px; color: {_TEXT}; }}
    """
    Style.STATUS = f"color: {_TEXT_DIM}; font-size: 11px;"


def set_base_color(base: str | None = None) -> str:
    """Rebuild the whole palette from `base` (None → Windows accent colour)."""
    global _BG, _BG_CARD, _BG_DEEP, _BORDER, _ACCENT, _ACCENT_HOVER, _TEXT, _TEXT_DIM, _WARN, _OK
    chosen = normalize_hex(base) or windows_accent_hex() or DEFAULT_BASE
    pal = derive_palette(chosen)
    _BG = chosen
    _BG_CARD = pal["card"]
    _BG_DEEP = pal["deep"]
    _BORDER = pal["border"]
    _ACCENT = pal["accent"]
    _ACCENT_HOVER = pal["accent_hover"]
    _TEXT = pal["text"]
    _TEXT_DIM = pal["text_dim"]
    _WARN = pal["warn"]
    _OK = pal["text_dim"]
    _rebuild_style()
    return chosen


_rebuild_style()


# ==========================================================================
# Home screen — search box + recommended formula cards
# ==========================================================================
class HomeScreen(QWidget):
    open_wizard = Signal(str)  # formula name

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 8, 14, 10)
        v.setSpacing(8)

        self.prompt_label = QLabel()
        self.prompt_label.setStyleSheet(Style.LABEL_TITLE)
        v.addWidget(self.prompt_label)

        self.search = QLineEdit()
        self.search.setStyleSheet(Style.INPUT)
        self.search.returnPressed.connect(self._on_search)
        v.addWidget(self.search)

        self.recom_head = QLabel()
        self.recom_head.setStyleSheet(Style.LABEL)
        v.addWidget(self.recom_head)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(Style.LIST)
        self.list_widget.itemDoubleClicked.connect(self._on_card_clicked)
        v.addWidget(self.list_widget, stretch=1)

    # ---------- runtime text refresh ----------
    def refresh_texts(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        self.prompt_label.setText(tr("home_prompt", lang=lang))
        self.search.setPlaceholderText(tr("home_placeholder", lang=lang))
        self._render_cards("")  # re-render all cards with current lang

    # ---------- search & card rendering ----------
    def _on_search(self) -> None:
        query = self.search.text().strip()
        self._render_cards(query)

    def _render_cards(self, query: str) -> None:
        cfg = Config.get()
        lang = cfg.lang
        self.list_widget.clear()

        if query:
            hits = detect(query)
            if not hits:
                item = QListWidgetItem(tr("home_no_match", lang=lang))
                item.setFlags(Qt.ItemFlag.NoItemFlags)
                self.list_widget.addItem(item)
                self.recom_head.setText(tr("home_recommend", lang=lang))
                return
            names = [h.formula for h in hits]
        else:
            names = all_formula_names()
            self.recom_head.setText(tr("home_all", lang=lang))

        for name in names:
            f = get_formula(name)
            if not f:
                continue
            card = _build_card(f, lang)
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, name)
            item.setSizeHint(card.sizeHint())
            self.list_widget.addItem(item)
            self.list_widget.setItemWidget(item, card)

    def _on_card_clicked(self, item: QListWidgetItem) -> None:
        name = item.data(Qt.ItemDataRole.UserRole)
        if name:
            self.open_wizard.emit(name)

    def apply_style(self) -> None:
        """Re-apply every widget-level stylesheet after a theme change."""
        self.prompt_label.setStyleSheet(Style.LABEL_TITLE)
        self.search.setStyleSheet(Style.INPUT)
        self.recom_head.setStyleSheet(Style.LABEL)
        self.list_widget.setStyleSheet(Style.LIST)


def _build_card(f: FormulaBase, lang: str) -> QWidget:
    card = QWidget()
    card.setStyleSheet(
        f"QWidget {{ background: {_BG_CARD}; border-radius: 8px; }} "
        f"QWidget:hover {{ background: {_ACCENT}; }}"
    )
    v = QVBoxLayout(card)
    v.setContentsMargins(10, 8, 10, 8)
    v.setSpacing(2)

    title = QLabel(f.zh_name if lang == "zh" else f.en_name)
    title.setStyleSheet(f"color: {_TEXT}; font-weight: bold; font-size: 13px;")
    v.addWidget(title)

    desc = QLabel(f.zh_desc if lang == "zh" else f.en_desc)
    desc.setStyleSheet(f"color: {_TEXT_DIM}; font-size: 11px;")
    desc.setWordWrap(True)
    v.addWidget(desc)

    example = QLabel((f.zh_example if lang == "zh" else f.en_example) or f.example)  # type: ignore[attr-defined]
    example = QLabel(getattr(f, "zh_example", "") or getattr(f, "en_example", "") or f.zh_example)
    example.setText(getattr(f, "zh_example", "") or getattr(f, "en_example", ""))
    example.setStyleSheet(f"color: {_TEXT}; font-family: Consolas; font-size: 11px; padding-top: 2px;")
    v.addWidget(example)
    return card


# ==========================================================================
# Wizard screen — step by step parameter input
# ==========================================================================
class WizardScreen(QWidget):
    prev = Signal()
    next_ready = Signal()
    done = Signal()
    cancelled = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.formula: FormulaBase | None = None
        self.steps: list[Any] = []  # ArgStep
        self.values: dict[str, str] = {}
        self.current = 0
        self._next_slot: Any = None  # slot currently wired to btn_next
        # Live, non-blocking cell picking (see _pick_from_excel):
        #   _pick_awaiting=True → the poll timer below watches Excel's selection
        self._pick_awaiting = False
        self._pick_mode = "cell"
        self._pick_base: dict | None = None   # {'sheet','address','ref'} at entry
        self._pick_last: str | None = None    # previous poll result (debounce)
        self._pick_ticks = 0
        self._pick_pressed = False            # left button seen down since enter
        self._pick_click_done = False         # a full click (down→up) completed
        self._pick_minimized = False  # did we minimise the window for pick mode?
        self._pick_timer = QTimer(self)
        self._pick_timer.setInterval(200)  # 200 ms → a click is captured in ~400 ms
        self._pick_timer.timeout.connect(self._poll_pick)
        self._build_ui()

    def _build_ui(self) -> None:
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 8, 14, 10)
        v.setSpacing(6)

        self.step_title = QLabel()
        self.step_title.setStyleSheet(Style.LABEL_STEP)
        v.addWidget(self.step_title)

        self.zh_label = QLabel()
        self.zh_label.setStyleSheet(Style.LABEL)
        self.zh_label.setWordWrap(True)
        v.addWidget(self.zh_label)

        self.en_label = QLabel()
        self.en_label.setStyleSheet(Style.LABEL)
        self.en_label.setWordWrap(True)
        v.addWidget(self.en_label)

        self.input_line = QLineEdit()
        self.input_line.setStyleSheet(Style.INPUT)
        self.input_line.returnPressed.connect(self._on_next)
        v.addWidget(self.input_line)

        self.input_combo = QComboBox()
        self.input_combo.setStyleSheet(Style.INPUT)
        self.input_combo.setVisible(False)
        v.addWidget(self.input_combo)

        pick_row = QHBoxLayout()
        self.btn_pick_cell = QPushButton()
        self.btn_pick_cell.setStyleSheet(Style.BTN_GHOST)
        self.btn_pick_cell.clicked.connect(lambda: self._pick_from_excel("cell"))
        pick_row.addWidget(self.btn_pick_cell)
        self.btn_pick_range = QPushButton()
        self.btn_pick_range.setStyleSheet(Style.BTN_GHOST)
        self.btn_pick_range.clicked.connect(lambda: self._pick_from_excel("range"))
        pick_row.addWidget(self.btn_pick_range)
        pick_row.addStretch(1)
        v.addLayout(pick_row)

        self.hint_label = QLabel()
        self.hint_label.setStyleSheet(Style.LABEL_TIP)
        self.hint_label.setWordWrap(True)
        v.addWidget(self.hint_label)

        self.warn_label = QLabel()
        self.warn_label.setStyleSheet(Style.LABEL_WARN)
        self.warn_label.setWordWrap(True)
        v.addWidget(self.warn_label)

        # tips for this formula
        self.tips_label = QLabel()
        self.tips_label.setStyleSheet(Style.LABEL_TIP)
        self.tips_label.setWordWrap(True)
        self.tips_label.setVisible(False)
        v.addWidget(self.tips_label)

        v.addStretch(1)

        # bottom buttons
        bottom = QHBoxLayout()
        self.btn_cancel = QPushButton()
        self.btn_cancel.setStyleSheet(Style.BTN_GHOST)
        self.btn_cancel.clicked.connect(self.cancelled.emit)
        bottom.addWidget(self.btn_cancel)
        bottom.addStretch(1)
        self.btn_prev = QPushButton()
        self.btn_prev.setStyleSheet(Style.BTN_GHOST)
        self.btn_prev.clicked.connect(self._on_prev)
        bottom.addWidget(self.btn_prev)
        self.btn_next = QPushButton()
        self.btn_next.setStyleSheet(Style.BTN)
        bottom.addWidget(self.btn_next)
        v.addLayout(bottom)

    # ---------- wizard flow ----------
    def start(self, formula_name: str) -> None:
        f = get_formula(formula_name)
        if not f:
            return
        self.formula = f
        self.steps = f.args()
        self.values = {s.key: s.default for s in self.steps}
        self.current = 0
        self._render_step()
        self._render_tips()

    def _render_tips(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        if not self.formula:
            return
        tips = self.formula.zh_tips if lang == "zh" else self.formula.en_tips
        if tips:
            self.tips_label.setText("📘 " + "\n".join(f"• {t}" for t in tips))
            self.tips_label.setVisible(True)
        else:
            self.tips_label.setVisible(False)

    def _render_step(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        total = len(self.steps)
        step = self.steps[self.current]

        title_text = tr("wizard_step", lang=lang, cur=self.current + 1, tot=total,
                        title=step.zh_label if lang == "zh" else step.en_label)
        self.step_title.setText(title_text)
        self.zh_label.setText(step.zh_hint)
        self.en_label.setText(step.en_hint)

        value = self.values.get(step.key, step.default)
        self.input_line.setText(value)
        self.input_line.setVisible(step.kind in ("cell", "range", "text", "number"))

        if step.kind == "choice" and step.choices:
            self.input_combo.clear()
            for zh, en in step.choices:
                self.input_combo.addItem(zh if lang == "zh" else en, zh if lang == "zh" else en)
            # try to match current value
            idx = self.input_combo.findData(value)
            if idx >= 0:
                self.input_combo.setCurrentIndex(idx)
            self.input_combo.setVisible(True)
        else:
            self.input_combo.setVisible(False)

        # reset any pending pick state when moving to a new step
        self.cancel_pick()
        self.btn_pick_cell.setVisible(step.kind == "cell")
        self.btn_pick_range.setVisible(step.kind == "range")
        self.hint_label.setText(step.zh_hint if lang == "zh" else step.en_hint)
        self.warn_label.setText(step.zh_warning if lang == "zh" else step.en_warning)

        # prev button only when not on first
        self.btn_prev.setEnabled(self.current > 0)

        # next button label: "完成" on last step
        if self.current == total - 1:
            self.btn_next.setText(f"✓ {tr('wizard_write', lang=lang)}")
            self._connect_next(self._on_done)
        else:
            self.btn_next.setText(tr("wizard_next", lang=lang))
            self._connect_next(self._on_next)

        self._refresh_static_texts()

    def _connect_next(self, slot) -> None:
        """Wire btn_next to exactly one slot.

        Note: a bare self.btn_next.disconnect() raises TypeError in PySide6
        (“not enough arguments”) and disconnects NOTHING — doing that on every
        _render_step() piled up one extra _on_next per step, so a single click
        then jumped several steps at once.  Track the slot we wired instead.
        """
        if self._next_slot is not None:
            try:
                self.btn_next.clicked.disconnect(self._next_slot)
            except (TypeError, RuntimeError):
                pass
        self._next_slot = slot
        self.btn_next.clicked.connect(slot)

    def _refresh_static_texts(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        self.btn_cancel.setText(tr("wizard_cancel", lang=lang))
        self._update_pick_buttons()

    def _current_value(self) -> str:
        step = self.steps[self.current]
        if step.kind == "choice":
            return self.input_combo.currentData() or self.input_combo.currentText()
        return self.input_line.text()

    def _on_prev(self) -> None:
        if self.current > 0:
            self.current -= 1
            self._render_step()

    def _on_next(self) -> None:
        step = self.steps[self.current]
        val = self._current_value()
        err = step.validator(val)
        if err:
            Config.get()
            self.warn_label.setText(err)
            return
        self.values[step.key] = val
        if self.current < len(self.steps) - 1:
            self.current += 1
            self._render_step()

    def _on_done(self) -> None:
        # save last step value too
        step = self.steps[self.current]
        val = self._current_value()
        err = step.validator(val)
        if err:
            self.warn_label.setText(err)
            return
        self.values[step.key] = val
        self.next_ready.emit()

    def _pick_from_excel(self, mode: str) -> None:
        """Live, non-blocking cell pick — one selection in Excel confirms it.

        Flow:
          1. Click “Pick cell / Pick range” → we **minimise the sidecar** so
             Excel is fully in front and clickable, and start a 200 ms poll.
          2. The user may freely switch worksheets first (that counts as
             navigation, not a selection).  Once they click / drag-select the
             target and **release the mouse**, the address is filled in and the
             window pops back by itself — no “confirm” click needed.
          3. The button doubles as a manual “read now” fallback (used when the
             poll times out, e.g. the target cell was already selected).

        Why minimise instead of staying on top: an always-on-top frameless
        window physically sits over Excel and intercepts the click.
        Minimising is safe since v1.2.2 removed the Qt.Tool flag (there is a
        taskbar button, so the window can always be restored).
        """
        cfg = Config.get()
        lang = cfg.lang

        # ---- manual read (second click, or after the poll timed out) ----
        if self._pick_awaiting:
            self._read_pick_result(lang)
            return

        # ---- enter pick mode ----
        bridge = self.window().bridge  # type: ignore[attr-defined]
        if not bridge.is_connected():
            self.warn_label.setText(tr("diag_conn_fail", lang=lang))
            return

        self._pick_awaiting = True
        self._pick_mode = mode
        self._pick_base = bridge.read_selection(mode)
        self._pick_last = None
        self._pick_ticks = 0
        self._pick_pressed = False
        self._pick_click_done = False
        self._update_pick_buttons()
        self.warn_label.setText(tr("pick_mode_active", lang=lang))
        # Get the sidecar out of the way so Excel receives the click cleanly.
        self._minimise_for_pick()
        self._pick_timer.start()

    def _minimise_for_pick(self) -> None:
        win = self.window()
        if win is not None and not win.isMinimized():
            self._pick_minimized = True
            win.showMinimized()

    def _restore_window(self) -> None:
        """Bring the sidecar back after pick mode (no-op if we did not minimise)."""
        if not self._pick_minimized:
            return
        self._pick_minimized = False
        win = self.window()
        if win is None:
            return
        win.showNormal()
        win.raise_()
        win.activateWindow()

    def _read_pick_result(self, lang: str) -> None:
        """Read Excel's current selection once and fill it in."""
        bridge = self.window().bridge  # type: ignore[attr-defined]
        ref = bridge.read_selection_address(self._pick_mode)
        if ref:
            self._apply_pick(ref)
        else:
            self.cancel_pick()
            self.warn_label.setText(tr("diag_conn_fail", lang=lang))

    def _poll_pick(self) -> None:
        """QTimer tick: capture the selection once the user has finished choosing.

        Three guards, learned from real use:
          * **button must be up** — a column-header drag from B to D pauses on
            B:B long enough for a poll to fire; confirming there froze the
            range at one column, and popping the window up mid-drag aborted
            the drag altogether;
          * **address stable across two polls** — a click confirms in ~400 ms;
          * **worksheet must not change** — clicking another sheet tab is
            navigation, not a pick.  Confirming on it was why “pick a range on
            Sheet2” was impossible: the moment Sheet2 was activated its A1 was
            captured and the window jumped back.
        """
        lang = Config.get().lang
        if not self.isVisible():
            self.cancel_pick()
            return
        self._pick_ticks += 1
        if self._pick_ticks > 150:  # 150 * 200 ms = 30 s
            # Timed out: pop the window back but STAY in pick mode, so the
            # user can click “Read selected cell” to use whatever is currently
            # selected (covers the “pick the cell that was already active” case).
            self._pick_timer.stop()
            self._restore_window()
            self._pick_awaiting = True
            self._update_pick_buttons()
            self.warn_label.setText(tr("pick_no_change", lang=lang))
            return

        bridge = self.window().bridge  # type: ignore[attr-defined]
        snap = bridge.read_selection(self._pick_mode)

        down = _left_button_down()
        if down:
            self._pick_pressed = True
        elif self._pick_pressed:
            self._pick_pressed = False
            self._pick_click_done = True

        if snap is None:
            return
        sig = snap["ref"]
        base = self._pick_base
        if base is None:
            # No baseline yet (Excel had no active cell at entry).
            self._pick_base = snap
            self._pick_last = sig
            return
        if snap["sheet"] != base["sheet"]:
            # Worksheet switch → navigation.  Re-baseline, never confirm.
            self._pick_base = snap
            self._pick_last = sig
            self._pick_click_done = False
            return
        if down:
            self._pick_last = sig  # mid-drag: wait for the button to come up
            return
        if sig != self._pick_last:
            self._pick_last = sig  # changed this tick → require one stable poll
            return
        # Same sheet, stable, button released.  Confirm a new address, or an
        # explicit re-click of the cell that was already selected.
        if sig != base["ref"] or self._pick_click_done:
            self._apply_pick(sig)

    def _apply_pick(self, addr: str) -> None:
        self.cancel_pick()  # stops the timer and restores the window
        self.input_line.setText(addr)
        self.values[self.steps[self.current].key] = addr
        self.warn_label.setText("")

    def cancel_pick(self) -> None:
        """Leave pick mode: stop polling and restore the normal window state."""
        self._pick_timer.stop()
        self._restore_window()
        self._pick_awaiting = False
        self._pick_base = None
        self._pick_last = None
        self._pick_ticks = 0
        self._pick_pressed = False
        self._pick_click_done = False
        self._update_pick_buttons()

    def _update_pick_buttons(self) -> None:
        lang = Config.get().lang
        if self._pick_awaiting:
            self.btn_pick_cell.setText(tr("pick_read_now", lang=lang))
            self.btn_pick_range.setText(tr("pick_read_now", lang=lang))
        else:
            self.btn_pick_cell.setText(tr("wizard_pick_cell", lang=lang))
            self.btn_pick_range.setText(tr("wizard_pick_range", lang=lang))

    def apply_style(self) -> None:
        """Re-apply every widget-level stylesheet after a theme change."""
        self.step_title.setStyleSheet(Style.LABEL_STEP)
        self.zh_label.setStyleSheet(Style.LABEL)
        self.en_label.setStyleSheet(Style.LABEL)
        self.input_line.setStyleSheet(Style.INPUT)
        self.input_combo.setStyleSheet(Style.INPUT)
        self.btn_pick_cell.setStyleSheet(Style.BTN_GHOST)
        self.btn_pick_range.setStyleSheet(Style.BTN_GHOST)
        self.hint_label.setStyleSheet(Style.LABEL_TIP)
        self.warn_label.setStyleSheet(Style.LABEL_WARN)
        self.tips_label.setStyleSheet(Style.LABEL_TIP)
        self.btn_cancel.setStyleSheet(Style.BTN_GHOST)
        self.btn_prev.setStyleSheet(Style.BTN_GHOST)
        self.btn_next.setStyleSheet(Style.BTN)

    def refresh_texts(self) -> None:
        self._refresh_static_texts()
        if self.formula:
            self._render_step()
            self._render_tips()


# ==========================================================================
# Preview screen — show generated formula, write to Excel, or copy
# ==========================================================================
class PreviewScreen(QWidget):
    write_requested = Signal()
    copy_requested = Signal()
    go_home = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.formula: FormulaBase | None = None
        self.generated: str = ""  # body without =
        self._build_ui()

    def _build_ui(self) -> None:
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 8, 14, 10)
        v.setSpacing(8)

        self.title = QLabel()
        self.title.setStyleSheet(Style.LABEL_TITLE)
        v.addWidget(self.title)

        self.preview_box = QLabel()
        self.preview_box.setStyleSheet(Style.PREVIEW_BOX)
        self.preview_box.setWordWrap(True)
        self.preview_box.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        v.addWidget(self.preview_box)

        self.status_label = QLabel()
        self.status_label.setStyleSheet(Style.LABEL)
        v.addWidget(self.status_label)

        # diagnostic section
        self.diag_head = QLabel()
        self.diag_head.setStyleSheet(f"color: {_WARN}; font-weight: bold; padding-top: 6px;")
        self.diag_head.setVisible(False)
        v.addWidget(self.diag_head)
        self.diag_text = QLabel()
        self.diag_text.setStyleSheet(Style.LABEL)
        self.diag_text.setWordWrap(True)
        self.diag_text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.diag_text.setVisible(False)
        v.addWidget(self.diag_text)
        self.diag_fix_btn = QPushButton()
        self.diag_fix_btn.setStyleSheet(Style.BTN_GHOST)
        self.diag_fix_btn.setVisible(False)
        self.diag_fix_btn.clicked.connect(self._apply_diag_fix)
        v.addWidget(self.diag_fix_btn)

        v.addStretch(1)

        bottom = QHBoxLayout()
        btn_home = QPushButton()
        btn_home.setStyleSheet(Style.BTN_GHOST)
        btn_home.clicked.connect(self.go_home.emit)
        bottom.addWidget(btn_home)
        bottom.addStretch(1)
        btn_copy = QPushButton()
        btn_copy.setStyleSheet(Style.BTN_GHOST)
        btn_copy.clicked.connect(self.copy_requested.emit)
        bottom.addWidget(btn_copy)
        self.btn_write = QPushButton()
        self.btn_write.setStyleSheet(Style.BTN)
        self.btn_write.clicked.connect(self.write_requested.emit)
        bottom.addWidget(self.btn_write)
        v.addLayout(bottom)

        self._btn_home_ref = btn_home
        self._btn_copy_ref = btn_copy

    # ---------- runtime ----------
    def set_formula(self, formula_name: str, generated_body: str) -> None:
        f = get_formula(formula_name)
        if not f:
            return
        self.formula = f
        self.generated = generated_body
        self.refresh_texts()

    def set_diagnostics(self, error: str | None, suggestions: list[Any]) -> None:
        cfg = Config.get()
        lang = cfg.lang
        if not error:
            self.status_label.setText(tr("wizard_ok", lang=lang))
            self.diag_head.setVisible(False)
            self.diag_text.setVisible(False)
            self.diag_fix_btn.setVisible(False)
            return
        self.diag_head.setText(tr("diag_title", lang=lang) + f"：{error}")
        parts = []
        fix_fomula = None
        for i, s in enumerate(suggestions, 1):
            title = s.zh_title if lang == "zh" else s.en_title
            detail = s.zh_detail if lang == "zh" else s.en_detail
            parts.append(f"• {title}\n  {detail}")
            if s.zh_fix or s.en_fix:
                fix_fomula = s.zh_fix or s.en_fix
        self.diag_text.setText("\n\n".join(parts))
        self.diag_head.setVisible(True)
        self.diag_text.setVisible(True)
        if fix_fomula:
            self.diag_fix_btn.setVisible(True)
            self._fix_formula_str = fix_fomula
        else:
            self.diag_fix_btn.setVisible(False)

    def _apply_diag_fix(self) -> None:
        if getattr(self, "_fix_formula_str", None):
            # _fix_formula_str already has full body (no leading '=')
            self.generated = self._fix_formula_str
            # re-render preview
            self.preview_box.setText("=" + self.generated)
            self.diag_head.setVisible(False)
            self.diag_text.setVisible(False)
            self.diag_fix_btn.setVisible(False)
            cfg = Config.get()
            self.status_label.setText(tr("wizard_ok", lang=cfg.lang))

    def refresh_texts(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        if self.formula:
            self.title.setText(
                tr("wizard_done", lang=lang) + " " +
                (self.formula.zh_name if lang == "zh" else self.formula.en_name)
            )
        self.preview_box.setText("=" + self.generated)
        self.btn_write.setText(tr("wizard_write", lang=lang))
        self._btn_copy_ref.setText(tr("wizard_copy", lang=lang))
        self._btn_home_ref.setText(tr("wizard_goto_home", lang=lang))
        self._refresh_diag_texts()

    def _refresh_diag_texts(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        if self.diag_fix_btn.isVisible():
            self.diag_fix_btn.setText(tr("diag_fix_action", lang=lang))

    def apply_style(self) -> None:
        """Re-apply every widget-level stylesheet after a theme change."""
        self.title.setStyleSheet(Style.LABEL_TITLE)
        self.preview_box.setStyleSheet(Style.PREVIEW_BOX)
        self.status_label.setStyleSheet(Style.LABEL)
        self.diag_head.setStyleSheet(f"color: {_WARN}; font-weight: bold; padding-top: 6px;")
        self.diag_text.setStyleSheet(Style.LABEL)
        self.diag_fix_btn.setStyleSheet(Style.BTN_GHOST)
        self._btn_home_ref.setStyleSheet(Style.BTN_GHOST)
        self._btn_copy_ref.setStyleSheet(Style.BTN_GHOST)
        self.btn_write.setStyleSheet(Style.BTN)


# ==========================================================================
# Main window — borderless, draggable, always on top, with custom title bar
# ==========================================================================
class SidecarWindow(QWidget):
    def __init__(self, bridge: ExcelComBridge) -> None:
        super().__init__()
        self.bridge = bridge
        self._drag_pos = None
        self._build_window()
        self._build_screens()
        self._connect_signals()
        self._refresh_static_texts()
        self._apply_geometry()
        self._refresh_status()

    # ---------- window chrome ----------
    def _build_window(self) -> None:
        # NOTE: no Qt.Tool here.  A tool window gets no taskbar button, so
        # showMinimized() (title-bar “—”) would hide the app with no way to
        # bring it back — exactly the “the app just died” report.
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setStyleSheet(Style.WINDOW)
        self.setMinimumSize(380, 500)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.frame = QWidget()
        self.frame.setObjectName("sidecarFrame")
        self.frame.setStyleSheet(
            f"QWidget#sidecarFrame {{ background: {_BG}; border-radius: 12px; "
            f"border: 1px solid {_BORDER}; }}"
        )
        v = QVBoxLayout(self.frame)
        v.setContentsMargins(10, 6, 10, 6)
        v.setSpacing(4)

        # title bar
        title_row = QHBoxLayout()
        self.title_label = QLabel()
        title_label = self.title_label
        title_label.setStyleSheet(f"color: {_TEXT}; font-weight: bold; font-size: 13px; padding: 0 4px;")
        title_row.addWidget(title_label, stretch=1)
        self.btn_theme = QPushButton()
        self.btn_theme.setStyleSheet(Style.BTN_TITLE)
        self.btn_theme.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_theme.clicked.connect(self._show_theme_menu)
        title_row.addWidget(self.btn_theme)
        self.btn_lang = QPushButton()
        self.btn_lang.setStyleSheet(Style.BTN_TITLE)
        self.btn_lang.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lang.clicked.connect(self._toggle_lang)
        title_row.addWidget(self.btn_lang)
        self.btn_min = QPushButton()
        self.btn_min.setStyleSheet(Style.BTN_TITLE)
        self.btn_min.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_min.clicked.connect(self.showMinimized)
        title_row.addWidget(self.btn_min)
        self.btn_close = QPushButton()
        self.btn_close.setStyleSheet(Style.BTN_TITLE)
        self.btn_close.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_close.clicked.connect(self._on_close)
        title_row.addWidget(self.btn_close)
        v.addLayout(title_row)

        # content (stacked)
        self.stack = QStackedWidget()
        v.addWidget(self.stack, stretch=1)

        # status bar
        self.status = QLabel()
        self.status.setStyleSheet(Style.STATUS)
        v.addWidget(self.status)

        outer.addWidget(self.frame)

    def _build_screens(self) -> None:
        self.home = HomeScreen()
        self.wizard = WizardScreen()
        self.preview = PreviewScreen()
        self.stack.addWidget(self.home)
        self.stack.addWidget(self.wizard)
        self.stack.addWidget(self.preview)
        self.stack.setCurrentWidget(self.home)

    def _connect_signals(self) -> None:
        self.home.open_wizard.connect(self._open_wizard)
        self.wizard.next_ready.connect(self._show_preview)
        self.wizard.cancelled.connect(self._go_home)
        self.preview.go_home.connect(self._go_home)
        self.preview.write_requested.connect(self._write_to_excel)
        self.preview.copy_requested.connect(self._copy_formula)

    def _refresh_static_texts(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        self.title_label.setText(tr("app_title", lang=lang, ver=__version__))
        self.btn_lang.setText(tr("btn_lang_zh", lang=lang) + " / " + tr("btn_lang_en", lang=lang))
        self.btn_theme.setText(tr("btn_theme", lang=lang))
        self.btn_min.setText(tr("btn_min", lang=lang))
        self.btn_close.setText(tr("btn_close", lang=lang))
        self.home.refresh_texts()
        self.wizard.refresh_texts()
        self.preview.refresh_texts()

    # ---------- geometry ----------
    def _apply_geometry(self) -> None:
        cfg = Config.get()
        if cfg.window_geom:
            x, y, w, h = cfg.window_geom
            self.setGeometry(x, y, w, h)
        else:
            # default top-right corner
            screen = QGuiApplication.primaryScreen().availableGeometry()
            w, h = 420, 600
            self.setGeometry(screen.right() - w - 20, screen.top() + 40, w, h)

    # ---------- theme ----------
    def _show_theme_menu(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        menu = QMenu(self)
        act_purple = menu.addAction(tr("theme_purple", lang=lang))
        act_purple.setCheckable(True)
        act_purple.setChecked(cfg.theme.upper() == DEFAULT_BASE)
        act_auto = menu.addAction(tr("theme_auto", lang=lang))
        act_auto.setCheckable(True)
        act_auto.setChecked(cfg.theme == "auto")
        act_custom = menu.addAction(tr("theme_custom", lang=lang))
        chosen = menu.exec(self.btn_theme.mapToGlobal(self.btn_theme.rect().bottomLeft()))
        if chosen is act_purple:
            cfg.theme = DEFAULT_BASE
            cfg.save()
            self._apply_theme()
        elif chosen is act_auto:
            cfg.theme = "auto"
            cfg.save()
            self._apply_theme()
        elif chosen is act_custom:
            color = QColorDialog.getColor(QColor(_BG), self, tr("theme_pick_title", lang=lang))
            if color.isValid():
                cfg.theme = color.name().upper()
                cfg.save()
                self._apply_theme()
        if chosen is not None:
            self.status.setText(tr("theme_applied", lang=cfg.lang))

    def _apply_theme(self) -> None:
        """Re-derive every colour from the stored base and repaint the UI."""
        cfg = Config.get()
        set_base_color(None if cfg.theme == "auto" else cfg.theme)
        self.setStyleSheet(Style.WINDOW)
        self.frame.setStyleSheet(
            f"QWidget#sidecarFrame {{ background: {_BG}; border-radius: 12px; "
            f"border: 1px solid {_BORDER}; }}"
        )
        self.home.apply_style()
        self.wizard.apply_style()
        self.preview.apply_style()
        self.status.setStyleSheet(Style.STATUS)
        self._refresh_static_texts()  # rebuilds the formula cards in new colours

    # ---------- language ----------
    def _toggle_lang(self) -> None:
        cfg = Config.get()
        cfg.lang = "en" if cfg.lang == "zh" else "zh"
        cfg.save()
        self._refresh_static_texts()

    # ---------- wizard flow ----------
    def _open_wizard(self, name: str) -> None:
        self.wizard.start(name)
        self.stack.setCurrentWidget(self.wizard)

    def _show_preview(self) -> None:
        self.wizard.cancel_pick()
        name = self.wizard.formula.name if self.wizard.formula else ""
        body = self.wizard.formula.build(**self.wizard.values) if self.wizard.formula else ""
        Config.get().add_recent(name)
        self.preview.set_formula(name, body)
        # Read current cell error (sync call — very fast thanks to _with_protection)
        if self.bridge.is_connected():
            addr = self.bridge.get_active_cell_address().get("data", "")
            err = self.bridge.get_cell_error(addr) if addr else None
            suggestions = self.wizard.formula.diagnose(body, err) if self.wizard.formula else []
            self.preview.set_diagnostics(err, suggestions)
        self.stack.setCurrentWidget(self.preview)

    def _go_home(self) -> None:
        self.wizard.cancel_pick()
        self.home.search.clear()
        self.home._render_cards("")
        self.stack.setCurrentWidget(self.home)

    # ---------- excel write ----------
    def _write_to_excel(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        body = self.preview.generated
        if not self.bridge.is_connected():
            QMessageBox.warning(self, "Error", tr("diag_conn_fail", lang=lang))
            return
        addr_result = self.bridge.get_active_cell_address()
        if not addr_result.get("ok"):
            QMessageBox.warning(self, "Error", str(addr_result.get("error", "unknown")))
            return
        addr = str(addr_result["data"])
        write_result = self.bridge.set_formula(addr, body)
        if write_result.get("ok"):
            name = self.bridge.workbook_name()
            suffix = f" · {name}" if name else ""
            self.status.setText(f"✓ {tr('written', lang=lang)} — {addr}{suffix}")
            # re-check for errors post-write
            err = self.bridge.get_cell_error(addr)
            suggestions = self.preview.formula.diagnose(body, err) if self.preview.formula else []
            self.preview.set_diagnostics(err, suggestions)
        else:
            QMessageBox.critical(self, "Error", str(write_result.get("error", "Unknown")))

    def _copy_formula(self) -> None:
        cfg = Config.get()
        cb = QApplication.clipboard()
        cb.setText("=" + self.preview.generated)
        self.status.setText(tr("copied", lang=cfg.lang))

    def _refresh_status(self) -> None:
        cfg = Config.get()
        lang = cfg.lang
        ok = self.bridge.is_connected()
        if ok:
            name = self.bridge.workbook_name()
            self.status.setText(tr("diag_conn_ok", lang=lang) + (f" · {name}" if name else ""))
        else:
            self.status.setText(tr("diag_conn_fail", lang=lang))

    # ---------- close / save geometry ----------
    def _on_close(self) -> None:
        cfg = Config.get()
        geom = self.geometry()
        cfg.window_geom = (geom.x(), geom.y(), geom.width(), geom.height())
        cfg.save()
        self.close()

    # ---------- drag & drop ----------
    def mousePressEvent(self, event):  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):  # noqa: N802
        if self._drag_pos is not None and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

    def mouseReleaseEvent(self, event):  # noqa: N802
        self._drag_pos = None


# ==========================================================================
# Entry: creates bridge, connects to Excel if possible, launches the window
# ==========================================================================
def launch() -> int:
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    app.setStyle("Fusion")

    bridge = ExcelComBridge()
    cfg = Config.get()
    set_base_color(None if cfg.theme == "auto" else cfg.theme)  # before widgets are built
    window = SidecarWindow(bridge)
    window.show()
    bridge.co_init()  # one-time COM apartment init on GUI thread
    if cfg.auto_connect:
        result = bridge.connect()
        window._refresh_status()
        if not result.get("ok") and result.get("error") == "excel_busy":
            window.status.setText(tr("diag_excel_busy", lang=cfg.lang))

    def _cleanup() -> None:
        cfg = Config.get()
        bridge.co_uninit()
        cfg.save()

    app.aboutToQuit.connect(_cleanup)
    return app.exec()
