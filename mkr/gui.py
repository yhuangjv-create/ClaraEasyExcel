"""PySide6 sidecar window — floating, always-on-top, borderless with custom title bar.

Three screens stacked in a QStackedWidget: Home → Wizard → Preview.
Language button toggles zh/en by mutating Config.lang and calling refresh_texts().
"""
from __future__ import annotations

import sys
from typing import Any

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from mkr import __version__
from mkr.config import Config
from mkr.excel_bridge import ExcelComBridge
from mkr.formulas import base as fbase  # noqa: F401 — triggers registry
from mkr.formulas.base import FormulaBase, all_formula_names, get_formula
from mkr.i18n import tr
from mkr.intent import detect

# ---------- Windows 11 “deep orange-yellow” palette — NO green anywhere ----------
_BG = "#F7630C"        # window base — Win11 deep orange-yellow accent
_BG_CARD = "#FFF1DE"   # cream card so dark ink text stays readable
_BG_DEEP = "#FFF8F0"   # inputs / preview boxes
_BORDER = "#8A3B00"    # burnt-orange border
_ACCENT = "#8A3B00"    # primary button / selected surface
_ACCENT_HOVER = "#A64A00"
_TEXT = "#331A00"      # dark ink
_TEXT_DIM = "#4A2600"
_WARN = "#5A0A00"      # deep red-brown — never green
_OK = "#5C3A14"


class Style:
    """Centralize all QSS strings so changing the palette is one edit."""
    WINDOW = f"""
        QWidget {{ background: {_BG}; color: {_TEXT}; font-family: "Microsoft YaHei", "Segoe UI"; font-size: 13px; }}
    """
    CARD = f"""
        QFrame {{ background: {_BG_CARD}; border-radius: 10px; padding: 12px; }}
        QFrame:hover {{ background: #FFD9AE; }}
    """
    BTN = f"""
        QPushButton {{ background: {_ACCENT}; color: white; border: none; border-radius: 6px;
                       padding: 6px 14px; font-weight: bold; }}
        QPushButton:hover {{ background: {_ACCENT_HOVER}; }}
        QPushButton:disabled {{ background: #D9C6B0; color: #8A7460; }}
    """
    BTN_GHOST = f"""
        QPushButton {{ background: transparent; color: {_TEXT}; border: 1px solid {_BORDER};
                       border-radius: 6px; padding: 5px 12px; }}
        QPushButton:hover {{ background: {_ACCENT}; color: white; }}
    """
    BTN_TITLE = f"""
        QPushButton {{ background: transparent; color: {_TEXT_DIM}; border: none; font-size: 16px;
                       padding: 2px 8px; }}
        QPushButton:hover {{ color: {_TEXT}; background: #FF8A3D; border-radius: 4px; }}
    """
    INPUT = f"""
        QLineEdit, QComboBox, QPlainTextEdit {{
            background: {_BG_DEEP}; border: 1px solid {_BORDER}; border-radius: 6px;
            color: {_TEXT}; padding: 6px; selection-background-color: #FFD9AE;
        }}
        QLineEdit:focus, QComboBox:focus, QPlainTextEdit:focus {{ border: 1px solid {_ACCENT_HOVER}; }}
    """
    LABEL = f"color: {_TEXT_DIM}; font-size: 12px;"
    LABEL_TITLE = f"color: {_TEXT}; font-size: 16px; font-weight: bold;"
    LABEL_STEP = f"color: {_TEXT}; font-size: 14px; font-weight: bold;"
    LABEL_TIP = f"color: {_TEXT_DIM}; font-size: 11px; padding: 2px 4px;"
    LABEL_WARN = f"color: {_WARN}; font-size: 11px; padding: 2px 4px;"
    LIST = f"""
        QListWidget {{ background: {_BG}; border: none; outline: none; }}
        QListWidget::item {{ background: {_BG_CARD}; border-radius: 8px; padding: 10px; margin: 3px 0; }}
        QListWidget::item:hover {{ background: #FFE2C4; color: {_TEXT}; }}
        QListWidget::item:selected {{ background: #FFD0A4; color: {_TEXT}; }}
    """
    PREVIEW_BOX = f"""
        QFrame {{ background: {_BG_DEEP}; border: 1px solid {_BORDER}; border-radius: 8px;
                  padding: 10px; font-family: "Consolas", "Courier New"; font-size: 13px; color: {_TEXT}; }}
    """
    STATUS = f"color: {_TEXT_DIM}; font-size: 11px;"


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


def _build_card(f: FormulaBase, lang: str) -> QWidget:
    card = QWidget()
    card.setStyleSheet(
        f"QWidget {{ background: {_BG_CARD}; border-radius: 8px; }} "
        "QWidget:hover { background: #FFD9AE; }"
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
        # Live, non-blocking cell picking (see _pick_from_excel):
        #   _pick_awaiting=True → the poll timer below watches Excel's selection
        self._pick_awaiting = False
        self._pick_mode = "cell"
        self._pick_base: str | None = None   # selection when pick mode started
        self._pick_last: str | None = None   # previous poll result (debounce)
        self._pick_ticks = 0
        self._pick_timer = QTimer(self)
        self._pick_timer.setInterval(500)
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
        self.btn_next.clicked.connect(self._on_next)
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
            self.btn_next.setText("✓ " + (tr("wizard_done", lang=lang)))
            self.btn_next.setText(f"✓ {tr('wizard_write', lang=lang)}")
            # switch to next_ready signal — window will move to preview itself
            try:
                self.btn_next.disconnect()
            except TypeError:
                pass
            self.btn_next.setStyleSheet(Style.BTN)
            self.btn_next.clicked.connect(self._on_done)
        else:
            try:
                self.btn_next.disconnect()
            except TypeError:
                pass
            self.btn_next.setText(tr("wizard_next", lang=lang))
            self.btn_next.setStyleSheet(Style.BTN)
            self.btn_next.clicked.connect(self._on_next)

        self._refresh_static_texts()

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
        """Live, non-blocking cell pick.

        Why this shape (the old version could not pick and froze the app):
          * It opened a **modal QMessageBox** on top of an always-on-top,
            frameless window.  The modal dialog spawned a nested Qt event
            loop, covered Excel so the user could not click a cell, and on
            close immediately fired a COM read while that loop was unwinding
            — deadlocking the UI thread.
          * It also used showMinimized().  Our window is a Qt.Tool window,
            which has **no taskbar button**, so once minimised the user had
            no way to bring it back — the app looked dead.

        New flow: the window stays put.  A QTimer polls Excel's selection
        every 500 ms; as soon as the user clicks a different cell (stable for
        two polls, so drag-selecting a range works) the address is filled in
        automatically.  The button doubles as a manual "read now" fallback.
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
        self._pick_base = bridge.read_selection_address(mode)
        self._pick_last = None
        self._pick_ticks = 0
        self._update_pick_buttons()
        self.warn_label.setText(tr("pick_mode_active", lang=lang))
        self._pick_timer.start()

    def _read_pick_result(self, lang: str) -> None:
        """Read Excel's current selection once and fill it in."""
        bridge = self.window().bridge  # type: ignore[attr-defined]
        result = (
            bridge.get_active_range_address()
            if self._pick_mode == "range"
            else bridge.get_active_cell_address()
        )
        if result.get("ok") and result.get("data"):
            self._apply_pick(str(result["data"]))
        else:
            self.cancel_pick()
            self.warn_label.setText(str(result.get("error", tr("diag_conn_fail", lang=lang))))

    def _poll_pick(self) -> None:
        """QTimer tick: auto-capture when the user selects a new cell."""
        lang = Config.get().lang
        if not self.isVisible():
            self.cancel_pick()
            return
        self._pick_ticks += 1
        if self._pick_ticks > 60:  # ~30 s
            self.cancel_pick()
            self.warn_label.setText(tr("pick_no_change", lang=lang))
            return

        bridge = self.window().bridge  # type: ignore[attr-defined]
        addr = bridge.read_selection_address(self._pick_mode)
        if not addr or self._pick_base is None:
            # no baseline → we cannot tell "new" from "old"; manual read only
            return
        # unchanged selection → keep waiting (also covers "pick the same cell")
        if addr == self._pick_base:
            self._pick_last = addr
            return
        # require the address to be stable across two polls so drag-selecting
        # a multi-cell range is not captured halfway through the drag
        if addr == self._pick_last:
            self._apply_pick(addr)
            return
        self._pick_last = addr

    def _apply_pick(self, addr: str) -> None:
        self.cancel_pick()
        self.input_line.setText(addr)
        self.values[self.steps[self.current].key] = addr
        self.warn_label.setText("")

    def cancel_pick(self) -> None:
        """Leave pick mode: stop polling and restore the normal button text."""
        self._pick_timer.stop()
        self._pick_awaiting = False
        self._pick_base = None
        self._pick_last = None
        self._pick_ticks = 0
        self._update_pick_buttons()

    def _update_pick_buttons(self) -> None:
        lang = Config.get().lang
        if self._pick_awaiting:
            self.btn_pick_cell.setText(tr("pick_read_now", lang=lang))
            self.btn_pick_range.setText(tr("pick_read_now", lang=lang))
        else:
            self.btn_pick_cell.setText(tr("wizard_pick_cell", lang=lang))
            self.btn_pick_range.setText(tr("wizard_pick_range", lang=lang))

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
            f"border: 1px solid #FFD9AE; }}"
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
            self.status.setText(f"✓ {tr('written', lang=lang)} — {addr}")
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
        self.status.setText(
            (tr("diag_conn_ok", lang=lang) if ok else tr("diag_conn_fail", lang=lang))
        )

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
    window = SidecarWindow(bridge)
    window.show()
    bridge.co_init()  # one-time COM apartment init on GUI thread
    cfg = Config.get()
    if cfg.auto_connect:
        bridge.connect()
        window._refresh_status()

    def _cleanup() -> None:
        cfg = Config.get()
        bridge.co_uninit()
        cfg.save()

    app.aboutToQuit.connect(_cleanup)
    return app.exec()
