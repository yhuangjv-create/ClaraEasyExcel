"""Regression: picking a range must survive worksheet switches and long drags.

Reported against 国开行转账明细.xlsx, wizard step 2 (框选区域):
  * the target range was B:D on **Sheet2**, but simply clicking the Sheet2 tab
    was captured as the pick (its A1), so another worksheet could never be
    chosen;
  * dragging across the column headers B→D was cut short at a single column,
    because the mid-drag selection (B:B) stayed stable for two polls and the
    window was popped back to the front, aborting the drag.

Fix under test (gui.WizardScreen._poll_pick):
  * never confirm while the left button is held down;
  * a worksheet change is navigation → re-baseline, never confirm;
  * confirm only a stable, same-sheet, button-released selection, and use the
    sheet-qualified reference (Sheet2!B:D).
"""

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication

import mkr.gui as gui
from mkr.gui import SidecarWindow


class PickBridge:
    """Fake Excel: `snap` mirrors what Selection/ActiveCell would report."""

    def __init__(self) -> None:
        self.snap = {"sheet": "Sheet1", "address": "B2", "ref": "Sheet1!B2"}

    def set(self, sheet: str, address: str) -> None:
        self.snap = {"sheet": sheet, "address": address, "ref": f"{sheet}!{address}"}

    def is_connected(self) -> bool:
        return True

    def workbook_name(self) -> str:
        return "国开行转账明细.xlsx"

    def read_selection(self, mode: str = "cell"):
        return dict(self.snap)

    def read_selection_address(self, mode: str = "cell"):
        return self.snap["ref"]


@pytest.fixture(scope="module")
def app():
    instance = QApplication.instance()
    return instance or QApplication([])


@pytest.fixture
def mouse(monkeypatch):
    """Drive the real Win32 button state from the test."""
    state = {"down": False}
    monkeypatch.setattr(gui, "_left_button_down", lambda: state["down"])
    return state


def _make_window(app, bridge):
    win = SidecarWindow(bridge)
    win.stack.setCurrentWidget(win.wizard)
    win.show()
    for _ in range(3):
        app.processEvents()
    win.wizard.start("VLOOKUP")
    win.wizard.current = 1          # step 2 = 查找表区域 (range)
    win.wizard._render_step()
    win.wizard._pick_from_excel("range")
    for _ in range(3):
        app.processEvents()
    return win


def test_switching_worksheet_is_not_a_pick(app, mouse):
    bridge = PickBridge()
    win = _make_window(app, bridge)
    wizard = win.wizard
    assert wizard._pick_awaiting is True

    # user clicks the Sheet2 tab: the new sheet's active cell is A1
    bridge.set("Sheet2", "A1")
    for _ in range(4):
        wizard._poll_pick()
    assert wizard._pick_awaiting is True, "sheet switch must not confirm the pick"
    assert wizard.input_line.text() != "Sheet2!A1"

    wizard.cancel_pick()


def test_drag_is_captured_only_after_release(app, mouse):
    bridge = PickBridge()
    win = _make_window(app, bridge)
    wizard = win.wizard

    # user switches to Sheet2 first
    bridge.set("Sheet2", "A1")
    wizard._poll_pick()
    wizard._poll_pick()

    # starts dragging the column headers: B:B → B:C → B:D, button still down
    mouse["down"] = True
    for addr in ("B:B", "B:B", "B:C", "B:D", "B:D"):
        bridge.set("Sheet2", addr)
        wizard._poll_pick()
        assert wizard._pick_awaiting is True, "must not confirm mid-drag"
    assert win.isMinimized() is True

    # releases the button → one more stable poll → confirmed
    mouse["down"] = False
    bridge.set("Sheet2", "B:D")
    wizard._poll_pick()
    wizard._poll_pick()
    for _ in range(3):
        app.processEvents()

    assert wizard._pick_awaiting is False
    assert wizard.input_line.text() == "Sheet2!B:D"
    assert win.isMinimized() is False


def test_manual_read_uses_sheet_qualified_ref(app, mouse):
    bridge = PickBridge()
    win = _make_window(app, bridge)
    wizard = win.wizard
    wizard._pick_timer.stop()          # simulate the 30 s timeout fallback
    wizard._restore_window()
    wizard._pick_awaiting = True

    bridge.set("Sheet2", "B:D")
    wizard._pick_from_excel("range")   # second click = manual read
    assert wizard.input_line.text() == "Sheet2!B:D"
    assert wizard._pick_awaiting is False
