"""Regression: the wizard must advance exactly one step per click.

Bug (v1.2.2 and earlier): _render_step() called a bare btn_next.disconnect(),
which raises TypeError in PySide6 and disconnects nothing, so every rendered
step added one more _on_next connection.  A single click then jumped several
steps (e.g. 1 → 4 after coming back to the home screen).
"""

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication

from mkr.gui import SidecarWindow


class FakeBridge:
    def is_connected(self) -> bool:
        return False

    def workbook_name(self) -> str:
        return ""


@pytest.fixture(scope="module")
def app():
    instance = QApplication.instance()
    return instance or QApplication([])


def test_one_click_advances_one_step(app):
    win = SidecarWindow(FakeBridge())
    win.wizard.start("VLOOKUP")
    assert win.wizard.current == 0
    for expected in (1, 2, 3):
        win.wizard.btn_next.click()
        assert win.wizard.current == expected, "one click must advance one step"
        assert win.wizard.btn_next.text().strip() != ""


def test_reopen_resets_to_first_step(app):
    win = SidecarWindow(FakeBridge())
    win.wizard.start("VLOOKUP")
    for _ in range(3):
        win.wizard.btn_next.click()

    win._go_home()
    win._open_wizard("VLOOKUP")

    assert win.wizard.current == 0
    win.wizard.btn_next.click()
    assert win.wizard.current == 1
