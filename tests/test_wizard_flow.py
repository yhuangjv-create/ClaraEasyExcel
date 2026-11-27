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


def test_back_button_left_of_next_is_labelled(app):
    """The button left of “Next” must read “Back” — it used to be blank.

    btn_prev was created and wired to _on_prev, but _refresh_static_texts()
    never set its text, so the wizard showed an empty ghost button.
    """
    from mkr.config import Config
    from mkr.i18n import tr

    win = SidecarWindow(FakeBridge())
    win.wizard.start("VLOOKUP")
    expected = tr("wizard_prev", lang=Config.get().lang)
    assert win.wizard.btn_prev.text() == expected
    assert expected.strip() != ""

    # it also has to survive a language switch (in-memory only, no config write)
    cfg = Config.get()
    original = cfg.lang
    try:
        cfg.lang = "en" if original == "zh" else "zh"
        win.wizard.refresh_texts()
        assert win.wizard.btn_prev.text() == tr("wizard_prev", lang=cfg.lang)
    finally:
        cfg.lang = original
        win.wizard.refresh_texts()
