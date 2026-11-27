"""Regression: hovering a formula card highlights the whole card, not one line.

Bug: the card stylesheet used a bare `QWidget` selector, which also matched
every QLabel inside the card.  Each line therefore painted its own rounded
background and had its own :hover, so sweeping the mouse down a 3-line
description lit up one line at a time instead of the card as a block.
"""

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel, QWidget

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


def test_card_text_is_never_squeezed(app):
    """Hover fix aside: the card must never be shorter than its own layout.

    v1.3.6 sized list rows from the card's preferred height, but the item style
    still had 6 px padding, so the card received a shorter rect than allocated;
    Qt then squeezed the labels and clipped the top of the bold headline.
    """
    win = SidecarWindow(FakeBridge())
    win.resize(320, 420)
    win.show()
    win.stack.setCurrentWidget(win.home)
    for _ in range(6):
        app.processEvents()

    cards = win.home.list_widget.findChildren(QWidget, "formulaCard")
    assert cards
    for card in cards:
        assert card.height() >= card.minimumSizeHint().height()
        for label in card.findChildren(QLabel):
            assert label.height() >= label.fontMetrics().height()
    win.close()


def test_card_hover_is_scoped_to_the_whole_card(app):
    win = SidecarWindow(FakeBridge())
    win.home.refresh_texts()

    cards = win.home.list_widget.findChildren(QWidget, "formulaCard")
    assert cards, "home screen should render formula cards"

    for card in cards:
        sheet = card.styleSheet()
        # hover must be scoped to the card and never hit a bare QWidget rule
        assert "QWidget#formulaCard" in sheet
        assert "QWidget:hover" not in sheet
        assert "QWidget {\n" not in sheet and "QWidget {" not in sheet

        labels = card.findChildren(QLabel)
        # only two lines now: the headline (formula + one-line summary) and the
        # example formula — the description paragraph was removed on request
        assert len(labels) == 2
        for label in labels:
            # text lines must not paint their own background …
            assert "transparent" in label.styleSheet()
            # … and must let the mouse reach the card for :hover
            assert label.testAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
