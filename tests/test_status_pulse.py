"""Regression: the bottom-left status line breathes and is painted with a gradient.

v1.3.9 replaced a static status label with a slow opacity pulse
(1.0 → 0.45 → 1.0 over 2.6 s, InOutSine).  A QLabel has no opacity property,
so the pulse runs on a QGraphicsOpacityEffect attached to the label.
v1.3.10 swapped the flat colour for a gold → white → ice-blue gradient, drawn
by `StatusLabel.paintEvent`.
"""

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import QAbstractAnimation
from PySide6.QtWidgets import QApplication, QGraphicsOpacityEffect

from mkr.gui import SidecarWindow, StatusLabel, _status_gradient


class FakeBridge:
    def is_connected(self) -> bool:
        return False

    def workbook_name(self) -> str:
        return ""


@pytest.fixture(scope="module")
def app():
    instance = QApplication.instance()
    return instance or QApplication([])


def test_status_label_pulses(app):
    win = SidecarWindow(FakeBridge())
    fx = win.status.graphicsEffect()
    assert isinstance(fx, QGraphicsOpacityEffect)

    anim = win._status_anim
    assert anim.state() == QAbstractAnimation.State.Running
    assert anim.loopCount() == -1
    assert anim.duration() == 2600
    assert fx.opacity() <= 1.0
    win.close()


def test_status_text_is_gradient_filled(app):
    win = SidecarWindow(FakeBridge())
    assert isinstance(win.status, StatusLabel)

    stops = _status_gradient(0, 100).stops()
    assert len(stops) == 3
    # three distinct colours: gold → white → blue on the dark theme
    assert len({color.name() for _, color in stops}) == 3
    # no green cast: the warm end leans yellow, the cool end leans blue
    warm, _, cool = (color for _, color in stops)
    assert warm.green() > warm.blue()
    assert cool.blue() > cool.green()
    win.close()
