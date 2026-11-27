"""Regression: the bottom-left status line breathes instead of hard-blinking.

v1.3.9 replaced a static status label with a slow opacity pulse
(1.0 → 0.45 → 1.0 over 2.6 s, InOutSine).  A QLabel has no opacity property,
so the pulse runs on a QGraphicsOpacityEffect attached to the label.
"""

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import QAbstractAnimation
from PySide6.QtWidgets import QApplication, QGraphicsOpacityEffect

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
