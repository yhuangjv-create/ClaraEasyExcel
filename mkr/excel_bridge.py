"""COM bridge between ClaraEasyExcel and Microsoft Excel.

CRITICAL DESIGN:
  * win32com COM proxies MUST be created AND used in the SAME Python thread.
  * The GUI thread only POSTS requests (via QMetaObject.invokeMethod).
    All COM calls run inside ExcelBridgeThread's own Qt event loop, which does
    its own pythoncom.CoInitialize() / CoUninitialize().
  * Excel's DisplayAlerts / ScreenUpdating / EnableEvents are disabled while
    we write formulas so background recalculation or open dialogs can never
    block the call and freeze the GUI.
  * GUI threads never touch the _InnerBridge or its COM proxies.
"""
from __future__ import annotations

from typing import Any

from PySide6.QtCore import Q_ARG, QMetaObject, QObject, Qt, QThread, Signal, Slot

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # non-Windows or pywin32 not installed
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]


# Excel error codes returned as integers from COM — map to human-readable strings
_XL_ERROR_CODES = {
    -2146826246: "#DIV/0!",
    -2146826245: "#N/A",
    -2146826273: "#NAME?",
    -2146826276: "#NULL!",
    -2146826288: "#NUM!",
    -2146826281: "#REF!",
    -2146826272: "#VALUE!",
}


class _InnerBridge:
    """All COM side effects live here.  Only touched from ExcelBridgeThread."""

    def __init__(self) -> None:
        self._excel: Any = None

    # ---------- connection ----------
    def connect(self) -> dict:
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}
        try:
            self._excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                self._excel = win32com.client.Dispatch("Excel.Application")
                self._excel.Visible = True
            except Exception as exc:
                self._excel = None
                return {"ok": False, "error": str(exc)}
        return {"ok": True, "data": "connected"}

    def disconnect(self) -> None:
        try:
            if self._excel is not None:
                self._excel.Quit()
        except Exception:
            pass
        self._excel = None

    def _touch(self) -> bool:
        if self._excel is None:
            return False
        try:
            _ = self._excel.Version
            return True
        except Exception:
            self._excel = None
            return False

    # ---------- protected execute ----------
    def _with_protection(self, fn):
        """Run `fn()` on the Excel COM proxy while disabling background work."""
        if not self._touch():
            return {"ok": False, "error": "not connected"}
        old_su = old_ea = old_da = None
        try:
            try:
                old_su = self._excel.ScreenUpdating
                old_ea = self._excel.EnableEvents
                old_da = self._excel.DisplayAlerts
                self._excel.ScreenUpdating = False
                self._excel.EnableEvents = False
                self._excel.DisplayAlerts = False
            except Exception:
                pass
            return fn()
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        finally:
            for name, val in (
                ("ScreenUpdating", old_su),
                ("EnableEvents", old_ea),
                ("DisplayAlerts", old_da),
            ):
                if val is not None:
                    try:
                        setattr(self._excel, name, val)
                    except Exception:
                        pass

    # ---------- reads ----------
    def get_active_cell_address(self) -> dict:
        return self._with_protection(
            lambda: {"ok": True, "data": self._excel.ActiveCell.Address(False, False)}
        )

    def get_active_range_address(self) -> dict:
        return self._with_protection(
            lambda: {"ok": True, "data": self._excel.Selection.Address(False, False)}
        )

    def get_cell_value(self, address: str) -> dict:
        def _work() -> dict:
            ws = self._excel.ActiveSheet
            val = ws.Range(address).Value
            if isinstance(val, (int, float)) and val in _XL_ERROR_CODES:
                return {"ok": True, "data": _XL_ERROR_CODES[val], "is_error": True}
            return {"ok": True, "data": val, "is_error": False}

        return self._with_protection(_work)

    def get_cell_error(self, address: str) -> str | None:
        result = self.get_cell_value(address)
        if not result.get("ok"):
            return None
        if result.get("is_error"):
            return str(result["data"])
        return None

    # ---------- writes ----------
    def set_formula(self, address: str, formula: str) -> dict:
        if not formula.startswith("="):
            formula = "=" + formula

        def _work() -> dict:
            ws = self._excel.ActiveSheet
            ws.Range(address).Formula = formula
            return {"ok": True, "data": address}

        return self._with_protection(_work)

    def set_value(self, address: str, value: Any) -> dict:
        def _work() -> dict:
            ws = self._excel.ActiveSheet
            ws.Range(address).Value = value
            return {"ok": True, "data": address}

        return self._with_protection(_work)


# ============================================================================
# Thread wrapper — owns the COM proxy and runs all COM calls in its own
# thread.  GUI posts work via QMetaObject.invokeMethod (QueuedConnection),
# and picks up answers via signals.
# ============================================================================
class ExcelBridgeThread(QThread):
    """One and only place that touches pythoncom / win32com."""

    # emitted whenever the GUI requests a connection (or reconnection)
    connected = Signal(dict)
    active_cell = Signal(str)
    active_range = Signal(str)
    cell_error = Signal(object)  # str | None
    formula_written = Signal(dict)  # {"ok", "data"/"error", "addr"}

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._inner = _InnerBridge()

    # ---------- QThread event loop ----------
    def run(self) -> None:  # noqa: D401
        if pythoncom is not None:
            pythoncom.CoInitialize()
        self.exec()  # Qt event loop — slots invoked here run on this thread
        if pythoncom is not None:
            pythoncom.CoUninitialize()

    def shutdown(self) -> None:
        """Graceful stop — call this from Window's close handler."""
        try:
            self.invokeMethod(self, "_do_disconnect", Qt.ConnectionType.DirectConnection)
        except Exception:
            pass
        self.quit()
        self.wait(3000)

    # ---------- sync snapshot (always safe to call from GUI) ----------
    def is_connected(self) -> bool:
        """Best-effort — returns True iff the thread is up AND we still hold an
        Excel COM proxy.  May lag by one user action, never deadlocks."""
        return self.isRunning() and self._inner._excel is not None

    # ---------- GUI-facing request methods (called from GUI thread) ----------
    def request_connect(self) -> None:
        self.start()  # idempotent — safe if already running
        QMetaObject.invokeMethod(self, "_do_connect", Qt.ConnectionType.QueuedConnection)

    def request_active_cell(self) -> None:
        QMetaObject.invokeMethod(self, "_do_active_cell", Qt.ConnectionType.QueuedConnection)

    def request_active_range(self) -> None:
        QMetaObject.invokeMethod(self, "_do_active_range", Qt.ConnectionType.QueuedConnection)

    def request_set_formula(self, address: str, formula_body: str) -> None:
        QMetaObject.invokeMethod(
            self, "_do_set_formula",
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(str, address),
            Q_ARG(str, formula_body),
        )

    def request_cell_error(self, address: str) -> None:
        QMetaObject.invokeMethod(
            self, "_do_cell_error",
            Qt.ConnectionType.QueuedConnection,
            Q_ARG(str, address),
        )

    # ---------- Slot implementations — ALWAYS run on worker thread ----------
    @Slot()
    def _do_connect(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        result = self._inner.connect()
        self.connected.emit(result)

    @Slot()
    def _do_disconnect(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        self._inner.disconnect()

    @Slot()
    def _do_active_cell(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        result = self._inner.get_active_cell_address()
        self.active_cell.emit(str(result.get("data", "")) if result.get("ok") else "")

    @Slot()
    def _do_active_range(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        result = self._inner.get_active_range_address()
        self.active_range.emit(str(result.get("data", "")) if result.get("ok") else "")

    @Slot(str, str)
    def _do_set_formula(self, address: str, formula_body: str) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        result = self._inner.set_formula(address, formula_body)
        self.formula_written.emit({**result, "addr": address})

    @Slot(str)
    def _do_cell_error(self, address: str) -> None:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        self.cell_error.emit(self._inner.get_cell_error(address))
