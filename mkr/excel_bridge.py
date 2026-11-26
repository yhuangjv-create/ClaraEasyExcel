"""COM bridge — runs ALL COM calls on the CALLER's thread.

Design principles:
  * COM proxies are THREAD-AFFINED to the apartment that created them.
    We create and use them on the GUI thread; that's perfectly legal and
    avoids all QThread / invokeMethod / apartment-marshaling bugs.
  * We guard EVERY Excel call with DisplayAlerts=False / ScreenUpdating=False
    / EnableEvents=False so Excel background work (recalculation, dialogs,
    VBA macros) can never block our write.
  * pythoncom.CoInitialize() is called ONCE at app start, CoUninitialize()
    ONCE at app shutdown — balanced pair on the same thread.
"""
from __future__ import annotations

from typing import Any

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # non-Windows or pywin32 not installed
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]


_XL_ERROR_CODES = {
    # Mapped against ACTUAL Excel COM return values (tested on Excel 365 dynamic dispatch).
    # Codes come directly from win32com.client Range.Value when the cell contains an error.
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def _a1(cell) -> str:
    """Convert a COM Range object to A1-style address WITHOUT dollar signs.

    Why not use Range.Address(False, False)?
      On newer Excel (365+) dynamic-dispatch returns `.Address` as a **string**
      property that already equals "$A$1" — calling it like `.Address(False, False)`
      triggers `TypeError: 'str' object is not callable`.
      We construct the address deterministically from Row + Column and therefore
      never depend on COM's Address method signature.
    """
    row = int(cell.Row)
    col = int(cell.Column)
    letters = ""
    n = col
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


class ExcelComBridge:
    """Single-threaded COM wrapper.  Safe to create / use from GUI thread."""

    def __init__(self) -> None:
        self._excel: Any = None
        # CoInitialize happens once on first connect, balanced by CoUninitialize on shutdown
        self._co_inited = False

    # ---------- init / teardown ----------
    def co_init(self) -> None:
        """Call ONCE at app start (GUI thread).  Safe to call multiple times —
        COM maintains a refcount; calling again is a no-op on the same thread."""
        if pythoncom is not None and not self._co_inited:
            pythoncom.CoInitialize()
            self._co_inited = True

    def co_uninit(self) -> None:
        if pythoncom is not None and self._co_inited:
            self.disconnect()
            pythoncom.CoUninitialize()
            self._co_inited = False

    # ---------- connection ----------
    def connect(self) -> dict:
        self.co_init()
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}
        try:
            self._excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            try:
                self._excel = win32com.client.Dispatch("Excel.Application")
                self._excel.Visible = True
                self._excel.Workbooks.Add()  # ensure ActiveSheet is never None
            except Exception as exc:
                self._excel = None
                return {"ok": False, "error": str(exc)}

        # Safety: even after GetActiveObject, user might have closed all workbooks
        if self._excel.ActiveSheet is None:
            try:
                self._excel.Workbooks.Add()
            except Exception:
                pass
        return {"ok": True, "data": "connected"}

    def disconnect(self) -> None:
        try:
            if self._excel is not None:
                self._excel.Quit()
        except Exception:
            pass
        self._excel = None

    def is_connected(self) -> bool:
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
        """Run `fn()` with Excel's screen/events/alerts all disabled so nothing
        on the Excel side can block the call or pop a dialog."""
        if not self.is_connected():
            return {"ok": False, "error": "Excel not connected"}
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
            lambda: {"ok": True, "data": _a1(self._excel.ActiveCell)}
        )

    def get_active_range_address(self) -> dict:
        def _work() -> dict:
            sel = self._excel.Selection
            if sel.Rows.Count == 1 and sel.Columns.Count == 1:
                return {"ok": True, "data": _a1(sel)}
            top_left = _a1(sel.Cells(1, 1))
            bottom_right = _a1(sel.Cells(sel.Rows.Count, sel.Columns.Count))
            return {"ok": True, "data": f"{top_left}:{bottom_right}"}

        return self._with_protection(_work)

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
