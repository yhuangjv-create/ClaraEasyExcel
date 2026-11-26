"""COM bridge — single-threaded on GUI thread, ALWAYS re-binds to the
user's currently-active Excel window before any read/write.

WHY WE RE-BIND BEFORE EVERY OPERATION (critical):
  Python's `win32com.client.GetActiveObject("Excel.Application")` fails
  (MK_E_UNAVAILABLE) for up to ~2 seconds after Excel starts, even though
  the process is already alive.  During that window our `connect()` falls
  through to `Dispatch()` which spawns a BRAND-NEW Excel instance, binds
  our proxy to it, and all subsequent writes land in THAT instance —
  which the user can't see.  Result: "write to active cell does nothing".
  Fix: **re-bind via GetActiveObject immediately before every COM call**.
  If it fails we know the user hasn't opened Excel yet and we error out
  cleanly — never silently write into a hidden Dispatch instance.

Threading:
  COM proxies are apartment-affined.  We create + use them on the GUI
  thread only — one CoInitialize at app start, one CoUninitialize at exit.
  See `_with_protection` for the DisplayAlerts/ScreenUpdating/EnableEvents
  shield that prevents Excel-side blocking.
"""
from __future__ import annotations

import time
from typing import Any

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # non-Windows or pywin32 not installed
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]


# Actual Excel 365 COM error codes — mapped by live `Range.Value` tests.
_XL_ERROR_CODES: dict[int, str] = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def _a1(cell) -> str:
    """Convert a COM Range to A1 address without dollar signs.

    We do NOT use Range.Address(False, False) because on current Excel
    (365) dynamic-dispatch proxies expose .Address as a STRING property
    (returns "$A$1" directly), not a callable.  Calling it throws
    TypeError.  Row + Column → A1 letters is deterministic across versions.
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
    """Sync COM wrapper.  Callers always run on the GUI thread."""

    def __init__(self) -> None:
        self._excel: Any = None
        self._co_inited = False
        # True when WE dispatched an Excel instance (not user-opened).
        # We will close OUR instances on quit; user-opened ones stay alive.
        self._we_started_it = False

    # ---------- apartment ----------
    def co_init(self) -> None:
        if pythoncom is not None and not self._co_inited:
            pythoncom.CoInitialize()
            self._co_inited = True

    def co_uninit(self) -> None:
        if pythoncom is not None and self._co_inited:
            # Only quit if *we* started the instance.  Never close user's Excel!
            if self._we_started_it and self._excel is not None:
                try:
                    self._excel.Quit()
                except Exception:
                    pass
            self._excel = None
            self._we_started_it = False
            pythoncom.CoUninitialize()
            self._co_inited = False

    # ---------- public connect ----------
    def connect(self) -> dict:
        """Try hard to bind to the USER's already-running Excel.

        Strategy:
          1. Try GetActiveObject up to 5 times with 0.5s gap — gives Excel
             time to register itself in the ROT after the user starts it.
          2. If still nothing, Dispatch a fresh Excel and mark
             _we_started_it=True so we know to quit it on shutdown.
          3. After any path, ensure ActiveSheet exists (Dispatch starts
             with no workbook = ActiveSheet is None).
        """
        self.co_init()
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}

        # Try to bind to user's Excel (retry because ROT registration lags)
        for _ in range(6):
            try:
                self._excel = win32com.client.GetActiveObject("Excel.Application")
                self._we_started_it = False
                break
            except Exception:
                time.sleep(1.0)

        if self._excel is None:
            try:
                # Dispatch fallback — WE own this instance, start hidden
                self._excel = win32com.client.Dispatch("Excel.Application")
                self._excel.Visible = False  # background only
                self._excel.Workbooks.Add()
                self._we_started_it = True
            except Exception as exc:
                self._excel = None
                self._we_started_it = False
                return {"ok": False, "error": str(exc)}

        # Safety: make sure ActiveSheet isn't None
        if self._excel.ActiveSheet is None:
            try:
                self._excel.Workbooks.Add()
            except Exception:
                pass

        return {"ok": True, "data": "connected"}

    def is_connected(self) -> bool:
        """Best-effort check — fast, non-blocking."""
        if self._excel is None:
            return False
        try:
            _ = self._excel.Version
            return True
        except Exception:
            self._excel = None
            self._we_started_it = False
            return False

    # ---------- private: ALWAYS re-bind before touching COM ----------
    def _rebind(self) -> bool:
        """Force _excel to point to the currently-active Excel window.

        This is called before EVERY read/write.  If the user opens a second
        Excel window, closes theirs, or our proxy went stale, we refresh.

        Returns True if we now have a usable proxy, False otherwise.
        """
        if win32com is None or pythoncom is None:
            return False
        try:
            self._excel = win32com.client.GetActiveObject("Excel.Application")
            self._we_started_it = False
        except Exception:
            # No running Excel — give up
            self._excel = None
            self._we_started_it = False
            return False
        # Make sure ActiveSheet is valid
        if self._excel.ActiveSheet is None:
            try:
                self._excel.Workbooks.Add()
            except Exception:
                return False
        return True

    # ---------- protected exec ----------
    def _with_protection(self, fn):
        """Run `fn()` on the active Excel instance with background work disabled."""
        if not self._rebind():
            return {"ok": False, "error": "Excel not connected. Please open Excel first."}

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
            tl = _a1(sel.Cells(1, 1))
            br = _a1(sel.Cells(sel.Rows.Count, sel.Columns.Count))
            return {"ok": True, "data": f"{tl}:{br}"}

        return self._with_protection(_work)

    def get_cell_value(self, address: str) -> dict:
        def _work() -> dict:
            ws = self._excel.ActiveSheet
            val = ws.Range(address).Value
            if isinstance(val, int) and val in _XL_ERROR_CODES:
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
