"""COM bridge — single-threaded on GUI thread, single proxy per session.

Connection strategy (deterministic, no guesswork):
  * **GetActiveObject("Excel.Application") is the ONLY way to find the user's
    already-running Excel.**  A Dispatch-created Excel always registers in
    the ROT immediately; a user-created Excel takes 2-15 seconds after the
    window first appears before COM's ROT knows about it.  We retry up to
    15 times × 1 second to wait for it.
  * If GetActiveObject still fails after 15 seconds, the user genuinely
    hasn't opened Excel yet — we Dispatch a **Visible=True** Excel (user can
    SEE it, so they know we created it) and mark _we_started_it=True so we
    close it on shutdown.  We NEVER create a hidden Dispatch instance — a
    hidden one is exactly what causes the "writes land nowhere" bug.
  * After Dispatch, we retry GetActiveObject **once more**.  Why?  Dispatch
    registers immediately, so now GetActiveObject WILL succeed (it returns
    whichever Excel registered last).  We use this new proxy so Dispatch
    can be garbage-collected.

_rebind() policy (called before every read/write):
  * Try GetActiveObject.  If it succeeds → swap proxy.  This means if the
    user JUST opened their Excel window and it finally registered, we pick
    it up next call.
  * If GetActiveObject fails → keep using whatever proxy we already have.
    No Dispatch fallback inside _rebind — that would double-spawn Excel.

CoInitialize / CoUninitialize:
  * Called ONCE at app start (co_init) and ONCE at app shutdown (co_uninit),
    both on the GUI thread.  Only co_uninit calls Quit() if _we_started_it.
"""
from __future__ import annotations

import time
from typing import Any

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except ImportError:
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
        self._we_started_it = False

    # ---------- apartment ----------
    def co_init(self) -> None:
        if pythoncom is not None and not self._co_inited:
            pythoncom.CoInitialize()
            self._co_inited = True

    def co_uninit(self) -> None:
        if pythoncom is not None and self._co_inited:
            # We created this Excel → safe to quit
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
        self.co_init()
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}

        # Phase 1: try GetActiveObject for up to 15 seconds
        for _ in range(15):
            try:
                self._excel = win32com.client.GetActiveObject("Excel.Application")
                self._we_started_it = False
                break
            except Exception:
                time.sleep(1.0)

        # Phase 2: if still nothing — create one VISIBLE (user sees it)
        _we_created_one = False
        if self._excel is None:
            try:
                dis = win32com.client.Dispatch("Excel.Application")
                dis.Visible = True
                dis.Workbooks.Add()
                _we_created_one = True
                self._we_started_it = True
                # Phase 3: retry GetActiveObject now that Dispatch registered
                # — GetActiveObject will find whichever registered last
                time.sleep(0.5)
                try:
                    self._excel = win32com.client.GetActiveObject("Excel.Application")
                    if not _we_created_one:
                        self._we_started_it = False
                    # else keep _we_started_it=True — Dispatch IS ours
                    dis = None  # orphan Dispatch proxy, GC it
                except Exception:
                    # Fallback — use Dispatch proxy directly
                    self._excel = dis
            except Exception as exc:
                self._excel = None
                self._we_started_it = False
                return {"ok": False, "error": str(exc)}

        # Safety: ActiveSheet can't be None
        if self._excel.ActiveSheet is None:
            try:
                self._excel.Workbooks.Add()
            except Exception:
                pass

        return {"ok": True, "data": "connected"}

    def is_connected(self) -> bool:
        if self._excel is None:
            return False
        try:
            _ = self._excel.Version
            return True
        except Exception:
            self._excel = None
            self._we_started_it = False
            return False

    # ---------- _rebind: called before every COM access ----------
    def _rebind(self) -> bool:
        """Try to swap proxy for the currently-active Excel.  Non-fatal if
        GetActiveObject fails — we keep using what we already have."""
        if win32com is None or pythoncom is None:
            return bool(self._excel)
        try:
            self._excel = win32com.client.GetActiveObject("Excel.Application")
            # Don't touch _we_started_it — it's set once at connect time and
            # stays sticky so we always know whether we own this instance.
        except Exception:
            pass
        if self._excel is None:
            return False
        # Ensure ActiveSheet exists
        if self._excel.ActiveSheet is None:
            try:
                self._excel.Workbooks.Add()
            except Exception:
                return False
        return True

    # ---------- protected exec ----------
    def _with_protection(self, fn):
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

    def read_selection_address(self, mode: str = "cell") -> str | None:
        """Silent address read used by the wizard's live pick polling.

        Deliberately skips _with_protection(): toggling ScreenUpdating /
        EnableEvents twice a second while the user drags a selection makes
        Excel flicker and stutter.  Returns None on any failure.
        """
        if not self._rebind():
            return None
        try:
            if mode == "range":
                sel = self._excel.Selection
                if sel is None:
                    return None
                rows = int(sel.Rows.Count)
                cols = int(sel.Columns.Count)
                if rows == 1 and cols == 1:
                    return _a1(sel)
                return f"{_a1(sel.Cells(1, 1))}:{_a1(sel.Cells(rows, cols))}"
            return _a1(self._excel.ActiveCell)
        except Exception:
            return None

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
