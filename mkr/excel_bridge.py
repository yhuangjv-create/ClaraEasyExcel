"""COM bridge — single-threaded on GUI thread, single proxy per session.
Connection strategy (deterministic, no guesswork):
  * We bind to the **visible Excel window** the user is actually working in,
    by enumerating top-level XLMAIN windows and asking MSAA for the native
    object behind its EXCEL7 sheet window (OBJID_NATIVEOM).  Excel started
    from Explorer / the Start menu usually does NOT appear in the COM ROT, so
    GetActiveObject alone cannot find it.
  * If several Excel windows exist we prefer, in order: the foreground one,
    the one we are already pinned to, any window with a real workbook file
    path, and only then a blank “工作簿1”.
  * GetActiveObject is a last-resort fallback and is trusted **only** when its
    `Application.Hwnd` is one of the visible windows.  A stale hidden instance
    (left over from older builds that used Dispatch) would otherwise hijack
    every write: its ActiveCell is A1, which is exactly the “it says it wrote
    to A1 but nothing appeared in my file” bug.
  * We never Dispatch while any Excel window exists.  Only when Excel is not
    running at all do we start a visible instance and mark _we_started_it, so
    the workbook we create is the one we write into and the one we close.

_rebind() policy (called before every read/write):
  * Keep the proxy pinned to the same Excel window, so the ActiveCell read and
    the following write always hit the SAME workbook.
  * Re-attach only when that window is gone, or the user switched to another
    Excel instance (foreground is a different Excel window).

CoInitialize / CoUninitialize:
  * Called ONCE at app start (co_init) and ONCE at app shutdown (co_uninit),
    both on the GUI thread.  Only co_uninit calls Quit() if _we_started_it.
"""
from __future__ import annotations

import ctypes
import time
from typing import Any

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32com.client.dynamic  # type: ignore[import-not-found]
except ImportError:
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]


class _GUID(ctypes.Structure):
    """Win32 GUID layout, used to pass IID_IDispatch to oleacc."""

    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


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
        self._hwnd: int | None = None  # main window of the Excel we are bound to
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

    # ---------- attaching to the Excel the user is actually using ----------
    @staticmethod
    def _excel_hwnds() -> list[int]:
        """Top-level windows of class XLMAIN — one per running Excel."""
        try:
            import win32gui
        except Exception:
            return []
        found: list[int] = []

        def _cb(hwnd: int, _param: Any) -> bool:
            try:
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "XLMAIN":
                    found.append(int(hwnd))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            return []
        return found

    @staticmethod
    def _foreground_excel(windows: list[int]) -> int | None:
        try:
            import win32gui

            root = win32gui.GetAncestor(win32gui.GetForegroundWindow(), 2)  # GA_ROOT
            if root in windows:
                return int(root)
        except Exception:
            pass
        return None

    @staticmethod
    def _find_child_class(root: int, cls: str, max_depth: int = 5) -> int | None:
        """Depth-first search for a child window of the given class."""
        try:
            import win32gui
        except Exception:
            return None
        pending: list[tuple[int, int]] = []
        try:
            child = win32gui.GetWindow(root, 5)  # GW_CHILD
            while child:
                pending.append((int(child), 1))
                child = win32gui.GetWindow(child, 2)  # GW_HWNDNEXT
        except Exception:
            return None
        while pending:
            hwnd, depth = pending.pop(0)
            try:
                if win32gui.GetClassName(hwnd) == cls:
                    return hwnd
                if depth < max_depth:
                    kid = win32gui.GetWindow(hwnd, 5)
                    while kid:
                        pending.append((int(kid), depth + 1))
                        kid = win32gui.GetWindow(kid, 2)
            except Exception:
                continue
        return None

    def _attach_window(self, hwnd: int) -> Any:
        """Application of a running Excel, obtained from its main window.

        Works even when Excel never registered in the COM ROT (the usual case
        when Excel is started from Explorer / the Start menu): we ask MSAA for
        the native object behind the sheet window (OBJID_NATIVEOM) and turn
        that raw interface pointer into a usable COM proxy.
        """
        try:
            import ctypes
            import uuid

            child = self._find_child_class(hwnd, "EXCEL7")
            if not child:
                return None
            iid = _GUID.from_buffer_copy(uuid.UUID(str(pythoncom.IID_IDispatch)).bytes_le)
            ptr = ctypes.POINTER(ctypes.c_void_p)()
            ctypes.oledll.oleacc.AccessibleObjectFromWindow(
                ctypes.c_void_p(child),
                ctypes.c_ulong(0xFFFFFFF0),  # OBJID_NATIVEOM
                ctypes.byref(iid),
                ctypes.byref(ptr),
            )
            raw = ctypes.cast(ptr, ctypes.c_void_p).value
            if not raw:
                return None
            unk = pythoncom.ObjectFromAddress(raw, pythoncom.IID_IDispatch)
            return win32com.client.dynamic.Dispatch(unk).Application
        except Exception:
            return None

    def _try_attach(self) -> bool:
        """Bind to a real, open workbook — foreground instance first.

        Windows whose workbook has a path (a real file the user opened) are
        preferred over blank "工作簿1" leftovers from older versions.
        """
        windows = self._excel_hwnds()
        order: list[int] = []
        fg = self._foreground_excel(windows)
        if fg is not None:
            order.append(fg)
        if self._hwnd is not None and self._hwnd in windows:
            order.append(int(self._hwnd))
        order += [h for h in windows if h not in order]

        blanks: list[int] = []
        for hwnd in order:
            app = self._attach_window(hwnd)
            if app is None:
                continue
            try:
                wb = app.ActiveWorkbook
                if wb is None:
                    continue  # nothing to write into
                has_path = bool(str(wb.Path or ""))
            except Exception:
                continue
            if not has_path:
                blanks.append(hwnd)
                continue
            self._excel = app
            self._hwnd = hwnd
            self._we_started_it = False
            return True

        # Nothing but unsaved blank workbooks → use the first one.
        for hwnd in blanks:
            app = self._attach_window(hwnd)
            if app is None:
                continue
            try:
                if app.ActiveWorkbook is None:
                    continue
            except Exception:
                continue
            self._excel = app
            self._hwnd = hwnd
            self._we_started_it = False
            return True

        # Fallback: the instance registered itself in the COM ROT.  Only trust
        # it if its main window is one we can actually see — a stale hidden
        # instance (left over from old Dispatch code) otherwise hijacks writes.
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            hwnd = int(app.Hwnd)
            if hwnd in windows and app.ActiveWorkbook is not None:
                self._excel = app
                self._hwnd = hwnd
                self._we_started_it = False
                return True
        except Exception:
            pass
        return False

    def workbook_name(self) -> str:
        try:
            wb = self._excel.ActiveWorkbook
            return str(wb.Name) if wb is not None else ""
        except Exception:
            return ""

    # ---------- public connect ----------
    def connect(self) -> dict:
        self.co_init()
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}

        # 1) Attach to the Excel the user has open.  We must NEVER spawn a
        #    second instance while Excel is running: a fresh blank workbook
        #    has its ActiveCell in A1, which is exactly the “says it wrote to
        #    A1 but nothing appeared in my file” bug.
        for _ in range(8):
            if self._try_attach():
                return {"ok": True, "data": self.workbook_name()}
            time.sleep(1.0)

        # 2) Excel really is not running → start one so the user can work.
        if not self._excel_hwnds():
            try:
                app = win32com.client.Dispatch("Excel.Application")
                app.Visible = True
                app.Workbooks.Add()
                self._excel = app
                self._hwnd = int(app.Hwnd)
                self._we_started_it = True
                return {"ok": True, "data": self.workbook_name()}
            except Exception as exc:
                self._excel = None
                self._we_started_it = False
                return {"ok": False, "error": str(exc)}

        # 3) Excel is up but unreachable (modal dialog, protected view, …).
        #    Say so instead of silently writing into some other workbook.
        return {"ok": False, "error": "excel_busy"}

    def is_connected(self) -> bool:
        if self._excel is None:
            return False
        try:
            return self._excel.ActiveWorkbook is not None
        except Exception:
            self._excel = None
            self._hwnd = None
            self._we_started_it = False
            return False

    # ---------- _rebind: called before every COM access ----------
    def _rebind(self) -> bool:
        """Keep the proxy pinned to the Excel window the user is working in.

        - our window is gone (closed) or the user switched to another Excel
          instance → re-attach;
        - otherwise keep the existing proxy so a read (ActiveCell) and the
          following write always hit the SAME workbook.
        """
        if win32com is None or pythoncom is None:
            return bool(self._excel)
        windows = self._excel_hwnds()
        fg = self._foreground_excel(windows)
        if fg is not None and fg != self._hwnd:
            if not self._try_attach():
                return False
        elif self._hwnd is None or (windows and self._hwnd not in windows):
            if not self._try_attach():
                return False
        if self._excel is None:
            return False
        try:
            return self._excel.ActiveWorkbook is not None
        except Exception:
            self._excel = None
            self._hwnd = None
            return False

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
