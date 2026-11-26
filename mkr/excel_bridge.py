"""COM bridge between our app and Excel.

All win32com calls live here so GUI code never imports pywin32 directly.
Every public method catches COM errors and returns a dict {ok, data, error}
so the GUI thread stays alive no matter what Excel does.
"""
from __future__ import annotations

from typing import Any

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # non-Windows or pywin32 not installed
    win32com = None  # type: ignore[assignment]
    pythoncom = None  # type: ignore[assignment]


# Excel error values returned as string by the COM interface
_XL_ERROR_CODES = {
    -2146826246: "#DIV/0!",
    -2146826245: "#N/A",
    -2146826273: "#NAME?",
    -2146826276: "#NULL!",
    -2146826288: "#NUM!",
    -2146826281: "#REF!",
    -2146826272: "#VALUE!",
}


class ExcelComBridge:
    """Thin wrapper over Excel.Application COM."""

    def __init__(self) -> None:
        self._excel: Any = None

    # ---------- connection ----------
    def connect(self) -> dict:
        if win32com is None:
            return {"ok": False, "error": "pywin32 not installed"}
        try:
            pythoncom.CoInitialize()
            # prefer attaching to a running Excel; otherwise start a new one visible
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
        self._excel = None
        if pythoncom is not None:
            pythoncom.CoUninitialize()

    def is_connected(self) -> bool:
        if self._excel is None:
            return False
        try:
            _ = self._excel.Version  # light health check
            return True
        except Exception:
            self._excel = None
            return False

    # ---------- read ----------
    def get_active_cell_address(self) -> dict:
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        try:
            cell = self._excel.ActiveCell
            return {"ok": True, "data": cell.Address(False, False)}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def get_active_range_address(self) -> dict:
        """Return selected range address like 'A1:B5'. Falls back to single cell."""
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        try:
            sel = self._excel.Selection
            return {"ok": True, "data": sel.Address(False, False)}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def get_cell_value(self, address: str) -> dict:
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        try:
            ws = self._excel.ActiveSheet
            val = ws.Range(address).Value
            if isinstance(val, (int, float)) and val in _XL_ERROR_CODES:
                return {"ok": True, "data": _XL_ERROR_CODES[val], "is_error": True}
            return {"ok": True, "data": val, "is_error": False}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def get_cell_error(self, address: str) -> str | None:
        """Return error string (#N/A etc.) or None if no error. Pure string, no dict."""
        result = self.get_cell_value(address)
        if not result.get("ok"):
            return None
        if result.get("is_error"):
            return str(result["data"])
        return None

    # ---------- write ----------
    def set_formula(self, address: str, formula: str) -> dict:
        """Write formula string to `address`. Formula should start with '='."""
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        if not formula.startswith("="):
            formula = "=" + formula
        try:
            ws = self._excel.ActiveSheet
            ws.Range(address).Formula = formula
            return {"ok": True, "data": address}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def set_value(self, address: str, value: Any) -> dict:
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        try:
            ws = self._excel.ActiveSheet
            ws.Range(address).Value = value
            return {"ok": True, "data": address}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    # ---------- navigation ----------
    def select_address(self, address: str) -> dict:
        if not self.is_connected():
            return {"ok": False, "error": "not connected"}
        try:
            ws = self._excel.ActiveSheet
            ws.Activate()
            ws.Range(address).Select()
            return {"ok": True, "data": address}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
