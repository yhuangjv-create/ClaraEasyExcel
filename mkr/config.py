"""Persistent config (language, window geometry, recent formulas, ...).

Storage location rules:
  - If the EXE directory is writable (PyInstaller frozen) → <exe-dir>/config.json (portable)
  - Otherwise → %APPDATA%/ClaraEasyExcel/config.json
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

_CONFIG_FILENAME = "config.json"
_APPNAME = "ClaraEasyExcel"

# Default window colour — the “deep purple” preset.  "auto" = Windows accent.
DEFAULT_THEME = "#5C2D91"


class Config:
    _instance: "Config | None" = None

    def __init__(self) -> None:
        self.lang: str = "zh"
        self.window_geom: tuple[int, int, int, int] | None = None  # x, y, w, h
        self.recent_formulas: list[str] = []  # formula names recently used
        self.auto_connect: bool = True
        self.theme: str = DEFAULT_THEME  # "#RRGGBB", or "auto" = Windows accent

        self._path = self._resolve_path()
        self.load()

    # ---------- singleton ----------
    @classmethod
    def get(cls) -> "Config":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    # ---------- path resolution ----------
    @staticmethod
    def _resolve_path() -> Path:
        if getattr(sys, "frozen", False):  # PyInstaller
            base = Path(sys.executable).parent
        else:
            base = Path(__file__).resolve().parent.parent

        target = base / _CONFIG_FILENAME
        try:
            # test writability
            test = target.with_suffix(".write_test")
            test.write_text("ok", encoding="utf-8")
            test.unlink()
            return target
        except OSError:
            pass

        # fallback to %APPDATA%
        appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        target = Path(appdata) / _APPNAME / _CONFIG_FILENAME
        target.parent.mkdir(parents=True, exist_ok=True)
        return target

    # ---------- load / save ----------
    def load(self) -> None:
        if not self._path.exists():
            return
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        self.lang = data.get("lang", "zh")
        self.auto_connect = bool(data.get("auto_connect", True))
        self.theme = str(data.get("theme", DEFAULT_THEME))
        self.recent_formulas = list(data.get("recent_formulas", []))
        geom = data.get("window_geom")
        if isinstance(geom, list) and len(geom) == 4:
            self.window_geom = tuple(geom)

    def save(self) -> None:
        data = {
            "lang": self.lang,
            "auto_connect": self.auto_connect,
            "theme": self.theme,
            "recent_formulas": self.recent_formulas[-10:],
            "window_geom": list(self.window_geom) if self.window_geom else None,
        }
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            self._path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass  # best effort — never crash the app

    # ---------- helpers ----------
    def add_recent(self, formula_name: str) -> None:
        if formula_name in self.recent_formulas:
            self.recent_formulas.remove(formula_name)
        self.recent_formulas.insert(0, formula_name)
        self.recent_formulas = self.recent_formulas[:10]
        self.save()
