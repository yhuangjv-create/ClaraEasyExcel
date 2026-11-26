"""Formula engine — base class + registry.

Each concrete formula inherits FormulaBase and implements 4 things:
  1. Static metadata (names, description, tips, syntax, example)
  2. args()       — ordered wizard steps
  3. build(...)   — turn filled args into a formula string (without leading '=')
  4. diagnose(...) — given an Excel error string, return fix suggestions
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable


# ---------- wizard step definition ----------
@dataclass
class ArgStep:
    """One parameter step in the formula wizard."""
    key: str
    zh_label: str
    en_label: str
    zh_hint: str = ""
    en_hint: str = ""
    zh_warning: str = ""
    en_warning: str = ""
    default: str = ""
    # "cell" | "range" | "text" | "number" | "choice"
    kind: str = "cell"
    # for kind="choice"
    choices: list[tuple[str, str]] = field(default_factory=list)  # (zh, en)
    # returns None if ok, or error string
    validator: Callable[[str], str | None] = lambda v: None


# ---------- suggestion from diagnose ----------
@dataclass
class Suggestion:
    zh_title: str
    en_title: str
    zh_detail: str = ""
    en_detail: str = ""
    # optional ready-to-apply replacement formula
    zh_fix: str | None = None
    en_fix: str | None = None


# ---------- base class ----------
class FormulaBase:
    # metadata (override in subclass)
    name: str = ""  # canonical formula name, e.g. "VLOOKUP"
    zh_name: str = ""
    en_name: str = ""
    zh_desc: str = ""
    en_desc: str = ""
    zh_tips: list[str] = []
    en_tips: list[str] = []
    zh_syntax: str = ""
    en_syntax: str = ""
    zh_example: str = ""
    en_example: str = ""

    def args(self) -> list[ArgStep]:
        raise NotImplementedError

    def build(self, **kwargs) -> str:
        """Return formula body WITHOUT leading '='."""
        raise NotImplementedError

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        """Return fix suggestions given the formula string and Excel error value."""
        _ = formula
        if error is None:
            return []
        return [
            Suggestion(
                zh_title=f"Excel 返回了 {error}",
                en_title=f"Excel returned {error}",
                zh_detail="请检查参数是否合理，或用 IFERROR 套一层容错。",
                en_detail="Check parameters or wrap with IFERROR.",
            )
        ]


# ---------- simple validators shared by multiple formulas ----------
def _not_empty(v: str) -> str | None:
    return None if v.strip() else "请输入值 / Please provide a value"


def _looks_like_range(v: str) -> str | None:
    """Loose check: either a cell 'A1' or range 'A1:B3' (optional $ signs and sheet prefix)."""
    v = v.strip().replace("$", "")
    if not v:
        return "请输入单元格或区域 / Provide a cell or range"
    # Allow single cell or range
    if ":" in v:
        parts = v.split(":")
        if len(parts) != 2:
            return "区域格式应为 A1:B3 / Range format should be A1:B3"
    return None


# ---------- registry ----------
_REGISTRY: dict[str, FormulaBase] = {}


def register(name: str):
    """Class decorator — registers a FormulaBase subclass under `name`.

    Usage:
        @register("VLOOKUP")
        class VLookup(FormulaBase): ...

    Also supports multi-registering with aliases:
        @register("CONCAT")
        @register("CONCATENATE")
        class Concatenate(FormulaBase): ...
    """

    def _decorator(cls: type) -> type:
        # only instantiate once and re-use for each alias
        if not isinstance(cls, FormulaBase):
            # instantiate the subclass (it has no __init__ override)
            instance = cls()  # type: ignore[call-arg]
        else:
            instance = cls  # already an instance
        _REGISTRY[name] = instance
        return cls

    return _decorator


def get_formula(name: str) -> FormulaBase | None:
    return _REGISTRY.get(name)


def all_formula_names() -> list[str]:
    return list(_REGISTRY.keys())
