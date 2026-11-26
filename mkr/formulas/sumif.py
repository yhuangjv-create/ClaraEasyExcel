"""SUMIF — sum cells meeting a single criterion."""
from __future__ import annotations

from mkr.formulas.base import (
    ArgStep,
    FormulaBase,
    Suggestion,
    _looks_like_range,
    _not_empty,
    register,
)


@register("SUMIF")
class SumIf(FormulaBase):
    name = "SUMIF"
    zh_name = "SUMIF 条件求和"
    en_name = "SUMIF — Sum by condition"
    zh_desc = "对满足指定条件的单元格求和。"
    en_desc = "Sums cells that meet a single criterion."
    zh_tips = [
        "条件可以是数字、文本（加引号）、或比较运算如 \">10\"",
        "条件区域和求和区域尺寸必须一致（或省略求和区域时，条件区域同时作为求和区域）",
        "条件里写文本要加双引号，如 \"苹果\"",
    ]
    en_tips = [
        "Criterion can be a number, quoted text, or an operator like \">10\"",
        "Range and sum_range must have identical dimensions, or omit sum_range to use range for both",
        "Text criteria need double quotes: \"Apple\"",
    ]
    zh_syntax = "SUMIF(条件区域, 条件, [求和区域])"
    en_syntax = "SUMIF(range, criteria, [sum_range])"
    zh_example = '=SUMIF(B:B, ">100", C:C)'
    en_example = '=SUMIF(B:B, ">100", C:C)'

    def args(self) -> list[ArgStep]:
        return [
            ArgStep(
                key="range",
                zh_label="① 条件区域",
                en_label="① Range to test",
                zh_hint="要判断条件的单元格区域",
                en_hint="Cells to evaluate against the criterion",
                default="B:B",
                kind="range",
                validator=_looks_like_range,
            ),
            ArgStep(
                key="criteria",
                zh_label="② 条件",
                en_label="② Criterion",
                zh_hint="数字直接写；文本加双引号；比较运算如 \">=100\" 或 \"=张三\"",
                en_hint="Plain number for numeric match; quoted text / operators like \">=100\" or \"=John\"",
                default='">=100"',
                kind="text",
                validator=_not_empty,
            ),
            ArgStep(
                key="sum_range",
                zh_label="③ 求和区域（可选）",
                en_label="③ Sum range (optional)",
                zh_hint="要相加的区域；留空 = 用条件区域同时求和",
                en_hint="Numbers to sum; empty means use the condition range itself",
                default="C:C",
                kind="range",
                validator=lambda v: None if not v.strip() else _looks_like_range(v),
            ),
        ]

    def build(self, **kwargs) -> str:
        r = kwargs["range"].strip()
        c = kwargs["criteria"].strip()
        s = kwargs.get("sum_range", "").strip()
        if s:
            return f"SUMIF({_auto_abs(r)}, {c}, {_auto_abs(s)})"
        return f"SUMIF({_auto_abs(r)}, {c})"

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        if error is None:
            return []
        if error == "#VALUE!":
            return [
                Suggestion(
                    zh_title="条件写法有问题",
                    en_title="Criterion format error",
                    zh_detail="文本条件必须加双引号，例如 \"苹果\"；比较运算符前后不能有多余空格。",
                    en_detail="Text criteria must be quoted: \"Apple\". No stray spaces around operators.",
                ),
            ]
        if error == "#NAME?":
            return [
                Suggestion(
                    zh_title="公式拼写错误",
                    en_title="Formula name misspelled",
                    zh_detail="检查 SUMIF 是否拼对，以及引号是否都是英文双引号。",
                    en_detail="Check SUMIF spelling and that all quotes are English double quotes.",
                ),
            ]
        return super().diagnose(formula, error)


def _auto_abs(range_str: str) -> str:
    """Same helper as VLOOKUP — anchor columns."""
    import re

    def _anchor(cell: str) -> str:
        if "!" in cell:
            sheet, ref = cell.rsplit("!", 1)
            return f"{sheet}!{_anchor(ref)}"
        m = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", cell)
        if not m:
            return cell
        _, letters, row_dollar, digits = m.groups()
        return f"${letters}{row_dollar}{digits}"

    range_str = range_str.strip()
    if not range_str:
        return range_str
    if ":" not in range_str:
        return _anchor(range_str)
    a, b = range_str.split(":", 1)
    return f"{_anchor(a)}:{_anchor(b)}"
