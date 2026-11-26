"""COUNTIF — count cells meeting a single criterion."""
from __future__ import annotations

from mkr.formulas.base import (
    ArgStep,
    FormulaBase,
    Suggestion,
    _looks_like_range,
    _not_empty,
    register,
)


@register("COUNTIF")
class CountIf(FormulaBase):
    name = "COUNTIF"
    zh_name = "COUNTIF 条件计数"
    en_name = "COUNTIF — Count by condition"
    zh_desc = "统计满足指定条件的单元格个数。"
    en_desc = "Counts cells that meet a single criterion."
    zh_tips = [
        "条件可以是数字、文本（加引号）、或比较运算如 \">0\"",
        "条件区域可以是单列 / 单行 / 矩形区域",
        "对空白单元格计数用 \"\"（两对双引号代表空）",
    ]
    en_tips = [
        "Criterion can be number, quoted text, or operator like \">0\"",
        "Range can be a single column/row or a rectangular block",
        "Count blanks with two double-quotes: \"\"",
    ]
    zh_syntax = "COUNTIF(计数区域, 条件)"
    en_syntax = "COUNTIF(range, criteria)"
    zh_example = '=COUNTIF(A:A, "已完成")'
    en_example = '=COUNTIF(A:A, "Done")'

    def args(self) -> list[ArgStep]:
        return [
            ArgStep(
                key="range",
                zh_label="① 计数区域",
                en_label="① Range to count",
                zh_hint="要统计的单元格区域",
                en_hint="Cells to be evaluated and counted",
                default="A:A",
                kind="range",
                validator=_looks_like_range,
            ),
            ArgStep(
                key="criteria",
                zh_label="② 条件",
                en_label="② Criterion",
                zh_hint='文本加双引号；比较如 ">0" 或 "<>0"；空白用 ""',
                en_hint='Quote text; operators like ">0" or "<>0"; "" for blanks',
                default='"已完成"',
                kind="text",
                validator=_not_empty,
            ),
        ]

    def build(self, **kwargs) -> str:
        r = kwargs["range"].strip()
        c = kwargs["criteria"].strip()
        return f"COUNTIF({_auto_abs(r)}, {c})"

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        if error is None:
            return []
        if error == "#VALUE!":
            return [
                Suggestion(
                    zh_title="条件写法有问题",
                    en_title="Criterion format error",
                    zh_detail="文本条件必须加英文双引号；不要混入中文引号。",
                    en_detail="Text criteria must use English double quotes, not full-width ones.",
                ),
            ]
        return super().diagnose(formula, error)


def _auto_abs(range_str: str) -> str:
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
