"""VLOOKUP — look up a value in a table's first column, return from another column."""
from __future__ import annotations

from mkr.formulas.base import (
    ArgStep,
    FormulaBase,
    Suggestion,
    _looks_like_range,
    _not_empty,
    register,
)


@register("VLOOKUP")
class VLookup(FormulaBase):
    name = "VLOOKUP"
    zh_name = "VLOOKUP 条件查找"
    en_name = "VLOOKUP — Lookup by condition"
    zh_desc = "在查找表第一列中搜索一个值，返回同一行指定列的内容。"
    en_desc = "Searches for a value in the first column of a table and returns a value from a specified column."
    zh_tips = [
        "查找表第一列必须包含要搜索的值",
        "返回列索引从 1 开始（第 1 列 = 查找列本身）",
        "常用精确匹配 (FALSE)，需要近似匹配时查找列必须先排序",
        "建议给查找表加 $ 绝对引用，防止公式下拉错位",
    ]
    en_tips = [
        "The first column of the table must contain the lookup key",
        "Column index is 1-based (1 = the leftmost column of the table)",
        "Use exact match (FALSE) most of the time; approximate match needs the lookup column sorted ascending",
        "Add $ to the table range so copying the formula down doesn't shift it",
    ]
    zh_syntax = "VLOOKUP(查找值, 查找表, 返回列索引, 匹配方式)"
    en_syntax = "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])"
    zh_example = "=VLOOKUP(A2, $D:$F, 3, FALSE)"
    en_example = "=VLOOKUP(A2, $D:$F, 3, FALSE)"

    def args(self) -> list[ArgStep]:
        return [
            ArgStep(
                key="lookup",
                zh_label="① 查找值",
                en_label="① Lookup value",
                zh_hint="要在查找表第一列中搜索的值（单元格引用或加引号的文本）",
                en_hint="Value to search for in the first column (cell reference or quoted text)",
                default="A2",
                kind="cell",
                validator=_not_empty,
            ),
            ArgStep(
                key="table",
                zh_label="② 查找表区域",
                en_label="② Table range",
                zh_hint="包含查找列 + 返回列的完整区域，第一列必须是搜索键列",
                en_hint="Full range containing both lookup column and return column; first column must be the key",
                default="$D:$F",
                kind="range",
                validator=_looks_like_range,
            ),
            ArgStep(
                key="col_index",
                zh_label="③ 返回列索引",
                en_label="③ Return column index",
                zh_hint="查找表内第几列返回（从 1 开始，不得超过区域列数）",
                en_hint="Which column inside the table to return (1-based; must not exceed total columns)",
                default="3",
                kind="number",
                validator=lambda v: (
                    None if (v.strip().isdigit() and int(v.strip()) >= 1)
                    else "请输入 ≥ 1 的整数 / Enter an integer ≥ 1"
                ),
            ),
            ArgStep(
                key="match_mode",
                zh_label="④ 匹配方式",
                en_label="④ Match mode",
                zh_hint="精确匹配 (FALSE/0) 最常用；近似匹配 (TRUE/1) 需要查找列排序",
                en_hint="Exact match (FALSE/0) is default; approximate (TRUE/1) requires sorted lookup column",
                default="FALSE",
                kind="choice",
                choices=[
                    ("精确匹配 (FALSE)", "Exact match (FALSE)"),
                    ("近似匹配 (TRUE)", "Approximate match (TRUE)"),
                ],
                validator=_not_empty,
            ),
        ]

    def build(self, **kwargs) -> str:
        lookup = kwargs["lookup"].strip()
        table = kwargs["table"].strip()
        col = kwargs["col_index"].strip()
        match = kwargs["match_mode"].strip().upper()
        # normalize choice text → TRUE / FALSE
        if "FALSE" in match or "EXACT" in match.upper():
            match = "FALSE"
        elif "TRUE" in match or "APPROX" in match.upper():
            match = "TRUE"
        # auto-absolute the table range (user sees this in wizard toggle too)
        # keep existing $ signs, just make sure columns are anchored
        table = _auto_absolute(table)
        return f"VLOOKUP({lookup}, {table}, {col}, {match})"

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        if error is None:
            return []
        if error == "#N/A":
            return [
                Suggestion(
                    zh_title="A2 在查找表第一列中找不到匹配值",
                    en_title="No match found in the first column of the table",
                    zh_detail="① 检查查找值拼写 / ② 查找表区域是否选错 / ③ 改用通配符模糊匹配",
                    en_detail="① Check lookup value spelling / ② Wrong table range / ③ Try wildcard fuzzy match",
                ),
                Suggestion(
                    zh_title="③ 用通配符模糊匹配",
                    en_title="③ Wildcard fuzzy match",
                    zh_detail='把查找值换成 "*"&A2&"*"（只适用于精确匹配场景）',
                    en_detail='Wrap lookup value with "*"&A2&"*"',
                    zh_fix=formula.replace("VLOOKUP(", "VLOOKUP(").replace(
                        "VLOOKUP(", 'VLOOKUP("*"&'
                    ).replace(", ", ", ", 1),  # placeholder — wizard uses this field
                    en_fix=None,
                ),
                Suggestion(
                    zh_title="用 IFERROR 套一层容错",
                    en_title="Wrap with IFERROR",
                    zh_detail="找不到时显示空白或友好提示",
                    en_detail="Show blank / friendly text when not found",
                    zh_fix=f"IFERROR({formula}, \"\")",
                    en_fix=f"IFERROR({formula}, \"\")",
                ),
            ]
        if error == "#REF!":
            return [
                Suggestion(
                    zh_title="返回列索引超范围",
                    en_title="Column index out of range",
                    zh_detail="查找表区域只有 3 列时，索引不能填 4。检查一下索引值。",
                    en_detail="Index 4 for a 3-column table is invalid. Lower the index number.",
                ),
            ]
        if error == "#VALUE!":
            return [
                Suggestion(
                    zh_title="参数类型不对",
                    en_title="Argument type mismatch",
                    zh_detail="返回列索引必须是数字，匹配方式必须是 TRUE/FALSE 或 1/0。",
                    en_detail="Column index must be a number; match mode must be TRUE/FALSE or 1/0.",
                ),
            ]
        return super().diagnose(formula, error)


def _auto_absolute(range_str: str) -> str:
    """Add $ before column letters of the first segment in each half of A1:B3, keeping existing $."""
    range_str = range_str.strip()
    if not range_str or ":" not in range_str:
        return _anchor_one(range_str)
    a, b = range_str.split(":", 1)
    return f"{_anchor_one(a)}:{_anchor_one(b)}"


def _anchor_one(cell: str) -> str:
    """Anchor column letters in an A1-style cell ref; preserve existing $ and row anchors."""
    import re

    # strip sheet prefix for processing
    if "!" in cell:
        sheet, ref = cell.rsplit("!", 1)
        return f"{sheet}!{_anchor_one(ref)}"

    m = re.match(r"^(\$?)([A-Z]+)(\$?)(\d+)$", cell)
    if not m:
        # Whole-column reference picked by dragging the headers, e.g. "B"
        # from "B:D" — anchor the column so it never shifts when copied.
        m_col = re.match(r"^(\$?)([A-Z]+)$", cell)
        if m_col:
            return f"${m_col.group(2)}"
        return cell  # leave as-is if it's not A1 style (e.g. a text literal)
    col_dollar, letters, row_dollar, digits = m.groups()
    col_dollar = "$"  # always anchor column
    return f"{col_dollar}{letters}{row_dollar}{digits}"
