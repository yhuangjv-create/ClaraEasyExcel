"""CONCAT / CONCATENATE — join multiple text items."""
from __future__ import annotations

from mkr.formulas.base import ArgStep, FormulaBase, Suggestion, register


@register("CONCATENATE")   # alias → bottom-up: CONCAT applies first (canonical), CONCATENATE becomes alias
@register("CONCAT")        # canonical (modern Excel 2019+ name)
class Concatenate(FormulaBase):
    name = "CONCAT"
    zh_name = "CONCAT 拼接文本"
    en_name = "CONCAT — Concatenate text"
    zh_desc = "把多个文本或单元格内容拼成一个字符串（新版 Excel 推荐用 CONCAT）。"
    en_desc = "Joins multiple text items / cell values into one string (CONCAT is modern, CONCATENATE works too)."
    zh_tips = [
        "要插入分隔符（如空格、逗号），把分隔符也加引号作为一个参数",
        "Excel 2019 之前版本用 CONCATENATE；CONCATENATE 兼容所有版本",
        "也可用 \"&\" 运算符：=A2&\" \"&B2，和 CONCAT(A2,\" \",B2) 效果一样",
    ]
    en_tips = [
        "To insert a separator (space, comma), add it as a quoted extra argument",
        "Excel < 2019 uses CONCATENATE instead; CONCATENATE is universally compatible",
        "Equivalent with the & operator: =A2&\" \"&B2",
    ]
    zh_syntax = "CONCAT(项1, 项2, 项3, ...)"
    en_syntax = "CONCAT(text1, text2, text3, ...)"
    zh_example = '=CONCAT(A2, "-", B2)'
    en_example = '=CONCAT(A2, "-", B2)'

    def args(self) -> list[ArgStep]:
        return [
            ArgStep(
                key="items",
                zh_label="① 要拼接的各项",
                en_label="① Items to join",
                zh_hint='多个项用英文逗号分隔；单元格直接写；文本加双引号。例：A2, " ", B2',
                en_hint='Comma-separated; cells as-is; text in quotes. Example: A2, " ", B2',
                default='A2, " ", B2',
                kind="text",
                validator=lambda v: None if v.strip() else "请至少填写一个项 / Provide at least one item",
            ),
        ]

    def build(self, **kwargs) -> str:
        items = kwargs["items"].strip()
        return f"CONCAT({items})"

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        if error is None:
            return []
        if error == "#NAME?":
            return [
                Suggestion(
                    zh_title="旧版 Excel 不支持 CONCAT",
                    en_title="CONCAT not supported in older Excel",
                    zh_detail="把 CONCAT 改成 CONCATENATE 即可，完全兼容。",
                    en_detail="Replace CONCAT with CONCATENATE — fully compatible.",
                    zh_fix=formula.replace("CONCAT(", "CONCATENATE("),
                    en_fix=formula.replace("CONCAT(", "CONCATENATE("),
                ),
            ]
        if error == "#VALUE!":
            return [
                Suggestion(
                    zh_title="引号写多/写少了",
                    en_title="Quotes mismatch",
                    zh_detail="检查每个文本片段的双引号是否成对。不要混用中文引号。",
                    en_detail="Make sure every text literal has a matched pair of English double quotes.",
                ),
            ]
        return super().diagnose(formula, error)
