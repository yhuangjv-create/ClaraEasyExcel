"""IFERROR — return a fallback when the wrapped formula errors."""
from __future__ import annotations

from mkr.formulas.base import ArgStep, FormulaBase, Suggestion, _not_empty, register


@register("IFERROR")
class IfError(FormulaBase):
    name = "IFERROR"
    zh_name = "IFERROR 错误容错"
    en_name = "IFERROR — Error guard"
    zh_desc = "公式报错时（如 #N/A、#DIV/0）返回你指定的内容，让表格更干净。"
    en_desc = "Returns a value you specify when the wrapped formula errors (#N/A, #DIV/0, etc.)."
    zh_tips = [
        "只捕获前一个参数的错误；嵌套公式要整体放进第一个参数",
        "第二个参数可以是文本（加引号）、数字、空值 \"\"、或另一条公式",
        "写空值显示空白，比显示 #N/A 更干净",
    ]
    en_tips = [
        "Only catches errors from the first argument — wrap the whole inner formula there",
        "Second arg can be quoted text, a number, \"\" (blank), or another formula",
        "Returning blank is usually cleaner than showing #N/A",
    ]
    zh_syntax = "IFERROR(要保护的公式, 出错时返回值)"
    en_syntax = "IFERROR(value, value_if_error)"
    zh_example = '=IFERROR(VLOOKUP(A2, $D:$F, 3, FALSE), "")'
    en_example = '=IFERROR(VLOOKUP(A2, $D:$F, 3, FALSE), "")'

    def args(self) -> list[ArgStep]:
        return [
            ArgStep(
                key="inner",
                zh_label="① 要保护的公式",
                en_label="① Formula to protect",
                zh_hint="完整的公式（可以引用单元格），写在等号之后的部分即可",
                en_hint="The full formula body (use the part after '=' if you already have one)",
                default="VLOOKUP(A2, $D:$F, 3, FALSE)",
                kind="text",
                validator=_not_empty,
            ),
            ArgStep(
                key="fallback",
                zh_label="② 出错时返回",
                en_label="② Value if error",
                zh_hint='文本加双引号；显示空白用 ""；也可以填数字或另一条公式',
                en_hint='Quote text; use "" for blank; or a number / another formula',
                default='""',
                kind="text",
                validator=_not_empty,
            ),
        ]

    def build(self, **kwargs) -> str:
        inner = kwargs["inner"].strip()
        fallback = kwargs["fallback"].strip()
        # strip leading = from inner if user pasted one
        if inner.startswith("="):
            inner = inner[1:]
        return f"IFERROR({inner}, {fallback})"

    def diagnose(self, formula: str, error: str | None) -> list[Suggestion]:
        if error is None:
            return []
        if error == "#NAME?":
            return [
                Suggestion(
                    zh_title="IFERROR 没生效",
                    en_title="IFERROR not working",
                    zh_detail="内层公式本身有拼写错误，导致 IFERROR 也识别不了。请检查内部公式。",
                    en_detail="Inner formula has a typo — fix that first, IFERROR only catches runtime errors.",
                ),
            ]
        return super().diagnose(formula, error)
