"""Bilingual dictionary (zh / en).

All UI strings live here so language switching is a one-line change.
Missing keys fall back to the zh value, then to the key itself.
"""

_TRANSLATIONS: dict[str, dict[str, str]] = {
    # ---------- GUI chrome ----------
    "app_title": {
        "zh": "ClaraEasyExcel",
        "en": "ClaraEasyExcel",
    },
    "btn_lang_zh": {"zh": "中文", "en": "中文"},
    "btn_lang_en": {"zh": "English", "en": "English"},
    "btn_min": {"zh": "—", "en": "—"},
    "btn_close": {"zh": "×", "en": "×"},
    "btn_theme": {"zh": "主题", "en": "Theme"},
    "theme_purple": {"zh": "深紫色（默认）", "en": "Deep purple (default)"},
    "theme_auto": {"zh": "跟随 Windows 主题色", "en": "Follow Windows accent"},
    "theme_custom": {"zh": "自定义颜色…", "en": "Custom colour…"},
    "theme_pick_title": {"zh": "选择主界面底色", "en": "Pick the window base colour"},
    "theme_applied": {"zh": "已应用主题色", "en": "Theme applied"},
    "diag_excel_busy": {
        "zh": "Excel 已打开但暂时连不上，请先点一下 Excel 的单元格再重试",
        "en": "Excel is open but unreachable — click a cell in Excel and retry",
    },

    # ---------- Home screen ----------
    "home_prompt": {
        "zh": "你想做什么？（例如：匹配、求和、统计个数、容错、拼接）",
        "en": "What do you want? (e.g. lookup, sum condition, count, error guard, concat)",
    },
    "home_placeholder": {
        "zh": "输入关键词或描述",
        "en": "Type keyword or describe your need",
    },
    "home_recommend": {"zh": "推荐公式", "en": "Recommended formulas"},
    "home_all": {"zh": "全部公式", "en": "All formulas"},
    "home_recent": {"zh": "最近使用", "en": "Recent"},
    "home_no_match": {
        "zh": "没匹配到合适的公式，试试：匹配 / 求和 / 计数 / 容错 / 拼接",
        "en": "No match. Try: lookup, sum, count, error guard, concatenate",
    },

    # ---------- Formula card ----------
    "card_syntax": {"zh": "语法", "en": "Syntax"},
    "card_example": {"zh": "示例", "en": "Example"},
    "card_start": {"zh": "开始填写 →", "en": "Start →"},
    "card_tips": {"zh": "注意事项", "en": "Tips"},

    # ---------- Wizard ----------
    "wizard_step": {"zh": "第 {cur}/{tot} 步 · {title}", "en": "Step {cur}/{tot} · {title}"},
    "wizard_value": {"zh": "当前值", "en": "Value"},
    "wizard_pick_cell": {"zh": "🔗 选单元格", "en": "🔗 Pick cell"},
    "wizard_pick_range": {"zh": "🔗 框选区域", "en": "🔗 Pick range"},
    "wizard_fix_abs": {"zh": "🔒 自动加绝对引用 ($)", "en": "🔒 Auto absolute ($)"},
    "wizard_prev": {"zh": "上一步", "en": "Back"},
    "wizard_next": {"zh": "下一步 →", "en": "Next →"},
    "wizard_done": {"zh": "完成填写！预览如下：", "en": "Done! Preview below:"},
    "wizard_ok": {"zh": "✓ 公式状态：语法正确，无明显错误", "en": "✓ Status: syntax looks fine"},
    "wizard_write": {"zh": "写入活动单元格", "en": "Write to active cell"},
    "wizard_copy": {"zh": "复制公式", "en": "Copy formula"},
    "wizard_cancel": {"zh": "取消", "en": "Cancel"},
    "wizard_goto_home": {"zh": "返回首页", "en": "Home"},

    # ---------- Errors / diagnostics ----------
    "diag_title": {"zh": "⚠️ 公式错误", "en": "⚠️ Formula error"},
    "diag_analysis": {"zh": "原因分析", "en": "Analysis"},
    "diag_solutions": {"zh": "可能的原因 & 修正方法", "en": "Possible causes & fixes"},
    "diag_fix_action": {"zh": "一键套用", "en": "Apply fix"},
    "diag_more": {"zh": "了解更多", "en": "Learn more"},
    "diag_conn_ok": {"zh": "Excel 已连接", "en": "Excel connected"},
    "diag_conn_fail": {"zh": "Excel 未运行 — 请先打开 Excel 文件", "en": "Excel not running — please open Excel first"},

    # ---------- Generic ----------
    "ok": {"zh": "好的", "en": "OK"},
    "copied": {"zh": "已复制到剪贴板 ✓", "en": "Copied to clipboard ✓"},
    "written": {"zh": "已写入活动单元格 ✓", "en": "Written to active cell ✓"},
    "pick_mode_active": {
        "zh": "已进入取数模式（窗口已最小化）：可以先切换工作表，再点击或框选目标区域。"
              "按住左键拖动时不会误确认，松开左键后自动填入并把窗口弹回，无需回来点确认。",
        "en": "Pick mode on (window minimised): switch worksheets first if you need "
              "to, then click or drag-select the target. Nothing is captured while "
              "the mouse is held down — release it and the address fills in and the "
              "window pops back, no confirm click needed.",
    },
    "pick_no_change": {
        "zh": "未检测到新的选择，窗口已弹回。若要使用当前选中的单元格，可点「读取选中单元格」。",
        "en": "No new selection detected — the window is back. To use the currently "
              "selected cell, click “Read selected cell”.",
    },
    "pick_read_now": {"zh": "✓ 读取选中单元格", "en": "✓ Read selected cell"},
    "confirm": {"zh": "确认", "en": "Confirm"},
    "cancel": {"zh": "取消", "en": "Cancel"},

    # ---------- Formula names / hints ----------
    "fml_vlookup": {"zh": "VLOOKUP 条件查找", "en": "VLOOKUP — Lookup by condition"},
    "fml_vlookup_desc": {
        "zh": "在查找表第一列中搜一个值，返回同一行指定列的内容。",
        "en": "Searches for a value in the first column of a table and returns a value from a specified column.",
    },
    "fml_sumif": {"zh": "SUMIF 条件求和", "en": "SUMIF — Sum by condition"},
    "fml_sumif_desc": {
        "zh": "对满足指定条件的单元格求和。",
        "en": "Sums cells that meet a single criterion.",
    },
    "fml_countif": {"zh": "COUNTIF 条件计数", "en": "COUNTIF — Count by condition"},
    "fml_countif_desc": {
        "zh": "统计满足指定条件的单元格个数。",
        "en": "Counts cells that meet a single criterion.",
    },
    "fml_iferror": {"zh": "IFERROR 错误容错", "en": "IFERROR — Error guard"},
    "fml_iferror_desc": {
        "zh": "公式报错时（如 #N/A）返回你指定的内容，让表格更干净。",
        "en": "Returns a value you specify when the wrapped formula errors (#N/A, #DIV/0, etc.).",
    },
    "fml_concat": {"zh": "CONCAT 拼接文本", "en": "CONCAT — Concatenate text"},
    "fml_concat_desc": {
        "zh": "把多个文本或单元格内容拼成一个字符串。",
        "en": "Joins multiple text items / cell values into one string.",
    },
}


def tr(key: str, lang: str = "zh", **fmt) -> str:
    """Return translated string for `key` in `lang`. Falls back to zh then to the key."""
    entry = _TRANSLATIONS.get(key)
    if entry is None:
        return key.format(**fmt)
    text = entry.get(lang) or entry.get("zh") or key
    return text.format(**fmt)
