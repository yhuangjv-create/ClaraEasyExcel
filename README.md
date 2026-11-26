# ClaraEasyExcel

> **Version 1.0.0** — [See CHANGELOG](#changelog)
>
> A lightweight **Excel sidecar assistant** — a floating, always-on-top window
> that guides novice Excel users through writing common formulas and (planned) VBA
> macros, with built-in intent detection, parameter wizards, and error diagnosis.
>
> 一款轻量 **Excel 外挂助手** — 悬浮置顶小窗口，引导 Excel 新手编写常用公式
> 和（规划中）VBA 宏，内置意图识别、参数向导、错误诊断。

---

## Why a sidecar, not an Excel add-in? / 为什么选外挂而非 Excel 插件？

|                       | Sidecar (this project) | Excel add-in (VSTO / Office JS) |
|-----------------------|:----------------------:|:-------------------------------:|
| Dev environment       | Python + PySide6 + pywin32 | VS + Office Developer Tools / Node + Yeoman |
| Deployment            | One PyInstaller EXE | ClickOnce / MSI + trust-center hoops |
| GUI freedom           | Full Qt | WPF only (VSTO) / HTML sandbox (Office JS) |
| Excel cross-version   | COM stable 2007 → 365 | Office JS spec shifts between versions |
| External / LLM calls  | Zero friction | Office JS sandbox restrictions |

For the MVP this matters because we want to ship to non-technical users **and** leave
a clean extension point for LLM integration later.

---

## Features / 功能

### Formula engine (v1.0.0, 5 formulas)

Each formula ships with a complete "guide-write-diagnose" pipeline:

| Formula | Intent keywords (examples) | Wizard steps | Diagnostics |
|---------|----------------------------|--------------|-------------|
| **VLOOKUP** | 匹配、查找、对应、vlookup | 4 steps (lookup / table / col / match) | #N/A、#REF!、#VALUE! → spelling / wildcard / IFERROR wrap |
| **SUMIF** | 求和、汇总、条件求和 | 3 steps (range / criteria / sum_range) | #VALUE! / #NAME? → quoting / spelling checks |
| **COUNTIF** | 计数、个数、次数 | 2 steps (range / criteria) | #VALUE! → quoting hints |
| **IFERROR** | 容错、错误避免、#N/A | 2 steps (inner formula / fallback) | #NAME? → inner formula check first |
| **CONCAT / CONCATENATE** | 拼接、合并文本 | 1 step (comma-separated items) | #NAME? → CONCATENATE fallback (older Excel) |

Every formula also carries syntax definition, usage examples, usage tips, and
auto-absolute (`$`) range anchoring.

### UI

- **Frameless floating window** (draggable, always-on-top, minimize/close)
- **Home screen** — keyword search → ranked formula cards
- **Wizard screen** — step-by-step parameter input with live hints, warnings,
  "pick from Excel" button (reads active cell / selection via COM)
- **Preview screen** — live formula preview, one-click write to active Excel
  cell, clipboard copy, error auto-diagnosis with "一键套用" fix buttons
- **EN / ZH language switch** in the title bar; persistence across sessions

---

## Architecture / 架构

```
┌─────────────────────────────────────────────────────────────────┐
│                          entry.py                               │
└──────────────────────────┬──────────────────────────────────────┘
                           ▼
                  ┌─────────────────┐
                  │  mkr / app.py   │  launch()
                  └────────┬────────┘
                           ▼
┌──────────────────────────────────────────────────────────────────┐
│                         mkr / gui.py                            │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐                        │
│  │ Home    │  │ Wizard  │  │ Preview │   QStackedWidget         │
│  └────┬────┘  └────┬────┘  └────┬────┘                        │
└───────┼────────────┼────────────┼───────────────────────────────┘
        │            │            │
        ▼            ▼            ▼
┌──────────────────────────────────────┐   ┌──────────────────────┐
│ mkr / intent.py  关键词 → 公式推荐 │──▶│ mkr / formulas /     │
│                                      │   base.py + vlookup, ... │
│                                      │   registry @register()  │
└──────────────────────────────────────┘   └───────────┬──────────┘
                                                       │
                                                       ▼
                                     ┌─────────────────────────────┐
                                     │ mkr / excel_bridge.py        │
                                     │ pywin32 · Excel.Application  │
                                     │ connect / get_active_cell /  │
                                     │ set_formula / get_cell_error │
                                     └──────────────┬──────────────┘
                                                    ▼
                                               Microsoft Excel
```

Supporting modules:

| Module | Role |
|--------|------|
| `mkr/i18n.py` | All zh/en strings — `tr(key, lang)` |
| `mkr/config.py` | Persistent config — language, window geometry, recent formulas, auto-connect flag. Portable on PyInstaller (`<exe>/config.json`), falls back to `%APPDATA%\ClaraEasyExcel` |

---

## Quick start / 快速开始

### Environment

- **Windows** only (pywin32 + Excel COM)
- Python ≥ 3.11 (tested on 3.15)
- Microsoft Excel installed (2019 / 2021 / 365 tested)

### Install

```powershell
cd ClaraEasyExcel
# create a venv (optional)
python -m venv .venv
.\.venv\Scripts\activate

pip install -e .          # runtime deps: PySide6 + pywin32
pip install -e ".[dev]"   # + pytest + ruff
```

### Run from source

```powershell
python entry.py
# or with pythonw (no console window):
pythonw entry.py
```

### Package to single-file EXE

```powershell
pyinstaller --onefile --windowed --name ClaraEasyExcel entry.py
# output: dist\ClaraEasyExcel.exe
```

### Dev checks

```powershell
ruff check .
pytest -v
```

---

## Roadmap

- **v1.1.0** — Add XLOOKUP, INDEX+MATCH, IFS, SUMIFS, COUNTIFS, IF, LEFT/RIGHT/MID, FIND, SUBTOTAL; VBA macro template engine with a 10-macro starter pack
- **v2.0.0** — Swap keyword-based intent → LLM-driven intent; support nested formula generation and natural-language VBA; optional Ollama local backend

---

## License

MIT — see [LICENSE](./LICENSE)

---

<a id="changelog"></a>
## CHANGELOG

### v1.0.0 — 2026-11-26

**First MVP release.**

- 5 core formulas implemented: VLOOKUP, SUMIF, COUNTIF, IFERROR, CONCAT (with CONCATENATE alias)
- Each formula ships with intent keyword detection, 1~4 step parameter wizard, auto-absolute `$` anchoring, syntax definition, usage tips, and per-error-type diagnostics
- PySide6 floating sidecar window: frameless, draggable, always-on-top, blue-only palette
- Three-screen flow: Home → Wizard → Preview
- COM bridge (`mkr/excel_bridge.py`) — connect/disconnect/is_connected/get_active_cell/set_formula/get_cell_error
- EN / ZH language toggle with session persistence; language, window geometry, recent formulas persisted to `config.json` (portable on EXE, fallback to `%APPDATA%`)
- 14 unit tests for intent detection and formula build pipeline; ruff enforced

#### v1.0.0 首次发布

- 实现 5 个核心公式：VLOOKUP、SUMIF、COUNTIF、IFERROR、CONCAT（兼容 CONCATENATE）
- 每个公式都包含意图关键词识别、1–4 步参数向导、自动绝对引用（$）、语法定义、使用技巧、分错误类型的诊断建议
- PySide6 悬浮外挂窗口：无边框、可拖动、置顶、仅使用蓝色调色板
- 三屏流程：首页 → 向导 → 预览
- COM 桥接层 (`mkr/excel_bridge.py`) — connect/disconnect/is_connected/get_active_cell/set_formula/get_cell_error
- 中英文切换 + 会话持久化；语言、窗口位置、最近使用公式保存到 `config.json`（EXE 同目录便携模式，不可写降级到 `%APPDATA%`）
- 14 个单元测试覆盖意图识别与公式生成管道；ruff 代码风格检查强制执行
