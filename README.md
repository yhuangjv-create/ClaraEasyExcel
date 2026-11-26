# ClaraEasyExcel

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.2-blue)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Platform](https://img.shields.io/badge/platform-windows-blue)

**A lightweight Excel sidecar assistant — floating, always-on-top, no installation required.**

English · [中文](./README_CN.md) · [Changelog](#changelog)

</div>

---

ClaraEasyExcel is a small **float-sidecar window** that sits next to Microsoft Excel and helps
novice Excel users write common internal formulas and (planned) VBA macros. It offers:

- **Intent detection** — type "lookup" or "匹配" and ClaraEasyExcel recommends VLOOKUP
- **Step-by-step wizards** — no need to memorize parameter order or quotes
- **Built-in tips & warnings** — why the lookup table should use `$`, when to use FALSE vs TRUE
- **Live error diagnosis** — Excel returns `#N/A`? ClaraEasyExcel tells you why and how to fix it
- **One-click write** — formula lands in your active cell, perfect

We ship as a **pure sidecar** (not an Excel add-in) — one PyInstaller EXE, zero registry
changes, zero deployment hassle. See [Why a sidecar?](#why-a-sidecar-not-an-excel-add-in).

---

## Table of contents

- [Why a sidecar?](#why-a-sidecar-not-an-excel-add-in)
- [Features](#features)
- [Architecture](#architecture)
- [Quick start](#quick-start)
- [Roadmap](#roadmap)
- [License](#license)
- [Changelog](#changelog)

---

## Why a sidecar, not an Excel add-in?

|                       | Sidecar (this project) | Excel add-in (VSTO / Office JS) |
|-----------------------|:----------------------:|:-------------------------------:|
| Dev environment       | Python + PySide6 + pywin32 | VS + Office Dev Tools / Node + Yeoman |
| Deployment            | One PyInstaller EXE     | ClickOnce / MSI + trust-center hoops |
| GUI freedom           | Full Qt widget set     | WPF only (VSTO) / HTML sandbox (Office JS) |
| Excel cross-version   | COM stable 2007 → 365  | Office JS spec shifts between versions |
| External / LLM calls  | Zero friction          | Office JS sandbox restrictions |

The sidecar path is **simpler to build, simpler to ship, simpler to extend**.
COM automation has been stable since Excel 2007 and will almost certainly remain
stable for a decade — we rely on that.

---

## Features

### Formula engine (v1.0.2 — 5 core formulas)

Every formula ships with a complete **guide → write → diagnose** pipeline:

| Formula | Intent keywords (examples) | Wizard steps | Diagnostics |
|---------|----------------------------|--------------|-------------|
| **VLOOKUP** | 匹配、查找、对应、vlookup | 4 steps (lookup / table / col / match) | `#N/A` / `#REF!` / `#VALUE!` → spelling / wildcard / IFERROR wrap |
| **SUMIF** | 求和、汇总、条件求和 | 3 steps (range / criteria / sum_range) | `#VALUE!` / `#NAME?` → quoting / spelling |
| **COUNTIF** | 计数、个数、次数 | 2 steps (range / criteria) | `#VALUE!` → quoting hints |
| **IFERROR** | 容错、错误避免、#N/A | 2 steps (inner / fallback) | `#NAME?` → inner formula check first |
| **CONCAT / CONCATENATE** | 拼接、合并文本 | 1 step (comma-separated items) | `#NAME?` → CONCATENATE fallback (older Excel) |

Each formula also carries syntax definition, usage examples, usage tips, and
**auto-absolute (`$`) range anchoring** — copy the formula down a column and the
lookup table won't shift by accident.

### UI

- **Frameless floating window** — draggable, always-on-top, minimize / close
- **Home screen** — keyword search → ranked formula cards
- **Wizard screen** — step-by-step parameter input with live hints, warnings, and a
  "pick from Excel" button (reads active cell / selection via COM)
- **Preview screen** — live formula preview, one-click write to active Excel cell,
  clipboard copy, auto-diagnosis with **apply-fix** buttons
- **EN / ZH language toggle** in the title bar; persists across sessions

---

## Architecture

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
│ mkr / intent.py  keyword → formula   │──▶│ mkr / formulas /     │
│                                      │   │ base.py + vlookup, ... │
│                                      │   │ @register() decorator │
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
| `mkr/i18n.py` | All zh / en strings — single `tr(key, lang)` |
| `mkr/config.py` | Persistent config — language, window geometry, recent formulas, auto-connect flag. Portable when frozen (`<exe>/config.json`), falls back to `%APPDATA%\ClaraEasyExcel` |

---

## Quick start

### Environment

- **Windows** only (pywin32 + Excel COM)
- Python ≥ 3.11 (tested on 3.15)
- Microsoft Excel installed (2019 / 2021 / 365 tested)

### Install

```powershell
cd ClaraEasyExcel

# create a venv (recommended but optional)
python -m venv .venv
.\.venv\Scripts\activate

pip install -e .          # runtime deps: PySide6 + pywin32
pip install -e ".[dev]"   # + pytest + ruff
```

### Run from source

```powershell
# double-click run.bat, or:
pythonw entry.py          # no console window
# or:
python entry.py           # with console (useful for debugging)
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
## Changelog

### Latest — v1.0.2 (2026-11-26)

- Fixed `run.bat` failing on Windows CMD with "'M' is not recognized"
  - **Root cause**: the .bat file was written with Unix LF-only line endings
    — Windows CMD requires CRLF (`\r\n`). Every multi-line `if ( ... )` block
    and `REM` comment got mangled into garbage commands, and `REM =====`
    degenerated into the `M` the user saw on screen.
  - **Fix**: rewrote `run.bat` using only simple, single-line `if / goto`
    branches (no multi-line blocks) and forcibly re-saved as **ASCII + CRLF**
    with no BOM. Verified: 43 CR + 43 LF pairs, every line ends with `0D 0A`.
- Quick test — run.bat now launches `pythonw.exe` directly, no errors on console.

### v1.0.1 (2026-11-26)

- Added `run.bat` one-click launcher (auto-detects venv / pythonw fallback)
- Split bilingual README into two stand-alone files: `README.md` (English) + `README_CN.md` (中文)
- Each README now has a top-bar language switch link and a direct anchor to the changelog
- Bumped version to **1.0.1** (`mkr/__init__.py` + `pyproject.toml` — gui title updates automatically)

### v1.0.0 (2026-11-26) — first MVP

- 5 core formulas implemented: VLOOKUP, SUMIF, COUNTIF, IFERROR, CONCAT (with CONCATENATE alias)
- Each formula ships with intent keyword detection, 1~4 step parameter wizard, auto-absolute `$` anchoring, syntax definition, usage tips, and per-error-type diagnostics
- PySide6 floating sidecar window: frameless, draggable, always-on-top, blue-only palette
- Three-screen flow: Home → Wizard → Preview
- COM bridge (`mkr/excel_bridge.py`) — connect / disconnect / is_connected / get_active_cell / set_formula / get_cell_error
- EN / ZH language toggle with session persistence; language, window geometry, recent formulas persisted to `config.json` (portable on EXE, fallback to `%APPDATA%`)
- 14 unit tests for intent detection and formula build pipeline; ruff enforced
