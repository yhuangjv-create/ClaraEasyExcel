# ClaraEasyExcel

<div align="center">

![Version](https://img.shields.io/badge/version-1.3.1-blue)
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

### Formula engine (v1.3.1 — 5 core formulas)

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
  "pick from Excel" button: switch to any worksheet, then click or drag-select
  once and it is filled in automatically (sheet-qualified, e.g. `Sheet2!B:D`)
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

### Latest — v1.3.1 (2026-11-27)

- **Fixed: you can now pick a range on any worksheet, and drag as long as you
  like.** Reported on 国开行转账明细.xlsx, step 2 (框选区域, target B:D on
  Sheet2) — the Sheet2 tab itself was captured as the pick, and a long drag was
  cut off at a single column.
  - **Worksheet switches are navigation, not a pick.** Clicking another sheet
    tab changes the selection to that sheet's active cell, which the old
    address-only comparison treated as “the user chose this”.  The poll now
    re-baselines on a sheet change and never confirms it.
  - **Nothing is captured while the left mouse button is held.** A B→D
    header drag pauses on `B:B` long enough for a poll; confirming there froze
    the range at one column, and popping the window back to the front also
    aborted the drag.  The button state is read with `GetAsyncKeyState`, so a
    drag is captured only after you release the mouse.
  - **Picked references now carry the worksheet name** (`Sheet2!B:D`), so the
    generated formula points at the right sheet: `VLOOKUP(A2, Sheet2!$B:$D, 3,
    FALSE)`.  Whole-column / whole-row selections are simplified from
    `B1:D1048576` / `A3:XFD3` to `B:D` / `3:3`.
  - Regression tests in `tests/test_pick_flow.py`.

### v1.3.0 (2026-11-27)

- **Picking a cell now confirms with a single selection in Excel.** Clicking
  “Pick cell / Pick range” minimises the sidecar so Excel is fully in front;
  click or drag-select the target once and the address is filled in
  automatically (~400 ms) while the window pops back by itself — no more
  switching back to the software to click “confirm”.
- Poll interval dropped from 500 ms to 200 ms, and a selection must be stable
  for two consecutive polls: a single click captures instantly, while a
  drag-select is captured only after the drag ends (never halfway through).
- Every cell / range picker in the wizard shares this one-click-to-confirm
  behaviour.
- If no new selection is detected within 20 s the window restores itself and
  stays in pick mode, so “Read selected cell” remains available as a manual
  fallback (e.g. to pick the cell that was already active when pick mode
  started).

### v1.2.3 (2026-11-27)

- **Fixed: the wizard jumped several steps at once** (e.g. after filling
  VLOOKUP steps 1-4, going back home and reopening, one “Next” click jumped
  from step 1 to step 4).
  Root cause: `_render_step()` called a bare `btn_next.disconnect()`, which in
  PySide6 raises `TypeError` (“not enough arguments”) and disconnects
  **nothing** — so every rendered step added one more `clicked` connection and
  a single click fired `_on_next` several times.
- `btn_next` is now wired through a tracked slot (`_connect_next`), so it
  always has exactly one handler.  Added a regression test
  (`tests/test_wizard_flow.py`) asserting one click = one step and that
  reopening the wizard starts at step 1.

### v1.2.2 (2026-11-27)

- **Fixed: “says it wrote to A1, but nothing appears in my file”.**
  Root cause: when Excel is opened from Explorer / the Start menu it usually
  is **not** in the COM ROT, so the old code silently started a second, blank
  Excel and wrote there (its ActiveCell is always A1).  Worse, leftovers of
  those hidden instances stayed registered and hijacked every later write.
- The bridge now **binds to the visible Excel window you are working in**:
  it enumerates `XLMAIN` windows and asks MSAA for the native object behind
  the sheet window (`OBJID_NATIVEOM`), preferring the foreground window and
  windows with a real workbook file.
- It **never starts a second Excel** while any Excel window exists, and a
  ROT instance is trusted only if its window is one you can actually see.
- The status bar now shows the bound workbook name
  (e.g. `已连接 · 国开行转账明细.xlsx`), and a successful write reads
  `✓ 已写入 — R2 · 国开行转账明细.xlsx`, so it is obvious **where** it went.

### v1.2.1 (2026-11-27)

- **Default colour**: the window now starts on the **deep purple** preset
  (`#5C2D91`) instead of the Windows accent colour.
- The theme menu gained a first entry **“Deep purple (default)”**, so the
  default is one click away after switching to another colour.
- Following the Windows accent and custom colours still work exactly as in
  v1.2.0.

### v1.2.0 (2026-11-27)

- **New: theme button** in the title bar (“Theme”).
  - **Follow Windows accent** (default) — reads the Windows 10/11 accent
    colour from `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent`
    (`AccentColorMenu`), so the sidecar matches your system theme.
  - **Custom colour…** — opens a colour picker; the whole palette is then
    derived from that single colour.
  - Every other colour (cards, inputs, preview box, borders, buttons, hover,
    warning) is **auto-derived** from the base.  Text colour follows WCAG AA:
    light text is used only when white reaches **4.5:1** on the base,
    otherwise dark ink on light cards.
  - The choice is persisted in `config.json` (`"theme": "auto"` or
    `"#RRGGBB"`).

### v1.1.2 (2026-11-27)

- **UI theme**: window background is now **deep purple** (`#5C2D91`), with
  deeper purple cards/inputs (`#4E2578` / `#3A1A5C`), lavender borders and
  near-white text — back to a dark surface theme, tuned to the new hue.
- Still **no green anywhere** (verified by an offscreen render + per-pixel
  scan).

### v1.1.1 (2026-11-27)

- **UI theme**: the window background is now Windows 11's built-in
  **deep orange-yellow** accent (`#F7630C`).  Cards, inputs, the formula
  preview box and ghost-button borders were re-tuned to warm cream surfaces
  with dark-ink text, and the frame edge uses a light cream line — the
  interface stays readable on a bright warm base.
- Still **no green anywhere** (verified by an offscreen render + per-pixel
  scan).

### v1.1.0 (2026-11-27)

- **UI theme**: the whole interface now uses Windows 11's built-in **“Iris”**
  accent (`#8764B8`) as its base colour; cards, inputs and the formula
  preview box use deeper iris shades so the hierarchy stays readable.
- **No green anywhere**: every green element is gone — the `✅` emoji in
  status/button labels was replaced with the text glyph `✓`, which inherits
  the theme colour.  Verified by rendering the window offscreen and scanning
  every pixel: **0 green pixels**.
- Palette tokens in `gui.py` were regrouped into one Iris scale
  (`_BG` / `_BG_CARD` / `_BG_DEEP` / `_BORDER` / `_ACCENT`), so future theme
  tweaks stay a one-block edit.

### v1.0.9 (2026-11-27)

- **Fixed**: "Pick cell" (选单元格) in the VLOOKUP wizard could not select a cell and froze the app
  - **Root cause 1 — modal dialog deadlock**: the button opened a **modal
    `QMessageBox`** on an always-on-top, frameless window.  The nested Qt
    event loop blocked the main window, covered Excel so no cell could be
    clicked, and on dismissal fired a COM read while that loop was still
    unwinding — the UI thread deadlocked against Excel.
  - **Root cause 2 — the window vanished**: the window used `Qt.Tool`, which
    gets **no taskbar button**.  Any `showMinimized()` (pick mode or the
    title-bar “—”) hid the app with no way to restore it.
  - **Fix — live, non-blocking pick (no dialogs, no minimise)**:
    1. Clicking "Pick cell" enters pick mode; a `QTimer` polls Excel's
       selection every **500 ms** through a silent COM read
       (`read_selection_address` — skips the ScreenUpdating/EnableEvents
       toggling so Excel does not flicker while you drag).
    2. As soon as you click another cell — stable across two polls, so
       drag-selecting a range is not captured mid-drag — the address is
       filled in automatically.
    3. The button turns into **“✓ Read selected cell”** as a manual
       fallback; after ~30 s with no change a hint asks you to click it.
    4. No `QMessageBox` anywhere in the pick flow and the window is never
       hidden, so nothing can deadlock or disappear.
  - **Also fixed**: `Qt.Tool` removed from the window flags, so the window
    now has a taskbar button and the title-bar minimise button is no longer
    a one-way trip.

### v1.0.8 (2026-11-26)

- **Fixed**: CONCAT / CONCATENATE both appeared as separate cards on the home screen
  - **Root cause**: `concatenate.py` stacked two `@register` decorators — one for
    `"CONCAT"`, one for `"CONCATENATE"` — and `register()` blindly wrote both into
    the `_REGISTRY` dict.  `all_formula_names()` then returned both keys, which the
    GUI rendered as two CONCAT cards.
  - **Fix** (minimal, 3 changes):
    1. `base.py` registry split into `_REGISTRY` (canonical names only) +
       `_ALIASES` (`alias → canonical`).  New `_INSTANCE_TO_NAME` reverse map
       lets `register()` detect when two decorators hit the same class and
       automatically demote subsequent ones to aliases.
    2. `get_formula(name)` now does a 2-step lookup: canonical first, then
       alias → same formula instance resolves whether you type CONCAT or
       CONCATENATE in the intent search.
    3. Decorator order swapped in `concatenate.py` so **CONCAT** (modern
       Excel 2019+ name, recommended) becomes the canonical card and
       **CONCATENATE** is stored as an alias.
  - **Verified**: `all_formula_names()` returns exactly 5 items (no dupes);
    `get_formula("CONCAT") is get_formula("CONCATENATE")` → True.

### v1.0.7 (2026-11-26)

- **Fixed**: click "Write to active cell" → always wrote to A1 (user's real cell untouched)
  - **Real root cause** (found by step-by-step COM trace of a real open Excel):
    * The user's Excel window (started from Start menu) does **NOT** register
      itself in the COM Running Object Table — not even after waiting 20s.
      Only a `win32com.client.Dispatch("Excel.Application")` instance
      registers.  So our connect() Phase-1 retries GetActiveObject × 15 would
      always fail — the user's Excel simply isn't findable via GetActiveObject.
    * v1.0.6 then fell through to Dispatch **with `Visible=False`**, which
      silently spawned a background Excel that DOES register in ROT.  After
      that, GetActiveObject **always** returned the Dispatch instance.  Every
      subsequent _rebind() → GetActiveObject call was also returning our
      hidden Dispatch proxy, whose ActiveCell is always A1 (fresh workbook).
      So we were writing into an invisible Excel; the user's real Excel
      stayed empty.  That's why every click reported "A1".
  - **Fix**:
    1. **NEVER Dispatch Visible=False** — if we must Dispatch, start it
       Visible=True so the user sees it and knows we created it.
    2. **connect() now has 3 deterministic phases**:
       * Phase 1 — retry GetActiveObject up to **15×1s** (user's Excel might
         take a long time to register).
       * Phase 2 — if still nothing, Dispatch Visible=True, mark
         `_we_started_it = True` (sticky — stays True forever).
       * Phase 3 — retry GetActiveObject once more after Dispatch registered.
         Now GetActiveObject WILL succeed because Dispatch is in ROT; we
         swap to that proxy but KEEP `_we_started_it = True`.
    3. **_rebind() no longer resets `_we_started_it`** — it's a sticky flag
       set ONCE at connect time.  _rebind only swaps the proxy if it can;
       otherwise keeps what we have.
  - **Verified**: live trace — connect() binds correctly to user's Excel
    (Visible=True, Caption='工作簿1 - Excel'), writes CONCAT formula to the
    active cell, reads back Value/Formula correctly.  co_uninit does NOT quit
    user's Excel (`_we_started_it=False`).

### v1.0.6 (2026-11-26)

- **Fixed**: "Write to active cell" writes into a hidden Excel, not the user's visible one
  - **Root cause traced with a live COM trace**: when `GetActiveObject` fails
    (COM ROT registration lags 1-3s after Excel starts — even when the process
    is alive), our `connect()` silently fell through to `Dispatch`, which
    spawns a **brand-new hidden Excel instance**.  The proxy stayed bound to
    that invisible instance forever after.  Every subsequent write landed in
    the Dispatch instance; the user's visible Excel stayed empty.  They saw
    "does nothing" — we were writing, just into the wrong Excel.
  - **Fix — `_rebind()` before EVERY COM call**: every `get_active_cell`,
    `set_formula`, `get_cell_error`, `set_value`, etc. now starts by calling
    `_rebind()`, which force-runs `win32com.client.GetActiveObject` **right
    before the operation** and swaps the proxy.  This guarantees we always
    write into whichever Excel window the user is currently looking at —
    regardless of what happened at `connect()` time.
  - **Extra safeguards**:
    1. `connect()` retries `GetActiveObject` **6 times × 1 second** before
       falling through to `Dispatch` — gives COM registration plenty of lag
       time.
    2. The Dispatch fallback now starts Excel `Visible=False` (background
       only) and sets `_we_started_it=True` so we **close only our own**
       instance at shutdown — the user's Excel stays open.
    3. `co_uninit()` only calls `Quit()` when `_we_started_it=True` —
       absolutely never closes a user-opened Excel.
  - **Verified**: live E2E trace with a real Excel already open — write to
    active cell A1, read back `Value=2.0`, `Formula==1+1`.  `_rebind()`
    always returns the correct instance.

### v1.0.5 (2026-11-26)

- **Fixed**: clicking "Write to active cell" does **nothing visible** (formula never appears in Excel)
  - **Three stacked root causes**, all found with a step-by-step COM bridge trace:
  - **1. `Dispatch` starts Excel with no workbook** — `win32com.client.Dispatch("Excel.Application")` creates a fresh Excel process but does NOT open a workbook, so `ActiveWorkbook`, `ActiveSheet`, and `ActiveCell` are all `None`. Any call that touches them returns `NoneType has no attribute X`.  Fix: after Dispatch we now call `Workbooks.Add()`, and we also re-check after GetActiveObject in case the user closed all workbooks.
  - **2. `.Address(False, False)` is obsolete** — on current Excel (365) dynamic-dispatch proxies, `.Address` is a **string property** that returns `"$A$1"` directly — it's no longer callable.  `Address(False, False)` triggers `TypeError: 'str' object is not callable`.  Fix: new `_a1(range)` helper computes A1-style address deterministically from `Row` + `Column` (never depends on COM method signatures that change between Excel versions).
  - **3. `_XL_ERROR_CODES` table was completely wrong** — the integer error codes returned by `Range.Value` on this Excel don't match what the handwritten table assumed.  All 7 error codes (#DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!) were re-mapped by live COM tests on Excel 365, and the table now reflects the actual codes we get back.
  - **What we tried first** and rejected:
    - QThread + `QMetaObject.invokeMethod` with QueuedConnection — the thread's STA COM proxy became invalid as soon as `CoInitialize` was called twice (COM apartment + QThread event-loop incompatibility).  We simplified back to **single-threaded on the GUI thread**.  Every COM call takes under 50 ms (confirmed by trace); `DisplayAlerts=False / ScreenUpdating=False / EnableEvents=False` already shields us from Excel-side blocking, so the GUI no longer freezes either.
  - **What we actually delivered**: `excel_bridge.py` is back to a straightforward sync wrapper on the GUI thread — just `co_init()` once at app start, all calls protected by `_with_protection()`, `co_uninit()` once at shutdown.  No more signal chains, `_pending_*` state machines, or QThread management.

### v1.0.4 (2026-11-26)

- **Fixed**: clicking "Write to active cell" froze the entire app
  - **Root cause**: every Excel COM call (`get_active_cell`, `set_formula`,
    `get_cell_error`, …) was being executed synchronously on the **PySide6 GUI
    thread**.  COM proxy calls wait on Excel's main thread; if Excel was busy
    (background recalculation, an open `DisplayAlerts` dialog, a VBA macro
    still running, or just sluggish formula calculation), the call blocked the
    GUI event loop and the window looked dead.
  - **Fix**: rewrote `excel_bridge.py` as `ExcelBridgeThread(QThread)`:
    1. All `win32com` COM proxies are now created and **only ever touched
       inside a dedicated QThread**.  `pythoncom.CoInitialize()` lives there
       too.  GUI threads can't touch COM — period.
    2. GUI requests go through `QMetaObject.invokeMethod(..., QueuedConnection)`
       to the worker thread's slots.  Answers come back via Qt `Signal`s.
    3. `ScreenUpdating=False`, `EnableEvents=False`, `DisplayAlerts=False`
       wrap every Excel write so background Excel work can't block the call.
  - **GUI changes**: `_write_to_excel`, `_show_preview` and Wizard's "pick
    cell from Excel" are all async signal chains now.  No more `bridge.xxx()`
    returning a dict that the GUI uses synchronously — every call that touches
    Excel goes through `request_xxx()` + a matching slot.
  - **Verified**: 14/14 tests green, 0 ruff warnings, module import OK.

### v1.0.3 (2026-11-26)

- Eliminated the CMD **black console window** when launching `run.bat`
  - **Root cause**: double-clicking any `.bat` necessarily spawns a CMD console; using `pythonw.exe` inside the bat only hides Python's console, not CMD's.
  - **Fix**: split the launcher into two files:
    1. `run.bat` — only detects which Python to use, then calls `start "" wscript run.vbs "python_path"` and **immediately** `exit /b 0`. The CMD window closes in 1~2 frames.
    2. `run.vbs` — uses `WScript.Shell.Run pythonCmd, 0, False`. The second argument `0` means "hidden window"; this is the official Windows way to launch a GUI app with zero flash.
  - **Tested**: `pythonw.exe` launches correctly, no lingering `cmd.exe` process, no visible black window.
  - This also means `run.vbs` can be double-clicked directly as an alternative entry point (falls back to `pythonw.exe` on PATH if no argument).

### v1.0.2 (2026-11-26)

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
