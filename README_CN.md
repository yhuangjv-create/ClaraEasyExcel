# ClaraEasyExcel

<div align="center">

![版本](https://img.shields.io/badge/version-1.0.7-blue)
[![许可证: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![平台](https://img.shields.io/badge/platform-windows-blue)

**轻量 Excel 外挂助手 — 悬浮置顶、一键启动、无需安装。**

[English](./README.md) · 中文 · [版本更迭](#版本更迭)

</div>

---

ClaraEasyExcel 是一款小而轻的 **悬浮外挂小窗口**，紧贴 Microsoft Excel 运行，帮助
Excel 新手快速编写常用内部公式和（规划中）VBA 宏：

- **智能意图识别** — 输入"匹配"或 lookup，ClaraEasyExcel 自动推荐 VLOOKUP
- **逐步参数向导** — 不用死记参数顺序和引号，跟着向导填就对
- **内置注意事项与警告** — 为什么查找表要加 `$`、什么时候用 FALSE / TRUE
- **实时错误诊断** — Excel 返回 `#N/A`？ClaraEasyExcel 告诉你原因和修正方法
- **一键写入** — 公式直接落到活动单元格，干净利落

我们以 **纯外挂** 形态交付（而非 Excel 插件）—— 一个 PyInstaller 打包的 EXE，
零注册表改动、零部署烦恼。详见 [为什么选外挂？](#为什么选外挂而非-excel-插件)。

---

## 目录

- [为什么选外挂而非 Excel 插件？](#为什么选外挂而非-excel-插件)
- [功能特性](#功能特性)
- [架构](#架构)
- [快速开始](#快速开始)
- [Roadmap](#roadmap)
- [许可证](#许可证)
- [版本更迭](#版本更迭)

---

## 为什么选外挂而非 Excel 插件？

|                       | 外挂（本项目）            | Excel 插件（VSTO / Office JS） |
|-----------------------|:-------------------------:|:------------------------------:|
| 开发环境              | Python + PySide6 + pywin32 | VS + Office 开发工具 / Node + Yeoman |
| 部署                  | 一个 PyInstaller EXE      | ClickOnce / MSI + 信任中心各种拦截 |
| GUI 自由度            | 完整 Qt 控件集            | VSTO 仅限 WPF；Office JS 受 HTML 沙箱限制 |
| Excel 跨版本兼容性    | COM 接口 2007 → 365 稳定  | Office JS 规范随版本变化       |
| 外部调用 / LLM 接入   | 零摩擦                    | Office JS 沙箱有限制           |

外挂路线 **开发更简单、发布更省心、扩展更自由**。COM 自动化接口自 Excel 2007
以来一直稳定，十年内大概率不会变 —— 我们就赌这一点。

---

## 功能特性

### 公式引擎（v1.0.7 — 5 个核心公式）

每个公式都自带完整的 **引导 → 生成 → 诊断** 流水线：

| 公式 | 意图关键词（示例） | 向导步数 | 诊断覆盖 |
|------|-------------------|----------|---------|
| **VLOOKUP** | 匹配、查找、对应、vlookup | 4 步（查找值 / 查找表 / 返回列 / 匹配方式） | `#N/A` / `#REF!` / `#VALUE!` → 拼写 / 通配符 / IFERROR 套层 |
| **SUMIF** | 求和、汇总、条件求和 | 3 步（条件区域 / 条件 / 求和区域） | `#VALUE!` / `#NAME?` → 引号 / 拼写检查 |
| **COUNTIF** | 计数、个数、次数 | 2 步（计数区域 / 条件） | `#VALUE!` → 引号提示 |
| **IFERROR** | 容错、错误避免、#N/A | 2 步（要保护的公式 / 出错返回值） | `#NAME?` → 先查内层公式 |
| **CONCAT / CONCATENATE** | 拼接、合并文本 | 1 步（逗号分隔的各项） | `#NAME?` → 旧版 Excel 降级为 CONCATENATE |

每个公式都附带语法定义、示例、使用技巧，以及 **自动绝对引用（`$`）** ——
公式下拉时查找表不会意外漂移。

### UI

- **无边框悬浮窗口** — 可拖动、置顶、可最小化 / 关闭
- **首页** — 关键词搜索 → 公式卡片列表
- **向导页** — 逐步填写参数，实时提示、警告、"从 Excel 选单元格"按钮
- **预览页** — 公式实时预览、一键写入、复制到剪贴板、自动诊断并提供"一键套用"修复
- **中英文切换** — 标题栏按钮切换，关闭时自动保存语言状态

---

## 架构

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
│                                      │   │ base.py + vlookup, ... │
│                                      │   │ @register() 装饰器    │
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

辅助模块：

| 模块 | 作用 |
|------|------|
| `mkr/i18n.py` | 全部中英文文本 —— 统一 `tr(key, lang)` 接口 |
| `mkr/config.py` | 配置持久化 —— 语言、窗口位置、最近公式、自动连接开关。PyInstaller 冻结时便携模式（EXE 同目录 `config.json`），不可写降级到 `%APPDATA%\ClaraEasyExcel` |

---

## 快速开始

### 环境

- **仅限 Windows**（pywin32 + Excel COM）
- Python ≥ 3.11（已在 3.15 测试）
- 已安装 Microsoft Excel（2019 / 2021 / 365 测试通过）

### 安装

```powershell
cd ClaraEasyExcel

# 创建虚拟环境（推荐但非必须）
python -m venv .venv
.\.venv\Scripts\activate

pip install -e .          # 运行时依赖：PySide6 + pywin32
pip install -e ".[dev]"   # + pytest + ruff
```

### 运行源码

```powershell
# 双击 run.bat，或者：
pythonw entry.py          # 无黑色控制台（推荐）
# 或者：
python entry.py           # 显示控制台（调试时用）
```

### 打包为单文件 EXE

```powershell
pyinstaller --onefile --windowed --name ClaraEasyExcel entry.py
# 输出：dist\ClaraEasyExcel.exe
```

### 开发检查

```powershell
ruff check .
pytest -v
```

---

## Roadmap

- **v1.1.0** — 新增 XLOOKUP、INDEX+MATCH、IFS、SUMIFS、COUNTIFS、IF、LEFT/RIGHT/MID、FIND、SUBTOTAL；VBA 宏模板引擎（预装 10 个常用宏）
- **v2.0.0** — 关键词匹配 → LLM 自然语言；支持嵌套公式生成与自然语言 VBA；可选 Ollama 本地后端

---

## 许可证

MIT — 详见 [LICENSE](./LICENSE)

---

<a id="版本更迭"></a>
## 版本更迭

### 最新 — v1.0.7（2026-11-26）

- **修复**：点击"写入活动单元格"永远提示写入 A1，实际什么都没写进用户的 Excel
  - **真正根因**（逐步 COM trace 实锤）：
    * 用户从开始菜单启动的 Excel 窗口，**根本不注册在 COM Running Object Table 里** —— 等 20 秒也没用。只有 `win32com.client.Dispatch` 创建的 Excel 实例才会注册。所以 connect() Phase-1 重试 GetActiveObject × 15 必然全部失败 —— 用户的 Excel 是找不到的。
    * v1.0.6 随后 fallback 到 Dispatch **`Visible=False`**，静默创建了一个后台 Excel（这个注册在 ROT）。之后所有 GetActiveObject 调用永远返回这个 Dispatch 实例。每个 _rebind() → GetActiveObject 也拿到这个隐藏代理，它的 ActiveCell 永远是 A1。所以我们的公式一直在写一个不可见的 Excel；用户眼前的 Excel 永远是空的 —— 这就是为什么每次点击都提示 "A1"。
  - **修复**：
    1. **绝不 Dispatch Visible=False** —— 要 Dispatch 就 Visible=True，让用户看得见
    2. **connect() 三阶段确定性连接**：
       * Phase 1 —— GetActiveObject 重试最多 **15×1s**
       * Phase 2 —— 还找不到就 Dispatch Visible=True，标记 `_we_started_it=True`（粘性）
       * Phase 3 —— Dispatch 后再 GetActiveObject 一次（这次必然成功），换 proxy 但保持 `_we_started_it=True`
    3. **_rebind() 不再重置 `_we_started_it`** —— 粘性标记一次设定终身不变
  - **实测验证**：connect() 正确绑到用户 Excel（Visible=True, Caption='工作簿1 - Excel'），写入 CONCAT 公式到活动单元格，读回 Value/Formula 正确。co_uninit 不退出用户 Excel。

### v1.0.6（2026-11-26）

- **修复**：点击"写入活动单元格"后什么都没发生 —— 实际上写到了一个后台隐藏的 Excel 实例
  - **根因追踪（COM trace 实锤）**：当 `GetActiveObject` 失败时（Excel 刚启动后 COM ROT 注册有 1~3s 延迟 —— 进程活着但 GetActiveObject 拿不到），`connect()` 悄悄 fallback 到了 `Dispatch` —— 这会创建一个**全新的 Excel 实例**。proxy 就一直绑在这个不可见的实例上。所有后续写入都进了 Dispatch 实例，用户眼前的 Excel 永远空的。
  - **修复 —— 每个 COM 调用前强制 `_rebind()`**：`get_active_cell` / `set_formula` / `get_cell_error` / `set_value` 等**每次操作前**先调 `_rebind()`，里面强制跑 `win32com.client.GetActiveObject` 重新换 proxy。这样无论 connect 时是什么状态，**永远写到用户当前正在看的那个 Excel 窗口**。
  - **额外防护**：
    1. `connect()` 重试 `GetActiveObject` **6 次 × 1 秒**才放弃 Dispatch —— 给 COM 注册足够的时间
    2. Dispatch fallback 现在以 `Visible=False`（后台静默）启动，并且设 `_we_started_it=True` —— 这样我们只在退出时关闭**自己创建的**实例，用户的 Excel 保持打开
    3. `co_uninit()` 只在 `_we_started_it=True` 时才调 `Quit()` —— 绝对不退出用户自己开的 Excel
  - **实测验证**：真实 Excel 已打开的场景下写入活动单元格 A1，读回 `Value=2.0`，`Formula==1+1`。`_rebind()` 始终返回正确实例。

### v1.0.5（2026-11-26）

- **修复**：点击"写入活动单元格"后 **Excel 里什么都没发生**（公式根本没被写入）
  - **三层叠加根因**，通过逐步 trace COM 桥接层全部定位：
  - **1. Dispatch 启动的 Excel 没有工作簿** — `win32com.client.Dispatch("Excel.Application")` 只创建 Excel 进程，**不自动打开工作簿**，所以 `ActiveWorkbook` / `ActiveSheet` / `ActiveCell` 全部是 `None`，任何访问都会报 `NoneType has no attribute X`。修复：Dispatch 后立即调用 `Workbooks.Add()`；GetActiveObject 后也二次检查用户是否关掉了所有工作簿。
  - **2. `.Address(False, False)` 已过时** — 当前 Excel（365）的 dynamic-dispatch 代理里，`.Address` 是 **字符串属性**（直接返回 `"$A$1"`），不是可调用方法。`Address(False, False)` 会抛 `TypeError: 'str' object is not callable`。修复：新增 `_a1(range)` 辅助函数，用 `Row` + `Column` 自行拼接 A1 地址 —— 确定性、不依赖 COM 方法签名、Excel 跨版本稳定。
  - **3. `_XL_ERROR_CODES` 表完全错误** — 本 Excel 版本 `Range.Value` 返回的整数错误码跟手写表完全对不上。7 种错误码（#DIV/0! / #N/A / #NAME? / #NULL! / #NUM! / #REF! / #VALUE!）全部按实测 COM 输出重新映射，表现在能正确识别。
  - **试过但放弃的方案**：
    - QThread + `QMetaObject.invokeMethod` QueuedConnection —— QThread 的 STA COM 代理在 `CoInitialize` 二次调用后直接失效（COM apartment 跟 QThread 事件循环不兼容），还不如同步方案。
  - **最终交付**：`excel_bridge.py` 回归 **GUI 线程同步调用** —— 只需在 app 启动时 `co_init()` 一次，所有 COM 调用走 `_with_protection()`（关闭 DisplayAlerts/ScreenUpdating/EnableEvents 防护 Excel 后台阻塞），退出时 `co_uninit()` 一次。无 signal 链、无 `_pending_*` 状态机、无 QThread 管理。

### v1.0.4（2026-11-26）

- **修复**：点击"写入活动单元格"后整个程序卡死
  - **根因**：所有 Excel COM 调用（读活动单元格、写公式、读单元格错误、……）都在 **PySide6 GUI 主线程**上同步执行。COM 代理调用内部需要等待 Excel 主线程响应 —— 一旦 Excel 正在做后台重算、弹出 `DisplayAlerts` 对话框、或者 VBA 宏还没跑完，COM 调用就会阻塞 GUI 事件循环，整个窗口看起来像"死"了。
  - **修复**：彻底重写 `excel_bridge.py` 为 `ExcelBridgeThread(QThread)`：
    1. 所有 `win32com` COM 代理都在独立 QThread 里创建，并且**只在这个线程里使用**。`pythoncom.CoInitialize()` 也只在这个线程里初始化。GUI 线程不碰 COM —— 一条都不碰。
    2. GUI 发起 Excel 请求时通过 `QMetaObject.invokeMethod(..., QueuedConnection)` 投递到 worker 线程的 slot；结果通过 Qt `Signal` 回传。
    3. 每次 Excel 写入前自动关闭 `ScreenUpdating=False`、`EnableEvents=False`、`DisplayAlerts=False`，防止 Excel 后台工作阻塞调用。
  - **GUI 同步改造**：`_write_to_excel`、`_show_preview`、Wizard 的"从 Excel 选单元格"全部改成异步 signal 链路。没有任何 `bridge.xxx()` 返回 dict 后 GUI 立即同步使用的情况 —— 所有碰 Excel 的调用都走 `request_xxx()` + 对应的 slot。
  - **验证通过**：14/14 单元测试、0 ruff 警告、模块 import OK。

### v1.0.3（2026-11-26）

- 彻底消除双击 `run.bat` 时出现的 CMD **黑色控制台窗口**
  - **根因**：双击任何 `.bat` 文件，Windows 必然先拉起一个 CMD 控制台；bat 里用 `pythonw.exe` 只能隐藏 Python 自己的黑窗，CMD 的黑框还是会出现。
  - **修复**：启动器拆分成两份文件：
    1. `run.bat` —— 只负责探测 Python 路径，然后调 `start "" wscript run.vbs "python路径"` 并 **立即** `exit /b 0`。CMD 窗口在 1~2 帧内就会关闭。
    2. `run.vbs` —— 用 `WScript.Shell.Run pythonCmd, 0, False`。第二个参数 `0` = 隐藏窗口；这是 Windows 官方推荐的 GUI 无闪现启动方式。
  - **实测验证**：`pythonw.exe` 正常启动，无遗留 `cmd.exe` 进程，屏幕上看不到任何黑窗。
  - `run.vbs` 也可以独立双击作为备用入口（无参数时回退到 PATH 里的 `pythonw.exe`）。

### v1.0.2（2026-11-26）

- 修复 `run.bat` 在 Windows CMD 下报 "'M' 不是内部或外部命令"
  - **根因**：.bat 文件以 Unix LF 行尾写入，而 Windows CMD 只认 CRLF（`\r\n`）。
    多行 `if ( ... )` 块与 `REM` 注释被 CMD 当成乱码解析，`REM =====`
    最终退化成用户看到的 `M`。
  - **修复**：重写 `run.bat` 为纯单行 `if / goto` 分支（避免多行块），
    并强制以 **ASCII + CRLF**（无 BOM）格式落盘。校验结果：43 个 CR + 43 个 LF，
    每行结尾都是 `0D 0A`。
- 快速验证通过：run.bat 直接启动 `pythonw.exe`，控制台无任何报错。

### v1.0.1（2026-11-26）

- 新增 `run.bat` 一键启动脚本（自动检测 venv / pythonw 回退）
- 原 README 拆分为两份独立文件：`README.md`（英文版）+ `README_CN.md`（中文版）
- 每份 README 顶部新增语言切换链接与版本更迭直达锚点
- 版本号提升至 **1.0.1**（`mkr/__init__.py` + `pyproject.toml` —— GUI 标题自动跟随同步）

### v1.0.0（2026-11-26）— 首次 MVP 发布

- 实现 5 个核心公式：VLOOKUP、SUMIF、COUNTIF、IFERROR、CONCAT（兼容 CONCATENATE）
- 每个公式都包含意图关键词识别、1–4 步参数向导、自动绝对引用（`$`）、语法定义、使用技巧、分错误类型的诊断建议
- PySide6 悬浮外挂窗口：无边框、可拖动、置顶、仅使用蓝色调色板
- 三屏流程：首页 → 向导 → 预览
- COM 桥接层（`mkr/excel_bridge.py`）— connect / disconnect / is_connected / get_active_cell / set_formula / get_cell_error
- 中英文切换 + 会话持久化；语言、窗口位置、最近使用公式保存到 `config.json`（EXE 同目录便携模式，不可写降级到 `%APPDATA%`）
- 14 个单元测试覆盖意图识别与公式生成管道；ruff 代码风格检查强制执行
