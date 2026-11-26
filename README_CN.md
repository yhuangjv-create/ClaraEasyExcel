# ClaraEasyExcel

<div align="center">

![版本](https://img.shields.io/badge/version-1.0.2-blue)
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

### 公式引擎（v1.0.2 — 5 个核心公式）

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

### 最新 — v1.0.2（2026-11-26）

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
