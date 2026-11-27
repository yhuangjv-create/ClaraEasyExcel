# ClaraEasyExcel

<div align="center">

![版本](https://img.shields.io/badge/version-1.3.11-blue)
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

### 公式引擎（v1.3.11 — 5 个核心公式）

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
- **向导页** — 逐步填写参数，实时提示、警告、"从 Excel 选单元格"按钮：
  可先切换到任意工作表，点击或框选一次即自动填入（带表名，如 `Sheet2!B:D`）
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

### 最新 — v1.3.11（2026-11-27）

- **状态栏渐变里不再出现绿色。** 原来的冷端是冰青色 `#8BE9FD`，它与白色或紫色背景
  混合时绿通道会占上风，看起来发绿。现在改为金黄 `#FFD479` → 白 → 蓝紫 `#9DB8FF`，
  每一段都是偏蓝的；浅色主题的冷端也从青色换成 `#1D4ED8`。测试中断言冷端
  B > G，避免绿色再次混入。

### v1.3.10（2026-11-27）

- **状态栏文字改为三段渐变**，不再是单一颜色：金黄 → 白 → 冰蓝横贯文字，在紫色
  背景上更醒目，同时保留呼吸式闪烁。`QLabel` 只能绘制纯色，因此由
  `StatusLabel.paintEvent` 用渐变画笔绘制文字，且渐变按文字实际宽度铺开（不会浪费
  在右侧空白上）。浅色主题下自动换成深琥珀 → 近黑 → 深青，保证对比度。

### v1.3.9（2026-11-27）

- **左下角的状态文字改为"呼吸式"闪烁**：透明度在 1.0 与 0.45 之间以 2.6 秒周期、
  InOutSine 曲线缓慢起伏，比生硬的开关式闪烁更有质感。`QLabel` 本身没有透明度
  属性，因此用 `QGraphicsOpacityEffect` + `QPropertyAnimation` 实现。新增测试
  `tests/test_status_pulse.py`。

### v1.3.8（2026-11-27）

- **左上角标题只保留 `ClaraEasyExcel`。** 后面的"— Excel 智能助手 vX.Y.Z"在紧凑
  窗口里本来就被截断了，现在连横线一起删除。

### v1.3.7（2026-11-27）

- **修复：卡片第一行顶部被裁、两行看起来压在一起。** 列表行的样式里还留着
  `padding: 6px`，卡片实际拿到的矩形比代码预留的高度少 12px（宽度也少 12px），
  Qt 于是把文字标签压到最小高度以下，粗体标题的上缘就被切掉了。现在去掉行的
  padding，行高改为 `max(heightForWidth, minimumSizeHint) + 4`。
- 第一行字号回到 12px（v1.3.5 曾把它加到 13px）。
- 新增回归测试 `test_card_text_is_never_squeezed`：在 320×420 下渲染首页，断言
  卡片和文字标签都不会小于自身最小高度。

### v1.3.6（2026-11-27）

- **修复公式卡片内两行文字挤在一起的问题**：加粗标题与下面的示例公式原来只隔
  1px。现在行间距 5px，标题下方再加 3px 内边距，两行层次分明。

### v1.3.5（2026-11-27）

- **公式卡片不再显示下面的说明段落。** 每张卡片现在只有标题行——公式名 + 一句
  简要说明（如 `VLOOKUP 条件查找`）——加上示例公式。省出来的高度给了这一行和行间
  距，5 个公式正好铺满首页，不再出现下方大片空白。
- `tests/test_home_cards.py` 现在锁定"卡片只有两行文字"。

### v1.3.4（2026-11-27）

- **窗口整体缩小：由 420×600 改为 320×420（宽约 -24%，高约 -30%）**，各处
  留白同步收紧，遮挡 Excel 的面积明显减少。
  - 三个页面、标题栏、公式卡片、列表行的边距与间距全部缩小；按钮、输入框、
    公式框的内边距和字号也各降一档。
  - 卡片宽度改为跟随可视区（以前按最长示例公式撑开，在窄窗口里会出现横向滚动条
    并截断文字），过长的示例公式自动换行。
  - 保存的窗口尺寸若仍等于旧默认值 420×600，视为"从未手动调整过"，将按新的紧凑
    尺寸打开；用户自己拖过的尺寸保持不变。
  - 最小尺寸调整为 300×360。

### v1.3.3（2026-11-27）

- **修复：鼠标扫过公式卡片时会一行一行地亮起。** 卡片样式用了裸 `QWidget`
  选择器，把卡片里每个 `QLabel` 也一起选中了——每行文字各自绘制圆角背景、各自
  带 `:hover`，所以鼠标沿三行说明往下扫时，亮起的是光标所在的那一行而不是整张卡片。
- 现在样式限定为 `QWidget#formulaCard`，内部文字标签背景透明且不接收鼠标事件：
  **整张卡片**作为一个整体高亮，鼠标不会"卡"在某一行上。
- 顺带修正：卡片示例公式跟随界面语言（英文界面此前显示的是中文示例）。
  新增回归测试 `tests/test_home_cards.py`。

### v1.3.2（2026-11-27）

- **修复：「下一步」左边的按钮是空白的。** `btn_prev` 虽然创建了、也绑定了返回
  的处理函数，但 `_refresh_static_texts()` 从未给它设置文字，所以「下一步」旁边
  显示的是一个没有字的按钮。现在它始终显示 **上一步 / Back**，切换语言后也正确。
- 新增回归测试：`tests/test_wizard_flow.py`。

### v1.3.1（2026-11-27）

- **修复：可以框选任意工作表的区域，且想拖多久就拖多久。** 以
  国开行转账明细.xlsx 第 2 步（框选区域，目标是 Sheet2 的 B:D 列）为例：
  原先连点 Sheet2 工作表标签都会被当成"选好了"，一次长距离拖选也会被截断成一列。
  - **切换工作表属于导航，不算选择。** 点击其他工作表标签会让选区变成新表的
    活动单元格，旧逻辑只比较地址，于是把它当成了"用户选的就是它"。现在检测到
    工作表变化时只更新基准、绝不确认。
  - **按住左键时不确认任何选区。** 从 B 拖到 D 的过程中会在 `B:B` 停留超过一次
    轮询；在那里确认会把区域固化成一列，而且把窗口弹回前台还会直接中断拖动。
    现在通过 `GetAsyncKeyState` 读取真实按键状态，只在松开鼠标后才确认。
  - **取到的引用带工作表名**（`Sheet2!B:D`），生成的公式因此指向正确的表：
    `VLOOKUP(A2, Sheet2!$B:$D, 3, FALSE)`。整列 / 整行选区自动简化：
    `B1:D1048576` → `B:D`，`A3:XFD3` → `3:3`。
  - 新增回归测试 `tests/test_pick_flow.py`。

### v1.3.0（2026-11-27）

- **选单元格现在在 Excel 中选一次即自动确认。** 点击「选单元格 / 框选区域」后，
  侧边窗口会自动最小化，让 Excel 完全置于前台；在 Excel 中点击或框选目标一次，
  约 400 毫秒后地址会自动填入，窗口也会自动弹回——无需再切回本软件点确认。
- 轮询间隔从 500 毫秒缩短到 200 毫秒，且要求连续两次轮询结果一致才确认：单击
  即时捕获，拖动框选则在拖动结束后才捕获，不会取到半途的选区。
- 向导中所有选单元格 / 选区域的步骤都统一为这套「一次选择即确认」的逻辑。
- 若 20 秒内未检测到新选择，窗口会自动弹回并保持取数模式，仍可点「读取选中单元格」
  手动兜底（例如想取进入模式时就已选中的那个单元格）。

### v1.2.3（2026-11-27）

- **修复：向导一次点击跳过好几步**（例如填完 VLOOKUP 第 1～4 步、返回首页再
  进入，点一次「下一步」会从第 1 步直接跳到第 4 步）。
  根因：`_render_step()` 里调用的是无参 `btn_next.disconnect()`，PySide6 会抛
  `TypeError`（“not enough arguments”）且**什么也没断开**，于是每渲染一步就多
  叠加一个 `clicked` 连接，点一次相当于连点多次。
- 现在 `btn_next` 通过记录槽位的 `_connect_next` 绑定，始终只有一个处理方法。
  新增回归测试 `tests/test_wizard_flow.py`：一次点击只前进一步、重新进入向导
  必从第 1 步开始。

### v1.2.2（2026-11-27）

- **修复「提示已写入 A1，但文件里什么都没有」。**
  根因：从资源管理器 / 开始菜单打开的 Excel 通常**没有注册到 COM ROT**，
  旧代码就会悄悄再启动一个空白 Excel 往里面写（它的活动单元格永远是 A1）；
  这些隐藏实例还会残留并劫持后续所有写入。
- 现在改为**绑定到你正在使用的那个可见 Excel 窗口**：枚举 `XLMAIN` 窗口，
  通过 MSAA 取工作表窗口背后的原生对象（`OBJID_NATIVEOM`），优先当前前台
  窗口和有真实文件路径的工作簿。
- 只要有 Excel 窗口存在就**绝不再启动第二个 Excel**；ROT 实例只有在它的
  窗口确实可见时才被采用。
- 状态栏会显示已绑定的工作簿名（如 `已连接 · 国开行转账明细.xlsx`），
  写入成功后显示 `✓ 已写入 — R2 · 国开行转账明细.xlsx`，写到哪里一目了然。

### v1.2.1（2026-11-27）

- **默认颜色**：窗口启动时使用**深紫色**预设（`#5C2D91`），不再默认跟随
  Windows 强调色。
- 主题菜单新增首项**「深紫色（默认）」**，切到别的颜色后可一键切回默认。
- 「跟随 Windows 主题色」与自定义颜色功能与 v1.2.0 一致，仍可正常使用。

### v1.2.0（2026-11-27）

- **新增：标题栏「主题」按钮**。
  - **跟随 Windows 主题色**（默认）：从
    `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Accent`
    （`AccentColorMenu`）读取 Win10/11 强调色，窗口自动与系统主题一致。
  - **自定义颜色…**：打开取色器，整套配色由这一个颜色推导。
  - 其余颜色（卡片、输入框、公式预览框、描边、按钮、悬停、警告）全部
    **自动推导**。文字颜色按 WCAG AA 选择：白色在底色上达到 **4.5:1** 才用
    浅色文字，否则改用浅色卡片 + 深墨文字。
  - 选择会保存在 `config.json`（`"theme": "auto"` 或 `"#RRGGBB"`）。

### v1.1.2（2026-11-27）

- **界面配色**：窗口底色改为**深紫色**（`#5C2D91`），卡片与输入框使用更深一层的
  紫色（`#4E2578` / `#3A1A5C`），描边为浅薰衣草色，文字为近白色 —— 回到深色底主题，
  并按新色相重新调校。
- 依旧**全面无绿**（离屏渲染 + 逐像素扫描验证）。

### v1.1.1（2026-11-27）

- **界面配色**：窗口底色改为 Windows 11 自带的**「深橙黄色」**（`#F7630C`）；
  卡片、输入框、公式预览框与次要按钮边框统一改为暖调奶油底 + 深墨色文字，
  窗体描边使用浅奶油色 —— 在明亮暖色底上依然清晰可读。
- 依旧**全面无绿**（离屏渲染 + 逐像素扫描验证）。

### v1.1.0（2026-11-27）

- **界面配色**：整体底色改为 Windows 11 自带的**「鸢尾花色」**（`#8764B8`）；
  卡片、输入框与公式预览框使用更深一层的鸢尾紫，层次分明、文字清晰。
- **彻底去绿**：界面上所有绿色元素全部移除 —— 状态/按钮标签中的 `✅` emoji
  改为文字符号 `✓`（跟随主题文字颜色）。已用离屏渲染逐像素扫描验证：
  **绿色像素为 0**。
- `gui.py` 色板收敛为一组鸢尾花色阶（`_BG` / `_BG_CARD` / `_BG_DEEP` /
  `_BORDER` / `_ACCENT`），以后调色只需改这一处。

### v1.0.9（2026-11-27）

- **修复**：VLOOKUP 向导中的「选单元格」无法选取，且点击后程序卡死
  - **根因 1 — 模态框死锁**：按钮弹出了**模态 `QMessageBox`**，而主窗口是置顶无边框窗口。模态框的嵌套 Qt 事件循环阻塞了主窗口、挡在 Excel 前面导致无法点选；关闭对话框时又在嵌套循环尚未退栈时立刻发起 COM 读取 —— UI 线程与 Excel 互相死锁。
  - **根因 2 — 窗口消失**：窗口使用了 `Qt.Tool`，**不生成任务栏按钮**。任何 `showMinimized()`（取数模式或标题栏「—」）都会让程序隐藏且无法再调出。
  - **修复 —— 实时、非阻塞取数（无对话框、不最小化）**：
    1. 点击「选单元格」进入取数模式，用一个 `QTimer` 每 **500 毫秒**通过静默 COM 读取（`read_selection_address`，不切换 ScreenUpdating / EnableEvents，避免拖动选区时 Excel 闪烁）轮询 Excel 的选择区域。
    2. 只要你在 Excel 中点了另一个单元格（地址连续两次轮询一致 —— 因此拖动选择区域不会被中途截断），就自动填入地址。
    3. 按钮同时变为 **「✓ 读取选中单元格」** 作为手动兜底；约 30 秒仍未检测到新选择时，会提示你点击它。
    4. 取数流程中没有任何 `QMessageBox`，窗口也不会隐藏，因此不可能再死锁或消失。
  - **同时修复**：窗口标志去掉 `Qt.Tool`，现在有任务栏按钮，标题栏的最小化按钮不再是「有去无回」。

### v1.0.8（2026-11-26）

- **修复**：首页"CONCAT 拼接文本"卡片出现两个重复
  - **根因**：`concatenate.py` 叠了两个 `@register` 装饰器（CONCAT + CONCATENATE），旧版 `register()` 把两个都写进了 `_REGISTRY` dict。`all_formula_names()` 返回两个 key → GUI 渲染两张 CONCAT 卡片。
  - **修复**（最小 3 处改动）：
    1. `base.py` registry 拆成 `_REGISTRY`（仅 canonical 主名）+ `_ALIASES`（alias → canonical 映射）。新增 `_INSTANCE_TO_NAME` 反向表让 `register()` 检测到同一类被多装饰器命中时自动降级后续为别名。
    2. `get_formula(name)` 先查主名再查别名 —— 意图搜索里无论用户打 CONCAT 还是 CONCATENATE 都能搜到同一个公式实例。
    3. `concatenate.py` 调换装饰器顺序，让 **CONCAT**（Excel 2019+ 新版推荐名）成为首页展示的 canonical 卡片，CONCATENATE 存为别名。
  - **实测**：`all_formula_names()` 精确返回 5 项（无重复）；`get_formula("CONCAT") is get_formula("CONCATENATE")` → True。

### v1.0.7（2026-11-26）

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
