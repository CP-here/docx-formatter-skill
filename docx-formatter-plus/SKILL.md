---
name: "docx-formatter-plus"
description: "生成专业排版的 Word (.docx) 文档，支持 OMML 数学公式、表格框图、标准排版、PDF 渲染验证、XSD 模式验证、编辑现有文档、追踪修订和批注。当用户要求生成带公式的 Word 文档、用数学方程格式化文档、在 Word 中创建流程图样式图表、编辑现有 docx、或验证文档结构时调用。"
---

# DOCX 排版工具（增强版）

本工具用于生成专业排版的 Word (.docx) 文档，支持以下功能：

1. **标准排版**（黑体标题/宋体正文、标准页边距、行距）
2. **OMML 数学公式**（使用 docx Math API 的行内公式和块级公式）
3. **表格框图**（使用 Word 表格 + Unicode 箭头绘制流程图，不使用 emoji 或图片）
4. **引用上标**（正文中的 [1] 等引用标记自动渲染为上标）
5. **PDF 渲染验证**（可选：通过 LibreOffice 转 PDF 并生成图片做视觉检查）
6. **XSD 模式验证**（可选：验证生成的 docx 文件符合 OOXML 标准）
7. **编辑现有文档**（可选：解包 → 合并碎片 run → 编辑 XML → 重新打包）
8. **追踪修订**（可选：tracked changes / redlining 支持与验证）
9. **批注**（可选：管理 6 文件交叉链接的 Word 批注系统）

## 调用时机

- 用户要求生成带数学公式的 Word 文档（.docx）
- 用户要求用方程、下标、分式、求和符号格式化文档
- 用户要求在 Word 中创建流程图样式图表（使用表格 + 箭头）
- 用户要求将文档中的参考文献引用渲染为上标
- 用户要求"排版"或"格式化"文档为 Word
- 用户要求"使用与之前相同的排版"生成 Word 文档
- 用户要求编辑现有 .docx 文件（解包 → 编辑 XML → 重新打包）
- 用户要求验证 .docx 文件结构（XSD 模式验证）
- 用户要求对文档进行追踪修订（redlining）
- 用户要求在文档中添加批注（comments）
- 用户要求将 .docx 转为 PDF 进行视觉验证

## 前置条件

- 用于中间脚本的工作目录（临时文件夹）
- 用于最终交付物的输出目录（用户的工作区文件夹）

## 环境配置

本工具需要**两个运行时**。引擎 A（Python）始终需要；引擎 B（Node.js）仅在文档包含 OMML 数学公式时需要。

### 引擎 A：Python + python-docx（始终需要）

**1. 安装 Python 3.8+**

从 https://www.python.org/downloads/ 下载并运行安装程序。Windows 安装时勾选 "Add Python to PATH"。

验证：
```bash
python --version
# 预期：Python 3.8.x 或更高
```

**2. 安装 python-docx**

```bash
pip install python-docx
```

验证：
```bash
python -c "import docx; print('python-docx version:', docx.__version__)"
# 预期：python-docx version: 1.1.0 或类似
```

> **注意**：`python-docx` 是 PyPI 包名。导入方式为 `from docx import Document`，而非 `import python-docx`。

### 引擎 B：Node.js + docx（数学公式时需要）

**1. 安装 Node.js 18+**

从 https://nodejs.org/ 下载（推荐 LTS 版本）并运行安装程序。

验证：
```bash
node --version
# 预期：v18.x.x 或更高
npm --version
# 预期：9.x.x 或更高
```

**2. 安装 docx 库（v8.5.0+）**

在脚本执行的工作目录中：
```bash
npm install docx
```

验证：
```bash
node -e "const d = require('docx'); console.log('docx loaded, Math available:', !!d.Math)"
# 预期：docx loaded, Math available: true
```

> **关键**：Math API（`Math`、`MathSubScript`、`MathFraction` 等）需要 `docx` 包 **8.5.0 或更高版本**。更低版本不导出这些类。如果 `d.Math` 为 `undefined`，请升级：`npm install docx@latest`。

**3.（可选）安装 mathHelpers.js 依赖**

`mathHelpers.js` 和 `formulas.js` 仅依赖 `docx` 包，无需额外 npm 包。

### 环境快速检查

运行以下组合检查验证两个引擎是否就绪：

```bash
# 检查 Python 引擎
python -c "import docx; print('[OK] python-docx:', docx.__version__)"

# 检查 Node.js 引擎（不需要数学公式时可跳过）
node -e "const d=require('docx'); console.log('[OK] docx Math:', !!d.Math, '| MathSubScript:', !!d.MathSubScript)"
```

如果两个都输出 `[OK]`，则环境完全就绪。

### 常见安装问题

| 问题 | 原因 | 解决方案 |
|-------|-------|----------|
| `ModuleNotFoundError: No module named 'docx'` | python-docx 未安装 | `pip install python-docx` |
| `Cannot find module 'docx'` | 当前目录未安装 docx npm 包 | `cd <工作目录> && npm install docx` |
| `d.Math is undefined` | docx 版本过低（< 8.5.0） | `npm install docx@latest` |
| `pip` 未识别 | Python 不在 PATH 中 | 重新安装 Python 时勾选 "Add to PATH" |
| `node` 未识别 | Node.js 不在 PATH 中 | 重新安装 Node.js 或手动添加到 PATH |
| 中文字体显示为方框 | 系统缺少宋体/SimHei | 安装东亚字体包；Windows 通常默认自带 |
| `ET.fromstring()` 失败 | 嵌套数组导致 XML 损坏 | 确保所有数学辅助函数中应用了 `flat()`（模板已处理） |

## 文件结构

```
.trae/skills/docx-formatter-plus/
├── SKILL.md                          # 本文件
└── scripts/
    ├── build_docx.js                 # Node.js 文档构建模板（含 OMML 数学公式）
    ├── build_docx.py                 # Python 文档构建模板（文本 + 图表）
    ├── mathHelpers.js                # OMML 数学元素构建器（Node.js）
    ├── formulas.js                   # 公式定义模板（Node.js）
    ├── merge_runs.py                 # 合并碎片 run（编辑现有文档前置步骤）
    ├── accept_changes.py             # 接受所有追踪修订（清理 redlining）
    ├── comment.py                    # 批注管理（6 文件交叉链接系统）
    ├── templates/                    # 批注 XML 模板
    │   ├── comments.xml
    │   ├── commentsExtended.xml
    │   ├── commentsExtensible.xml
    │   ├── commentsIds.xml
    │   └── people.xml
    └── office/                       # 验证与转换工具
        ├── __init__.py
        ├── soffice.py                # LibreOffice 跨平台调用（Linux LD_PRELOAD shim + Windows 原生）
        ├── validate.py               # XSD 模式验证入口
        ├── helpers/                  # 通用辅助函数包
        │   ├── __init__.py           # safe_extract, rezip, opc_target 等
        │   ├── pptx_chart.py         # PPTX 图表验证辅助
        │   ├── pptx_slide.py         # PPTX 幻灯片验证辅助
        │   └── pptx_theme.py         # PPTX 主题验证辅助
        ├── schemas/                  # OOXML XSD 模式文件
        │   ├── ISO-IEC29500-4_2016/  # 核心模式（wml, dml, shared 等）
        │   ├── ecma/fouth-edition/   # OPC 模式（content types, relationships）
        │   ├── microsoft/            # 微软扩展模式（wml-2010/2012/2018 等）
        │   └── mce/                  # Markup Compatibility 模式
        └── validators/               # 验证器
            ├── __init__.py
            ├── base.py               # 基类：XML/XSD/命名空间/ID/引用验证
            ├── docx.py               # DOCX 专用验证（空白保留、删除/插入、批注标记）
            ├── pptx.py               # PPTX 验证器（附带）
            └── redlining.py          # 追踪修订验证（检测未追踪的编辑）
```

## 架构：双引擎

### 引擎 A：Python（python-docx）
**适用场景：** 不需要 OMML 数学公式，仅需文本、表格和表格框图。

**需复制到工作目录的脚本文件：**
- `scripts/build_docx.py` — 包含所有辅助函数的主模板

**核心函数（build_docx.py）：**
| 函数 | 用途 |
|----------|---------|
| `setup_document()` | 创建文档，2.54cm 页边距，宋体默认字体 |
| `add_title(doc, text)` | 黑体 三号(16pt) 居中加粗，1.5倍行距，段前=12pt，段后=12pt |
| `add_h1(doc, text)` | 黑体 三号(16pt) 左对齐加粗，1.5倍行距，段前=10pt，段后=24pt |
| `add_h2(doc, text)` | 黑体 小四(12pt) 左对齐加粗，1.5倍行距，段前=18pt，段后=12pt |
| `add_h3(doc, text)` | 黑体 11pt 左对齐加粗，1.5倍行距，段前=14pt，段后=8pt |
| `add_body(doc, text)` | 宋体 小四(12pt) 两端对齐，首行缩进2字符，1.5倍行距。支持 `**bold**` 和 `[n]` 引用上标 |
| `add_item_para(doc, label, text)` | 加粗标签 + 正文，首行缩进。正文支持 `**bold**` 和 `[n]` 引用上标 |
| `add_box(doc, text, width_cm, font_size)` | 单个居中框，用于流程图（默认宽度14cm，含单元格边距） |
| `add_multi_line_box(doc, lines, width_cm, font_size)` | 多行居中框，用于流程图（一个框内多行文字） |
| `add_multi_col_table(doc, cells, col_width_cm, font_size)` | 并排框行（默认总宽14cm，含单元格边距） |
| `add_arrow_down(doc)` | 居中 ↓ 箭头（16pt），连接上下框 |
| `add_arrow_row(doc, left_text, right_text, total_cm, arrow_cm, font_size, left_shade, right_shade)` | 横向"过程 → 产出"行：3列表格（左\|箭头\|右），左右等宽，固定布局，箭头恒居中。融合中间竖线，可选底纹（默认全白，左灰右白传 `left_shade='D9D9D9'`） |
| `add_separator_note(doc, text)` | 居中虚线分隔注释（如 '----- 前处理止于此处 -----') |
| `add_fig_caption(doc, text)` | 图标题，位于图下方（自动："图N 描述"） |
| `add_table_caption(doc, text)` | 表标题，位于表上方（自动："表N 描述"） |
| `reset_counters()` | 重置图表计数器为零（由 `setup_document()` 自动调用） |
| `add_note(doc, text)` | 居中斜体注释（9pt），位于图下方 |
| `add_toc(doc, title, levels)` | **按需**目录（黑体三号居中标题 + TOC域，levels '1-2'） |
| `set_table_border(table)` | 设置所有边框为单线黑色 |
| `set_cell_shading(cell, color_hex)` | 设置单元格背景色（如 'D9D9D9' 灰色） |

### 引擎 B：Node.js（docx 库）
**适用场景：** 需要 OMML 数学公式（下标、上标、分式、求和符号）。

**需复制到工作目录的脚本文件：**
- `scripts/build_docx.js` — 主模板
- `scripts/mathHelpers.js` — OMML 数学元素构建器（build_docx.js 依赖）
- `scripts/formulas.js` — 公式定义（build_docx.js 依赖）

**核心函数（build_docx.js）：**
| 函数 | 用途 |
|----------|---------|
| `title(text)` | 文档主标题段落（居中，加粗，1.5倍行距，段前=12pt，段后=12pt） |
| `h1(text)` | H1 标题段落（左对齐，加粗，1.5倍行距，段前=10pt，段后=24pt） |
| `h2(text)` | H2 标题段落（左对齐，加粗，1.5倍行距，段前=18pt，段后=12pt） |
| `para(text)` | 正文段落（1.5倍行距）。支持 `**bold**` 和 `[n]` 引用上标 |
| `itemPara(label, text)` | 加粗标签 + 正文（返回 Paragraph）。正文支持 `**bold**` 和 `[n]` 引用上标 |
| `eqPara(mathObj)` | 居中块级公式段落 |
| `box(text, opts)` | 单个居中表格框 |
| `multiLineBox(lines, opts)` | 多行居中表格框（一个框内多行文字） |
| `twoColBox(leftTitle, leftLines, rightTitle, rightLines)` | 并排对比框 |
| `multiCellRow(cells, opts)` | 水平排列的并行单元格行 |
| `arrowDown()` | ↓ 箭头段落 |
| `arrowRow(leftText, rightText, opts)` | 横向"过程 → 产出"行：3列表格，左右等宽，固定布局，箭头居中。融合中间竖线，可选底纹（默认全白，左灰右白传 `{ leftShade: "D9D9D9" }`） |
| `separatorNote(text)` | 居中虚线分隔注释（如 '----- 前处理止于此处 -----') |
| `figCaption(text)` | 图标题，位于图下方（自动："图N 描述"），空字符串仅生成间距 |
| `tableCaption(text)` | 表标题，位于表上方（自动："表N 描述"） |
| `resetCounters()` | 重置图表计数器为零（在 CONTENT 区段自动调用） |
| `note(text)` | 斜体注释段落，位于图下方 |
| `pageBreak()` | 分页符段落 |
| `toc(title, levels)` | **按需**目录（返回数组：标题段落 + TOC + 分页符） |

## 文档结构与函数映射

生成文档时，请根据结构元素选择正确的函数：

| 场景 | 示例 | Python 函数 | Node.js 函数 | 格式 | 进入目录 |
|----------|---------|-----------------|------------------|--------|-----|
| 文档主标题 | `技术交底书` | `add_title(doc, text)` | `title(text)` | 黑体 16pt, 居中, 加粗 | 否 |
| 一级章节（一、二、） | `一、发明名称` | `add_h1(doc, text)` | `h1(text)` | 黑体 16pt, 左对齐, 加粗 | **是 (Heading 1)** |
| 二级编号标题 | `3.1 现有技术...` | `add_h2(doc, text)` | `h2(text)` | 黑体 12pt, 左对齐, 加粗 | **是 (Heading 2)** |
| 二级步骤标题 | `步骤1：构建...` | `add_h2(doc, text)` | `h2(text)` | 黑体 12pt, 左对齐, 加粗 | **是 (Heading 2)** |
| 二级括号标题 | `（一）系统总体架构` | `add_h2(doc, text)` | `h2(text)` | 黑体 12pt, 左对齐, 加粗 | **是 (Heading 2)** |
| 三级标题 | `6.1 xxx方法` | `add_h3(doc, text)` | `h2(text)` | 黑体 11pt, 左对齐, 加粗 | **是 (Heading 3)** |
| **编号分点（加粗标签）** | `（1）定义核心概念...：内容` | `add_item_para(doc, "（1）定义...：", "内容")` | `itemPara("（1）定义...：", "内容")` | 宋体 12pt, 加粗标签 + 正文, 首行缩进 |
| **字母子步骤（加粗标签）** | `a. 任务分解：内容` | `add_item_para(doc, "a. 任务分解：", "内容")` | `itemPara("a. 任务分解：", "内容")` | 宋体 12pt, 加粗标签 + 正文, 首行缩进 |
| 普通正文段落 | 技术描述段落 | `add_body(doc, text)` | `para(text)` | 宋体 12pt, 两端对齐, 缩进2字符 |
| 图标题 | `系统架构图` → 自动 "图1 系统架构图" | `add_fig_caption(doc, text)` | `figCaption(text)` | 宋体 10.5pt, 居中, 图下方 |
| 表标题 | `参数对比` → 自动 "表1 参数对比" | `add_table_caption(doc, text)` | `tableCaption(text)` | 宋体 10.5pt, 居中, 表上方 |
| 图注（斜体） | `（核心：...）` | `add_note(doc, text)` | `note(text)` | 宋体 9pt, 居中, 斜体 |

### 关键规则
1. **（1）/（2）/（3）分点**：label 部分加粗（含冒号），正文部分正常，使用 `add_item_para` / `itemPara`，首行缩进2字符
2. **a. / b. / c. 子步骤**：同上格式，label 加粗（含冒号），使用 `add_item_para` / `itemPara`
3. **步骤1：/ 步骤2：**：作为二级标题处理，使用 `add_h2` / `h2`，黑体小四加粗左对齐
4. **（一）/（二）**：作为二级标题处理，使用 `add_h2` / `h2`
5. **3.1 / 3.2**：作为二级标题处理，使用 `add_h2` / `h2`
6. **正文中的 `**bold**`**：`add_body` / `para` 自动解析加粗标记
7. **正文中的 `[n]` 引用**：`add_body` / `para` 自动渲染为上标
8. **Word 目录（TOC）支持**：`add_h1`/`add_h2`/`add_h3`（Python）和 `h1`/`h2`（Node.js）使用 Word 内置 Heading 样式。使用 `add_toc`/`toc` 按需插入目录域（TOC field），在 Word 中右键"更新域"生成目录条目。`add_title`/`title` 不使用 Heading 样式，不进入目录
9. **标题颜色为黑色**：Word 内置 Heading 样式默认为蓝色，所有标题函数已显式覆盖为黑色（`RGBColor(0,0,0)` / `color: "000000"`），确保标题显示为黑色而非蓝色
10. **图表自动编号**：`add_fig_caption`/`figCaption` 和 `add_table_caption`/`tableCaption` 自动递增编号（图N / 表N），无需手动填写编号。传空字符串 `""` 给 `figCaption` 可生成纯间距段落（不编号）。图标题在图**下方**，表标题在表**上方**

## 行距与标题间距

所有间距值对齐 **NJUThesis** LaTeX 模板（`linespread = 1.625`，等效 Word 1.5倍行距）。

### 间距参数

| 元素 | 行距 | 段前 | 段后 | 说明 |
|---------|-------------|-------------|-------------|-------|
| 文档标题 | 1.5倍 | 12pt | 12pt | 居中，黑体 16pt |
| H1（章） | 1.5倍 | 10pt | 24pt | 左对齐，黑体 16pt。NJUThesis 章 after=60pt，Word 中缩减为24pt |
| H2（节） | 1.5倍 | 18pt | 12pt | 左对齐，黑体 12pt。匹配 ctex section 默认值 |
| H3（小节） | 1.5倍 | 14pt | 8pt | 左对齐，黑体 11pt。匹配 ctex subsection 默认值 |
| 正文段落 | 1.5倍 | 0 | 7pt | 两端对齐，宋体 12pt，首行缩进2字符 |
| 分点段落 | 1.5倍 | 0 | 7pt | 加粗标签 + 正文，首行缩进 |
| 块级公式 | 1.5倍 | 6pt | 6pt | 居中 |
| 目录条目 | 固定22pt | 0 | 0 | TOC 1/2 样式，独立于正文间距 |

### NJUThesis 行距背景

NJUThesis 全局使用 `linespread = 1.625`。计算方式：LaTeX 默认行距倍数为 1.2，Word 默认为 1.3，要求 1.5倍 Word 行距，故 `1.5 × (1.3 / 1.2) = 1.625`。在 Word 中设置 1.5倍行距即可达到相同视觉效果。

### 实现说明

- **Python**：`pf.line_spacing = 1.5` 设置 Word 1.5倍行距。`pf.space_before` / `pf.space_after` 控制段落间距。
- **Node.js**：`spacing: { line: 360 }` 设置 1.5倍行距（360/240 = 1.5）。`before` / `after` 单位为 twips（1pt = 20 twips）。
- 标题前的空段落（`doc.add_paragraph()`）已移除，间距改由 `space_before` / `space_after` 精确控制，结构更干净。

## 图表自动编号

图表由内部计数器**自动编号**。无需手动写"图1"或"表1"，只需传入描述文字。

### 格式

| 元素 | 自动格式 | 位置 | 字体 |
|---------|-------------|----------|------|
| 图标题 | `图N 描述文字` | 图下方 | 宋体 10.5pt, 居中 |
| 表标题 | `表N 描述文字` | 表上方 | 宋体 10.5pt, 居中 |

### 用法

**Python:**
```python
# 图：编号自动递增（图1, 图2, ...）
add_box(doc, "流程步骤1", width_cm=10)
add_fig_caption(doc, "系统架构流程图")    # → "图1 系统架构流程图"

# 表：编号自动递增（表1, 表2, ...）
add_table_caption(doc, "参数对比表")       # → "表1 参数对比表"
add_multi_col_table(doc, ["方法A", "方法B"])

# 重置计数器（仅在单进程多次生成文档时需要）
# reset_counters()  # 通常由 setup_document() 自动调用
```

**Node.js:**
```javascript
// 图
children.push(box("流程步骤1", { width: 6000 }));
children.push(figCaption("系统架构流程图"));   // → "图1 系统架构流程图"

// 表
children.push(tableCaption("参数对比表"));     // → "表1 参数对比表"
children.push(multiCellRow([{title: "方法A"}, {title: "方法B"}]));

// 重置计数器（仅在单进程多次生成文档时需要）
// resetCounters();  // 通常在 CONTENT 区段自动调用
```

### 关键规则
- 图表编号相互独立（图1, 图2... 和 表1, 表2...）
- `figCaption("")` — 空字符串生成间距段落，**不**递增计数器
- `tableCaption(text)` 始终递增（无空字符串跳过机制）
- 计数器自动重置：`setup_document()`（Python）和 CONTENT 区段（Node.js）会自动调用 `reset_counters()` / `resetCounters()`，每次生成文档从 图1/表1 开始。仅在单进程多次生成文档且未重新调用 setup 时需手动重置

## 横向箭头行（过程 → 产出）

`add_arrow_row` / `arrowRow` 创建 3 列表格（左框 | → | 右框），左右等宽并锁定 `tblLayout=fixed`，使箭头始终位于整表几何正中心，与文字长短无关。**中间竖线融合**（insideV=none），左右单元格视觉上连通。

### 底纹规则

| 场景 | 参数 | 效果 |
|------|------|------|
| 左右地位相等 | 不传 `left_shade` / `right_shade`（默认 None） | 左右全白 |
| 左右层级不一 | `left_shade='D9D9D9'`，`right_shade` 不传 | 左灰右白 |
| 自定义 | 传入任意 hex 色值 | 对应单元格着色 |

### 适用场景

- "过程 → 产出"步骤（如 `① 数据采集 → 原始数据集`）
- 需要用底纹区分层级的流程图（如数据层 → 模型层，左灰右白）
- 任何需要横向箭头连接不同长度文字框的流程图

### 用法

以下示例展示两种底纹模式：前两步地位相等（全白），后两步层级转换（左灰右白）。

**Python:**
```python
# 前两步：同属数据层，地位相等 → 全白
add_arrow_row(doc, "① 数据采集", "原始数据集")
add_arrow_down(doc)
add_arrow_row(doc, "② 数据标注", "标注数据集")
add_arrow_down(doc)

# 后两步：从数据层进入模型层，层级不一 → 左灰右白
add_arrow_row(doc, "③ 模型训练", "预测模型", left_shade='D9D9D9')
add_arrow_down(doc)
add_arrow_row(doc, "④ 模型部署", "推理服务", left_shade='D9D9D9')
add_arrow_down(doc)

add_separator_note(doc, "----- 离线训练止于此处，后续由在线服务承接 -----")
add_arrow_down(doc)
add_box(doc, "API 网关 + 业务系统：请求路由 / 负载均衡 / 结果返回 / 监控告警")
add_fig_caption(doc, "模型训练与部署流程图")
```

**Node.js:**
```javascript
// 前两步：同属数据层，地位相等 → 全白
children.push(arrowRow("① 数据采集", "原始数据集"));
children.push(arrowDown());
children.push(arrowRow("② 数据标注", "标注数据集"));
children.push(arrowDown());

// 后两步：从数据层进入模型层，层级不一 → 左灰右白
children.push(arrowRow("③ 模型训练", "预测模型", { leftShade: "D9D9D9" }));
children.push(arrowDown());
children.push(arrowRow("④ 模型部署", "推理服务", { leftShade: "D9D9D9" }));
children.push(arrowDown());

children.push(separatorNote("----- 离线训练止于此处，后续由在线服务承接 -----"));
children.push(arrowDown());
children.push(box("API 网关 + 业务系统：请求路由 / 负载均衡 / 结果返回 / 监控告警"));
children.push(figCaption("模型训练与部署流程图"));
```

### 实现原理

1. **3 列表格**：左框 | 箭头列 | 右框
2. **左右等宽**：`side = (total - arrow) / 2`，布局常数，不依赖内容
3. **固定布局**（`tblLayout=fixed`）：防止 Word autofit 重排窄箭头列
4. **融合中间竖线**（`insideV=none`）：去掉左/箭头/右之间的竖线，外框 + insideH 保留
5. **可选底纹**：`left_shade` / `right_shade` 默认 None（全白），层级不同时左灰右白

## 目录（TOC）— 按需生成

目录**默认不生成**。仅在文档需要目录时调用 `add_toc()`（Python）或展开 `toc()`（Node.js）。

### 格式规范

对齐 NJUThesis LaTeX 模板。

| 元素 | 字体 | 字号 | 样式 |
|---------|------|------|-------|
| 目录标题 | 黑体 (SimHei) | 16pt (三号) | 加粗，居中 |
| 一级条目 | 黑体 (SimHei) | 14pt (四号) | 常规 |
| 二级条目 | 宋体 (SimSun) | 12pt (小四) | 常规 |
| 目录行距 | — | 固定22pt | — |

### 工作原理

1. **标题样式**：`add_h1`/`h1` 和 `add_h2`/`h2` 使用 Word 内置 Heading 1/2 样式，TOC 域可识别
2. **目录条目样式**：TOC 1 和 TOC 2 样式按上述格式预定义：
   - **Python**：`_setup_toc_styles(doc)` 在 `add_toc()` 内部调用，配置 TOC 1/TOC 2 样式
   - **Node.js**：TOC1/TOC2 段落样式在 Document 构造函数的 `styles.paragraphStyles` 中定义
3. **域生成**：目录是 Word 域（`TOC \o "1-2" \h \z \u`）。在 Word 中打开后，右键目录区域 → "更新域"生成条目
4. **层级**：默认 `'1-2'` 仅包含 H1 和 H2 标题

### 用法

**Python:**
```python
doc = setup_document()
add_title(doc, "文档标题")
add_toc(doc)                    # 在标题后、正文前插入目录
add_h1(doc, "一、概述")
add_body(doc, "正文内容...")
```

**Node.js:**
```javascript
children.push(title("文档标题"));
children.push(...toc());        // 将目录元素展开到 children 数组
children.push(h1("一、概述"));
children.push(para("正文内容..."));
```

### 注意事项
- 目录条目在用户于 Word 中更新域之前不会显示（右键 → "更新域"）
- 更新前显示占位文字"（请在 Word 中右键此处选择"更新域"以生成目录）"（仅 Python；Node.js 使用库默认值）
- `add_title`/`title` 不使用 Heading 样式，文档主标题不会出现在目录中
- 目录后跟分页符，与正文分隔

## 排版标准

### 字体与字号
| 元素 | 字体 | 字号 | 样式 |
|---------|------|------|-------|
| 文档标题 | 黑体 (SimHei) | 16pt (三号) | 加粗，居中 |
| H1（一级标题） | 黑体 (SimHei) | 16pt (三号) | 加粗，左对齐 |
| H2（二级标题） | 黑体 (SimHei) | 12pt (小四) | 加粗，左对齐 |
| H3（三级标题） | 黑体 (SimHei) | 11pt | 加粗，左对齐 |
| 正文 | 宋体 (SimSun) | 12pt (小四) | 常规，两端对齐 |
| 目录标题 | 黑体 (SimHei) | 16pt (三号) | 加粗，居中 |
| 一级条目 | 黑体 (SimHei) | 14pt (四号) | 常规 |
| 二级条目 | 宋体 (SimSun) | 12pt (小四) | 常规 |
| 表格单元格 | 宋体 (SimSun) | 9-10.5pt | 常规 |
| 表格框图 | 宋体 (SimSun) | 10-14pt | 标题加粗 |

**注意：** docx 库中字号单位为半磅：12pt = 24，16pt = 32，10.5pt = 21。

### 页面设置
- 纸张大小：A4（11907 × 16840 twips）
- 页边距：四周 2.54cm（1440 twips）
- 行距：1.5倍（docx 库中为 360）
- 首行缩进：24pt = 2个中文字符（docx 库中为 480）
- 表格对齐：居中

## 引用上标（双引擎）

正文中的参考文献引用标记自动渲染为上标。适用于 `add_body()` / `para()` 和 `add_item_para()` / `itemPara()` 函数。

### 支持的引用格式
| 格式 | 示例 | 渲染效果 |
|--------|---------|------------|
| 单个 | `[1]` | <sup>[1]</sup> |
| 多个 | `[1,2]` 或 `[1, 2]` | <sup>[1,2]</sup> |
| 范围 | `[1-3]` | <sup>[1-3]</sup> |

### 工作原理

`add_body`（Python）和 `para`（Node.js）函数使用组合正则解析正文，同时检测 `**bold**` 片段和 `[n]` 引用标记。引用标记渲染为上标 `TextRun` 对象。

**Python（python-docx）：**
```python
# 引用上标通过以下方式设置：
run.font.superscript = True
```

**Node.js（docx 库）：**
```javascript
// 引用上标通过以下方式设置：
new TextRun({ text: "[1]", font: FONT, size: 21, superScript: true })
```

### 用法示例

直接在正文中包含 `[n]` 标记即可，无需特殊语法：

```python
# Python
add_body(doc, "如文献[1]所述，该方法在工业场景中表现出色。后续研究[2,3]进一步验证了这一结论。")
```
```javascript
// Node.js
para("如文献[1]所述，该方法在工业场景中表现出色。后续研究[2,3]进一步验证了这一结论。")
```

两者均生成 `[1]` 和 `[2,3]` 渲染为上标字符的段落。

### 内部辅助函数
- **Python**：`_add_runs_with_formatting(p, text)` — 被 `add_body()` 和 `add_item_para()` 共用，解析 `**bold**` 和 `[n]` 引用
- **Node.js**：`parseRichText(text)` — 返回 `TextRun` 对象数组，被 `para()` 和 `itemPara()` 使用

## OMML 数学公式指南（仅引擎 B）

### 基本元素
```javascript
const { r, sub, sup, frac, sumOp, func, math, inlineMath } = require("./mathHelpers");

// 纯数学文本
r("E_total")

// 下标：E_loss
sub("E", "loss")

// 上标：x²
sup("x", "2")

// 分式：a/b
frac([r("a")], [r("b")])

// 求和：Σ_i（不要空上标！）
sumOp([r("i")], [sub("y", "i")])  // 返回数组，自动展平

// 函数：log(p_i)
func("log", [sub("p", "i")])

// 块级公式（居中，独立行）
math([sub("E", "loss"), r(" = -"), sumOp([r("i")], [sub("y","i"), r(" log("), sub("p","i"), r(")")])])

// 行内公式（与文本混排在段落中）
inlineMath([sub("y", "i")])
```

### 关键：求和符号
**问题**：`MathSum` 使用空 `superScript: []` 会在 Word 中渲染一个不可见的上标框，导致文件看起来损坏或显示异常。

**解决方案**：改用 `MathSubScript`，以 `∑`（U+2211）作为基础字符。

```javascript
// 错误 — 渲染空上标，可能导致文件损坏
new MathSum({ subScript: [r("i")], superScript: [], children: [...] })

// 正确 — 无上标，干净渲染
new MathSubScript({ children: [r("∑")], subScript: [r("i")] })
```

此问题已在 `mathHelpers.js` → `sumOp()` 中处理。切勿直接使用 `MathSum`。

### 关键：数组展平
`sumOp()` 返回数组 `[sumSymbol, ...body]`。如果不展平，嵌套数组会导致 XML 损坏，.docx 文件无法打开。

`flat()` 已在所有接收 children 的函数中应用：`math()`、`frac()`、`func()`、`paren()`、`bracket()`、`inlineMath()`。

### 正文中的行内公式

在单个 Paragraph 中将 `TextRun`（中文文本）与 `inlineMath()`（数学）混合使用：

```javascript
new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  indent: { firstLine: 480 },
  spacing: { after: 140, line: 360 },
  children: [
    new TextRun({ text: "其中，", font: "SimSun", size: 21 }),
    inlineMath([sub("E", "loss")]),
    new TextRun({ text: "为损失函数，计算公式为：", font: "SimSun", size: 21 }),
  ],
});
```

### 希腊字母与特殊字符

在 `r()` 调用中直接使用 Unicode：
| 字符 | Unicode | 用途 |
|-----------|---------|-------|
| λ (lambda) | `\u03bb` | 系数 / 权重参数 |
| τ (tau) | `\u03c4` | 温度参数 |
| α (alpha) | `\u03b1` | 权重系数 |
| β (beta) | `\u03b2` | 权重系数 |
| Σ (sigma) | `\u2211` | 求和（在 sumOp 中使用） |
| − (减号) | `\u2212` | 数学减号 |
| × (乘号) | `\u00d7` | 乘法 |
| · (点号) | `\u00b7` | 点积 |
| ∈ (属于) | `\u2208` | 集合隶属 |
| ↓ (下箭头) | `\u2193` | 流程图连接符 |
| → (右箭头) | `\u2192` | 横向流程 |

## 表格框图模式

在 Word 中绘制流程图/图表（不使用 emoji 或图片）：

```
[框 A]
   ↓
[多行框 B]
  (第1行)
  (第2行)
   ↓
[左框]  [右框]
   ↓
[框 C]
图1 示意图
(注：斜体注释)
```

**实现模式：**
1. `figCaption("")` — 图前空间距段落
2. `box("text")` — 单个框
3. `arrowDown()` — 连接符
4. `multiLineBox(["line1", "line2"])` — 多行框（Python: `add_multi_line_box(doc, ["line1", "line2"])`）
5. `twoColBox(...)` — 并排对比
6. `multiCellRow([...])` — 并行项行
7. `figCaption("图N 描述")` — 图后标题
8. `note("annotation text")` — 图下斜体注释（Python: `add_note(doc, "annotation text")`）

## 工作流程

1. **阅读**用户提供的源文本/内容
2. **确定引擎**：
   - 内容包含数学符号（Σ、下标、分式、方程）→ 引擎 B（Node.js）
   - 仅文本、表格、图表 → 引擎 A（Python）
3. **复制模板脚本**从 `.trae/skills/docx-formatter-plus/scripts/` 到工作目录
4. **修改**模板的 CONTENT 区段为实际文档内容
5. **运行脚本** — 输出到用户工作区文件夹
6. **验证** — 检查 ZIP 完整性和 XML 可解析性：
   ```python
   import zipfile, xml.etree.ElementTree as ET
   z = zipfile.ZipFile(path)
   bad = z.testzip()
   assert bad is None, f"Corrupt: {bad}"
   ET.fromstring(z.read('word/document.xml'))
   print(f"OK, size={len(z.read('word/document.xml'))} bytes")
   ```
7. **分享**文件，使用 `computer://` 链接

## 常见陷阱

1. **禁止使用 emoji** — 文档中仅使用 Unicode 符号
2. **文件锁定**：如果 Word 已打开文件，保存会静默失败 — 使用不同文件名
3. **docx 版本**：Math API 需要 `docx` npm 包 v8.5.0+
4. **禁止使用 MathSum**：始终使用 mathHelpers.js 中的 `sumOp()` — 它使用 `MathSubScript` 配合 Unicode ∑，避免空上标渲染问题
5. **始终展平数组**：所有数学辅助函数使用 `flat()` — 新增函数时需对所有 children 参数应用 `flat()`
6. **输出路径**：Node.js 脚本中始终使用 Windows 反斜杠路径（`C:\\Users\\...`）
7. **字体东亚设置**：python-docx 中必须设置 `rFonts.set(qn('w:eastAsia'), font_name)` 才能正确渲染中文字体

## 文件验证清单

交付 .docx 文件前：
- [ ] ZIP 完整性：`zipfile.testzip()` 返回 None
- [ ] XML 有效：`ET.fromstring(document.xml)` 成功
- [ ] 文件大小 > 10KB（空文档约 3KB）
- [ ] 内容中无 emoji 字符
- [ ] 所有数学符号渲染为 OMML（非纯文本）— 仅引擎 B
- [ ] 图表使用表格框模式（非图片）
- [ ] 表格有边框和标题底纹（D9D9D9，灰色）
- [ ] 所有公式使用 `math()` / `inlineMath()`，非原始 `TextRun` 加下标字符

## 可选依赖

核心文档生成功能（引擎 A/B）不依赖以下组件。仅在需要验证、编辑、批注等增强功能时安装。

| 依赖 | 安装方式 | 用途 | 级别 |
|------|---------|------|------|
| `defusedxml` | `pip install defusedxml` | 安全 XML 解析（防 XXE），编辑/验证/批注使用 | 推荐 |
| `lxml` | `pip install lxml` | XSD 模式验证引擎 | 可选（XSD 验证时需要） |
| LibreOffice | 系统安装 | PDF 渲染验证、接受追踪修订 | 可选（PDF 验证/redlining 时需要） |
| Poppler (`pdftoppm`) | 系统安装 | PDF 转图片做视觉检查 | 可选（PDF 视觉验证时需要） |
| pandoc | 系统安装 | 读取 .docx 内容为 Markdown | 可选（读取现有文档时需要） |

### 可选依赖安装

```bash
# 推荐：安全 XML 解析（编辑/验证/批注都需要）
pip install defusedxml lxml

# 可选：PDF 渲染验证
# Windows: 从 https://www.libreoffice.org/ 下载安装
# Linux: sudo apt install libreoffice poppler-utils
# 验证: soffice --version

# 可选：读取现有 docx 内容
# Windows: 从 https://pandoc.org/ 下载安装
# Linux: sudo apt install pandoc
```

## docx-js 完整陷阱清单

除前述"常见陷阱"外，以下为 docx-js (Node.js) 的全部已知陷阱：

| 陷阱 | 说明 | 解决方案 |
|------|------|---------|
| 页面默认 A4 | docx-js 默认 A4 | US Letter 需设 `page: { size: { width: 12240, height: 15840 } }`（DXA; 1440=1英寸） |
| 横向页面 | 需要交换宽高 | 传入纵向尺寸 + `orientation: PageOrientation.LANDSCAPE`，库内部交换 |
| 表格双宽度 | PERCENTAGE 在 Google Docs 中失效 | `columnWidths` 和每个 cell 的 `width` 都要设，用 `WidthType.DXA`，宽度之和 = 表宽 |
| 表格底纹 | SOLID 渲染为黑色 | 用 `ShadingType.CLEAR`，不用 `SOLID` |
| 列表 | 不要直接插入 `•` | 用 `numbering` 配置 + `LevelFormat.BULLET` |
| ImageRun | 缺少 type 字段会报错 | 必须指定 `type:` (`"png"`, `"jpg"` 等) |
| PageBreak | 不能直接放 | 必须放在 `Paragraph` 内部 |
| 换行 | 不要用 `\n` | 用独立的 `Paragraph` 元素 |
| TOC 标题 | 自定义样式不进目录 | 标题必须用内置 `HeadingLevel.*`；自定义样式需设 `outlineLevel` |
| 水平线 | 不要用表格做 | 用段落底边框 |
| 点引导符 | 不要用字面量点号 | 用 `PositionalTab`（`alignment: PositionalTabAlignment.RIGHT`, `leader: PositionalTabLeader.DOT`） |

## 安全处理规范

处理来自外部方的 .docx 文件时，必须遵循以下安全措施：

### 1. 安全解压（防路径穿越）

`office/helpers.py` 中的 `safe_extract()` 函数在解压时执行两项检查：
- **拒绝符号链接**：`stat.S_ISLNK(m.external_attr >> 16)` 检测并拒绝
- **路径不逃逸**：`(dest / m.filename).resolve()` 确保解压目标在目标目录内

```python
from office.helpers import safe_extract
import zipfile

with zipfile.ZipFile(untrusted_docx, "r") as zf:
    safe_extract(zf, Path("unpacked/"))  # 安全解压
```

### 2. 安全 XML 解析（防 XXE）

所有 XML 解析应使用 `defusedxml` 替代标准库，防止 XML External Entity (XXE) 攻击：

```python
# 错误 — 标准 XML 库易受 XXE 攻击
import xml.etree.ElementTree as ET
tree = ET.parse("word/document.xml")

# 正确 — defusedxml 阻止 XXE
import defusedxml.ElementTree as ET
tree = ET.parse("word/document.xml")
```

### 3. 清除符号链接

处理外部文档前，解压后立即清除符号链接条目：

```bash
unzip -q document.docx -d unpacked/
# Linux: find unpacked -type l -delete
# PowerShell: Get-ChildItem unpacked -Recurse -Link | Remove-Item
```

## PDF 渲染验证（可选）

生成 .docx 后，可通过 LibreOffice 转 PDF 并生成图片进行视觉检查，确保排版正确。

### 工作流程

```bash
# 1. docx 转 PDF
python scripts/office/soffice.py --headless --convert-to pdf output.docx

# 2. PDF 转图片（需要 Poppler 的 pdftoppm）
pdftoppm -jpeg -r 100 output.pdf page

# 3. 检查生成的图片
ls page-*.jpg   # page-01.jpg, page-02.jpg, ...
```

> `pdftoppm` 会将页码零填充到总页数宽度（如 12 页 → `page-01.jpg`…`page-12.jpg`）。

### 跨平台说明

`soffice.py` 同时支持 Linux 和 Windows：
- **Linux 沙箱**：自动检测 AF_UNIX socket 限制，编译 LD_PRELOAD shim 绕过
- **Windows**：自动搜索 `C:\Program Files\LibreOffice\program\soffice.exe` 等常见路径
- **通用**：每次调用使用临时用户配置目录，避免与运行中的 LibreOffice 实例冲突

```python
from office.soffice import run_soffice

# 编程方式调用
result = run_soffice(["--headless", "--convert-to", "pdf", "output.docx"])
```

## XSD 模式验证（可选）

验证生成的 .docx 文件是否符合 OOXML (ISO/IEC 29500-4) 标准。

### 用法

```bash
# 基本验证（检查所有 XSD 错误）
python scripts/office/validate.py output.docx

# 与原始文件对比验证（仅报告新引入的错误）
python scripts/office/validate.py output.docx --original original.docx

# 自动修复常见问题
python scripts/office/validate.py output.docx --auto-repair

# 追踪修订验证（检查所有编辑是否正确追踪）
python scripts/office/validate.py output.docx --original original.docx --author "你的名字"
```

### 验证项目（11 项）

| 验证项 | 说明 |
|-------|------|
| XML 良构性 | 所有 XML 文件可解析 |
| 命名空间声明 | Ignorable 引用的命名空间必须声明 |
| ID 唯一性 | comment/bookmark ID 不重复 |
| 文件引用完整性 | .rels 引用的文件必须存在，无孤立文件 |
| Content-Type 声明 | 所有部件在 `[Content_Types].xml` 中声明 |
| 空白保留 | 带首尾空白的 `w:t` 必须有 `xml:space="preserve"` |
| 删除验证 | `w:del` 内不能有 `w:t`（应为 `w:delText`） |
| 插入验证 | `w:ins` 内不能有 `w:delText` |
| 关系ID引用 | `r:id` 等引用必须指向已定义的 Relationship |
| ID 约束 | paraId < 0x80000000, durableId < 0x7FFFFFFF |
| 批注标记配对 | commentRangeStart/End 必须配对 |

### 自动修复

`--auto-repair` 可修复两类常见问题：
- **paraId/durableId 溢出**：超过 OOXML 限制的 ID 值
- **缺失 xml:space="preserve"**：带空白的 `w:t` 元素未设置保留属性

## 编辑现有文档

docx-js 无法打开现有 .docx 文件。编辑现有文档需通过 XML 操作：解包 → 编辑 → 重新打包。

### 编辑流程

```bash
# 1. 旧版 .doc 先转换
python scripts/office/soffice.py --headless --convert-to docx file.doc

# 2. 解包（安全解压，防路径穿越）
python -c "
import zipfile
from pathlib import Path
from office.helpers import safe_extract
with zipfile.ZipFile('file.docx', 'r') as zf:
    safe_extract(zf, Path('unpacked/'))
"

# 3. 合并碎片 run（关键步骤）
python scripts/merge_runs.py unpacked/

# 4. 编辑 unpacked/word/document.xml（不要重新格式化或美化 XML）

# 5. 重新打包
python -c "
from pathlib import Path
from office.helpers import rezip
rezip(Path('unpacked/'), Path('output.docx'))
"

# 6. 验证
python scripts/office/validate.py output.docx --original file.docx
```

### merge_runs.py — 合并碎片 run

Word 会将段落文字碎片化到多个 `w:r` run 中（修订标记、拼写检查、编辑历史），导致查找替换不可靠。`merge_runs.py` 合并相邻同格式 run，不影响渲染。

```bash
# 解包目录模式
python scripts/merge_runs.py unpacked/

# 直接处理 .docx 文件
python scripts/merge_runs.py document.docx

# 输出到新文件
python scripts/merge_runs.py document.docx -o merged.docx
```

### 追踪修订（Tracked Changes / Redlining）

编辑文档时使用追踪修订，所有改动需包裹在 `<w:ins>`/`<w:del>` 标记中：

- **`<w:ins>`**：插入内容，需设 `w:id`, `w:author`, `w:date` 属性
- **`<w:del>`**：删除内容，文本元素用 `<w:delText>` 而非 `<w:t>`
- **删除段落标记**：`<w:pPr><w:rPr><w:del w:id=".." w:author=".." w:date=".."/></w:rPr></w:pPr>` 表示"合并到下一段"
- **`<w:del/>` 位置**：必须在 rPr 的其他子元素之前（schema 强制顺序）

验证追踪修订：
```bash
# --author 检查所有编辑是否被正确追踪
python scripts/office/validate.py output.docx --original original.docx --author "你的名字"
```

接受所有追踪修订（生成干净副本）：
```bash
python scripts/accept_changes.py input.docx output.docx
```

> **注意**：`accept_changes.py` 使用 LibreOffice 宏接受修订。删除段落标记应将该段落合并到下一段；全删除的段落应消失。`pandoc --track-changes=accept` 不会合并段落，`accept_changes.py` 在段落后跟空段时也可能不合并。空项目符号是查看工具的产物，不是文档缺陷。

### 批注（Comments）

批注需要 6 个交叉链接的 XML 文件。使用 `comment.py` 辅助工具：

```bash
# 对解包目录添加批注（推荐：同时编辑 document.xml 时使用）
python scripts/comment.py unpacked/ "费用上限过低"

# 添加回复（批注 id=0 的子批注）
python scripts/comment.py unpacked/ "同意" --parent 0

# 对 .docx 文件直接添加批注
python scripts/comment.py contract.docx "此上限过低" -o annotated.docx
```

脚本自动创建 `comments.xml`、`commentsExtended.xml`、`commentsIds.xml`、`commentsExtensible.xml`、关系文件和内容类型覆盖。批注 ID 自动分配。

添加批注后，需在 `word/document.xml` 中插入标记使批注可见：
```xml
<w:commentRangeStart w:id="0"/>
... 被批注的内容 ...
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

### 读取文档内容

使用 pandoc 将 .docx 转为 Markdown：
```bash
pandoc -t markdown file.docx
```
