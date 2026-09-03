***

name: docx-formatter
description: 生成专业排版的中文 Word (.docx) 文档，支持 OMML 数学公式、表格框图、标准排版、引用上标和轻量完整性验证。当用户要求生成带公式的 Word 文档、用数学方程格式化文档、在 Word 中创建流程图样式图表、将参考文献引用渲染为上标、或排版/格式化文档为 Word 时调用。
-------------------------------------------------------------------------------------------------------------------------------------------------------------

# 中文 DOCX 排版工具

本工具用于生成专业排版的 Word (.docx) 文档。核心功能：

1. **标准排版**（黑体标题/宋体正文、标准页边距、行距）
2. **OMML 数学公式**（原生 Word 公式：行内公式和块级公式，直接生成 OMML XML，无中间库依赖）
3. **表格框图**（使用 Word 表格 + Unicode 箭头绘制流程图，不使用 emoji 或图片）
4. **引用上标**（正文中的 \[1] 等引用标记自动渲染为上标）
5. **轻量验证**（纯 stdlib 的 ZIP/XML 完整性检查，生成后默认执行）

## 可选功能 — 默认跳过

以下功能**默认不执行、不检查环境、不安装依赖**。仅当用户**明确要求**时才启用，详见文末「可选功能详解」章节：

| 可选功能      | 触发条件（仅用户明确要求）             | 所在脚本                                                                        |
| --------- | ------------------------- | --------------------------------------------------------------------------- |
| PDF 渲染验证  | 用户要求转 PDF 做视觉检查           | `scripts/optional/office/soffice.py`                                        |
| XSD 模式验证  | 用户要求 OOXML 标准深度验证         | `scripts/optional/office/validate.py`                                       |
| 编辑现有 docx | 用户要求修改已有 .docx 文件         | `scripts/optional/merge_runs.py` + `build_docx.py` 内 `safe_extract`/`rezip` |
| 追踪修订 / 批注 | 用户要求 redlining / comments | `scripts/optional/accept_changes.py`、`comment.py`                           |

> **规则**：默认工作流只有「生成 → 轻量验证 → 交付」。遇到可选功能请求时，先参考对应章节确认依赖（LibreOffice、lxml、pandoc 等均按需安装），再执行。

## 调用时机

- 用户要求生成带数学公式的 Word 文档（.docx）

- 用户要求用方程、下标、分式、求和符号格式化文档

- 用户要求在 Word 中创建流程图样式图表（使用表格 + 箭头）

- 用户要求将文档中的参考文献引用渲染为上标

- 用户要求"排版"或"格式化"文档为 Word

- 用户要求"使用与之前相同的排版"生成 Word 文档

- 用户明确要求编辑现有 docx、追踪修订、批注、PDF/XSD 验证 → 调用本 skill 并启用对应可选功能

## 前置条件

- 用于中间脚本的工作目录（临时文件夹）

- 用于最终交付物的输出目录（用户的工作区文件夹）

## 环境配置

本工具仅需**一个运行时**：Python 3.8+ 与 python-docx。数学公式由 `mathHelpers.py` 直接生成 OMML XML，无需 Node.js 或其他任何外部依赖。

### 安装 Python 与 python-docx

**1. 安装 Python 3.8+**

从 <https://www.python.org/downloads/> 下载并运行安装程序。Windows 安装时勾选 "Add Python to PATH"。

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

### 环境快速检查

```bash
python -c "import docx; print('[OK] python-docx:', docx.__version__)"
```

输出 `[OK]` 即环境就绪。

### 常见安装问题

| 问题                                            | 原因               | 解决方案                          |
| --------------------------------------------- | ---------------- | ----------------------------- |
| `ModuleNotFoundError: No module named 'docx'` | python-docx 未安装  | `pip install python-docx`     |
| `pip` 未识别                                     | Python 不在 PATH 中 | 重新安装 Python 时勾选 "Add to PATH" |
| 中文字体显示为方框                                     | 系统缺少宋体/SimHei    | 安装东亚字体包；Windows 通常默认自带        |
| `parse_xml()` 失败                              | OMML 字符串格式错误     | 确认使用 mathHelpers.py 的构建函数     |

## 文件结构

```
skill 根目录（SKILL.md 所在目录）/
├── SKILL.md                     # 本文件
├── README.md                    # 项目说明
├── LICENSE
└── scripts/
    ├── build_docx.py            # Python 文档构建模板（核心；含 safe_extract/rezip 供可选编辑功能使用）
    ├── mathHelpers.py           # OMML 数学元素构建器（核心，纯 stdlib）
    ├── formulas.py              # 公式定义模板（核心）
    ├── validate_docx.py         # 轻量验证脚本（核心，纯 stdlib，默认执行）
    └── optional/                # ★ 可选功能脚本 — 默认不使用、不复制到工作目录
        ├── merge_runs.py        # 合并碎片 run（编辑现有文档前置步骤）
        ├── accept_changes.py    # 接受所有追踪修订（LibreOffice 宏）
        ├── comment.py           # 批注管理（6 文件交叉链接系统）
        ├── templates/           # 批注 XML 模板
        │   ├── comments.xml
        │   ├── commentsExtended.xml
        │   ├── commentsExtensible.xml
        │   ├── commentsIds.xml
        │   └── people.xml
        └── office/              # 验证与转换工具（可选）
            ├── soffice.py       # LibreOffice 跨平台调用
            ├── validate.py      # XSD 模式验证入口
            ├── helpers/         # 通用辅助函数（safe_extract, rezip, opc_target 等）
            ├── schemas/         # OOXML XSD 模式文件
            └── validators/      # 验证器（docx / redlining / pptx）
```

## 核心函数（build\_docx.py）

**需复制到工作目录的脚本文件：**

- `scripts/build_docx.py` — 包含所有排版辅助函数的主模板

- `scripts/mathHelpers.py` — OMML 数学构建器（仅文档含公式时需要）

- `scripts/formulas.py` — 公式定义示例（仅文档含公式时需要）

| 函数                                                                                                  | 用途                                                                                        |
| --------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| `setup_document()`                                                                                  | 创建文档，2.54cm 页边距，宋体默认字体                                                                    |
| `add_title(doc, text)`                                                                              | 黑体 三号(16pt) 居中加粗，1.5倍行距，段前=12pt，段后=12pt                                                   |
| `add_h1(doc, text)`                                                                                 | 黑体 三号(16pt) 左对齐加粗，1.5倍行距，段前=10pt，段后=24pt                                                  |
| `add_h2(doc, text)`                                                                                 | 黑体 小四(12pt) 左对齐加粗，1.5倍行距，段前=18pt，段后=12pt                                                  |
| `add_h3(doc, text)`                                                                                 | 黑体 11pt 左对齐加粗，1.5倍行距，段前=14pt，段后=8pt                                                       |
| `add_body(doc, text)`                                                                               | 宋体 小四(12pt) 两端对齐，首行缩进2字符，1.5倍行距。支持 `**bold**` 和 `[n]` 引用上标                                |
| `add_item_para(doc, label, text)`                                                                   | 加粗标签 + 正文，首行缩进。正文支持 `**bold**` 和 `[n]` 引用上标                                               |
| `add_eq_para(doc, math_xml)`                                                                        | 居中块级公式段落（接收 mathHelpers 的 OMML XML 字符串）                                                   |
| `add_body_with_math(doc, parts)`                                                                    | 正文与行内公式混排段落（parts 为 `("text", str)` / `("math", xml)` 列表）                                 |
| `add_code_block(doc, code, font_size)`                                                              | 代码块：Consolas 等宽字体（默认9pt），浅灰底纹（F2F2F2），逐行段落，左缩进                                       |
| `add_data_table(doc, headers, rows, col_widths, font_size)`                                         | 数据表：灰色表头（D9D9D9 黑体加粗居中），数据行宋体，末列左对齐其余居中，固定列宽                                  |
| `add_math_to_cell(cell, omml_xml)`                                                                  | 向表格单元格插入行内 OMML 公式（居中），配合 `inlineMath()` 用于公式对照表等场景                            |
| `add_box(doc, text, width_cm, font_size)`                                                           | 单个居中框，用于流程图（默认宽度14cm，含单元格边距）                                                              |
| `add_multi_line_box(doc, lines, width_cm, font_size)`                                               | 多行居中框，用于流程图（一个框内多行文字）                                                                     |
| `add_multi_col_table(doc, cells, col_width_cm, font_size)`                                          | 并排框行（默认总宽14cm，含单元格边距）                                                                     |
| `add_arrow_down(doc)`                                                                               | 居中 ↓ 箭头（16pt），连接上下框                                                                       |
| `add_arrow_row(doc, left_text, right_text, total_cm, arrow_cm, font_size, left_shade, right_shade)` | 横向"过程 → 产出"行：3列表格（左\|箭头\|右），左右等宽，固定布局，箭头恒居中。融合中间竖线，可选底纹（默认全白，左灰右白传 `left_shade='D9D9D9'`） |
| `add_separator_note(doc, text)`                                                                     | 居中虚线分隔注释（如 '----- 前处理止于此处 -----')                                                         |
| `add_fig_caption(doc, text)`                                                                        | 图标题，位于图下方（自动："图N 描述"）                                                                     |
| `add_table_caption(doc, text)`                                                                      | 表标题，位于表上方（自动："表N 描述"）                                                                     |
| `reset_counters()`                                                                                  | 重置图表计数器为零（由 `setup_document()` 自动调用）                                                      |
| `add_note(doc, text)`                                                                               | 居中斜体注释（9pt），位于图下方                                                                         |
| `add_toc(doc, title, levels)`                                                                       | **按需**目录（黑体三号居中标题 + TOC域，levels '1-2'）                                                    |
| `set_table_border(table)`                                                                           | 设置所有边框为单线黑色                                                                               |
| `set_cell_shading(cell, color_hex)`                                                                 | 设置单元格背景色（如 'D9D9D9' 灰色）                                                                   |
| `safe_extract(zf, dest)`                                                                            | 安全解压 .docx ZIP 包，防止路径遍历和符号链接攻击（可选功能用）                                                     |
| `rezip(src_dir, out_path)`                                                                          | 将目录重新打包为 .docx 文件，确保 `[Content_Types].xml` 首位存储（可选功能用）                                    |

## 文档结构与函数映射

生成文档时，请根据结构元素选择正确的函数：

| 场景              | 示例                      | Python 函数                               | 格式                       | 进入目录              |
| --------------- | ----------------------- | --------------------------------------- | ------------------------ | ----------------- |
| 文档主标题           | `技术交底书`                 | `add_title(doc, text)`                  | 黑体 16pt, 居中, 加粗          | 否                 |
| 一级章节（一、二、）      | `一、发明名称`                | `add_h1(doc, text)`                     | 黑体 16pt, 左对齐, 加粗         | **是 (Heading 1)** |
| 二级编号标题          | `3.1 现有技术...`           | `add_h2(doc, text)`                     | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 二级步骤标题          | `步骤1：构建...`             | `add_h2(doc, text)`                     | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 二级括号标题          | `（一）系统总体架构`             | `add_h2(doc, text)`                     | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 三级标题            | `6.1 xxx方法`             | `add_h3(doc, text)`                     | 黑体 11pt, 左对齐, 加粗         | **是 (Heading 3)** |
| **编号分点（加粗标签）**  | `（1）定义核心概念...：内容`       | `add_item_para(doc, "（1）定义...：", "内容")` | 宋体 12pt, 加粗标签 + 正文, 首行缩进 | <br />            |
| **字母子步骤（加粗标签）** | `a. 任务分解：内容`            | `add_item_para(doc, "a. 任务分解：", "内容")`  | 宋体 12pt, 加粗标签 + 正文, 首行缩进 | <br />            |
| 普通正文段落          | 技术描述段落                  | `add_body(doc, text)`                   | 宋体 12pt, 两端对齐, 缩进2字符     | <br />            |
| 块级数学公式          | `L = Σ ...`             | `add_eq_para(doc, math([...]))`         | 居中，1.5倍行距                | <br />            |
| 行内公式混排          | `其中 L_LLM 为...`         | `add_body_with_math(doc, parts)`        | 宋体 12pt, 两端对齐            | <br />            |
| 代码块             | 命令/源码片段                 | `add_code_block(doc, code)`             | Consolas 9pt, 灰底, 无首行缩进  | <br />            |
| 数据表             | 参数表 / 检查项表 / 功能矩阵    | `add_data_table(doc, headers, rows, col_widths)` | 灰色表头, 固定列宽        | <br />            |
| 表格单元格内公式       | 公式对照表第三列              | `add_math_to_cell(cell, inlineMath([...]))` | OMML 居中              | <br />            |
| 图标题             | `系统架构图` → 自动 "图1 系统架构图" | `add_fig_caption(doc, text)`            | 宋体 10.5pt, 居中, 图下方       | <br />            |
| 表标题             | `参数对比` → 自动 "表1 参数对比"   | `add_table_caption(doc, text)`          | 宋体 10.5pt, 居中, 表上方       | <br />            |
| 图注（斜体）          | `（核心：...）`              | `add_note(doc, text)`                   | 宋体 9pt, 居中, 斜体           | <br />            |

### 关键规则

1. **（1）/（2）/（3）分点**：label 部分加粗（含冒号），正文部分正常，使用 `add_item_para`，首行缩进2字符
2. **a. / b. / c. 子步骤**：同上格式，label 加粗（含冒号），使用 `add_item_para`
3. **步骤1：/ 步骤2：**：作为二级标题处理，使用 `add_h2`，黑体小四加粗左对齐
4. **（一）/（二）**：作为二级标题处理，使用 `add_h2`
5. **3.1 / 3.2**：作为二级标题处理，使用 `add_h2`
6. **正文中的** **`**bold**`**：`add_body` 自动解析加粗标记
7. **正文中的** **`[n]`** **引用**：`add_body` 自动渲染为上标
8. **Word 目录（TOC）支持**：`add_h1`/`add_h2`/`add_h3` 使用 Word 内置 Heading 样式。使用 `add_toc` 按需插入目录域（TOC field），在 Word 中右键"更新域"生成目录条目。`add_title` 不使用 Heading 样式，不进入目录
9. **标题颜色为黑色**：Word 内置 Heading 样式默认为蓝色，所有标题函数已显式覆盖为黑色（`RGBColor(0,0,0)`），确保标题显示为黑色而非蓝色
10. **图表自动编号**：`add_fig_caption` 和 `add_table_caption` 自动递增编号（图N / 表N），无需手动填写编号。传空字符串 `""` 给 `add_fig_caption` 可生成纯间距段落（不编号）。图标题在图**下方**，表标题在表**上方**

## 行距与标题间距

所有间距值对齐 **NJUThesis** LaTeX 模板（`linespread = 1.625`，等效 Word 1.5倍行距）。

### 间距参数

| 元素     | 行距     | 段前   | 段后   | 说明                                               |
| ------ | ------ | ---- | ---- | ------------------------------------------------ |
| 文档标题   | 1.5倍   | 12pt | 12pt | 居中，黑体 16pt                                       |
| H1（章）  | 1.5倍   | 10pt | 24pt | 左对齐，黑体 16pt。NJUThesis 章 after=60pt，Word 中缩减为24pt |
| H2（节）  | 1.5倍   | 18pt | 12pt | 左对齐，黑体 12pt。匹配 ctex section 默认值                  |
| H3（小节） | 1.5倍   | 14pt | 8pt  | 左对齐，黑体 11pt。匹配 ctex subsection 默认值               |
| 正文段落   | 1.5倍   | 0    | 7pt  | 两端对齐，宋体 12pt，首行缩进2字符                             |
| 分点段落   | 1.5倍   | 0    | 7pt  | 加粗标签 + 正文，首行缩进                                   |
| 块级公式   | 1.5倍   | 6pt  | 6pt  | 居中                                               |
| 目录条目   | 固定22pt | 0    | 0    | TOC 1/2 样式，独立于正文间距                               |

### NJUThesis 行距背景

NJUThesis 全局使用 `linespread = 1.625`。计算方式：LaTeX 默认行距倍数为 1.2，Word 默认为 1.3，要求 1.5倍 Word 行距，故 `1.5 × (1.3 / 1.2) = 1.625`。在 Word 中设置 1.5倍行距即可达到相同视觉效果。

### 实现说明

- `pf.line_spacing = 1.5` 设置 Word 1.5倍行距。`pf.space_before` / `pf.space_after` 控制段落间距。

- 标题前的空段落（`doc.add_paragraph()`）已移除，间距改由 `space_before` / `space_after` 精确控制，结构更干净。

## 图表自动编号

图表由内部计数器**自动编号**。无需手动写"图1"或"表1"，只需传入描述文字。

### 格式

| 元素  | 自动格式      | 位置  | 字体            |
| --- | --------- | --- | ------------- |
| 图标题 | `图N 描述文字` | 图下方 | 宋体 10.5pt, 居中 |
| 表标题 | `表N 描述文字` | 表上方 | 宋体 10.5pt, 居中 |

### 用法

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

### 关键规则

- 图表编号相互独立（图1, 图2... 和 表1, 表2...）

- `add_fig_caption("")` — 空字符串生成间距段落，**不**递增计数器

- `add_table_caption(text)` 始终递增（无空字符串跳过机制）

- 计数器自动重置：`setup_document()` 自动调用 `reset_counters()`，每次生成文档从 图1/表1 开始。仅在单进程多次生成文档且未重新调用 setup 时需手动重置

## 横向箭头行（过程 → 产出）

`add_arrow_row` 创建 3 列表格（左框 | → | 右框），左右等宽并锁定 `tblLayout=fixed`，使箭头始终位于整表几何正中心，与文字长短无关。**中间竖线融合**（insideV=none），左右单元格视觉上连通。

### 底纹规则

| 场景     | 参数                                       | 效果      |
| ------ | ---------------------------------------- | ------- |
| 左右地位相等 | 不传 `left_shade` / `right_shade`（默认 None） | 左右全白    |
| 左右层级不一 | `left_shade='D9D9D9'`，`right_shade` 不传   | 左灰右白    |
| 自定义    | 传入任意 hex 色值                              | 对应单元格着色 |

### 适用场景

- "过程 → 产出"步骤（如 `① 数据采集 → 原始数据集`）

- 需要用底纹区分层级的流程图（如数据层 → 模型层，左灰右白）

- 任何需要横向箭头连接不同长度文字框的流程图

### 用法

以下示例展示两种底纹模式：前两步地位相等（全白），后两步层级转换（左灰右白）。

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

### 实现原理

1. **3 列表格**：左框 | 箭头列 | 右框
2. **左右等宽**：`side = (total - arrow) / 2`，布局常数，不依赖内容
3. **固定布局**（`tblLayout=fixed`）：防止 Word autofit 重排窄箭头列
4. **融合中间竖线**（`insideV=none`）：去掉左/箭头/右之间的竖线，外框 + insideH 保留
5. **可选底纹**：`left_shade` / `right_shade` 默认 None（全白），层级不同时左灰右白

## 目录（TOC）— 按需生成

目录**默认不生成**。仅在文档需要目录时调用 `add_toc()`。

### 格式规范

对齐 NJUThesis LaTeX 模板。

| 元素   | 字体          | 字号        | 样式    |
| ---- | ----------- | --------- | ----- |
| 目录标题 | 黑体 (SimHei) | 16pt (三号) | 加粗，居中 |
| 一级条目 | 黑体 (SimHei) | 14pt (四号) | 常规    |
| 二级条目 | 宋体 (SimSun) | 12pt (小四) | 常规    |
| 目录行距 | —           | 固定22pt    | —     |

### 工作原理

1. **标题样式**：`add_h1` 和 `add_h2` 使用 Word 内置 Heading 1/2 样式，TOC 域可识别
2. **目录条目样式**：TOC 1 和 TOC 2 样式按上述格式预定义：`_setup_toc_styles(doc)` 在 `add_toc()` 内部调用，配置 TOC 1/TOC 2 样式
3. **域生成**：目录是 Word 域（`TOC \o "1-2" \h \z \u`）。在 Word 中打开后，右键目录区域 → "更新域"生成条目
4. **层级**：默认 `'1-2'` 仅包含 H1 和 H2 标题

### 用法

```python
doc = setup_document()
add_title(doc, "文档标题")
add_toc(doc)                    # 在标题后、正文前插入目录
add_h1(doc, "一、概述")
add_body(doc, "正文内容...")
```

### 注意事项

- 目录条目在用户于 Word 中更新域之前不会显示（右键 → "更新域"）

- 更新前显示占位文字"（请在 Word 中右键此处选择"更新域"以生成目录）"

- `add_title` 不使用 Heading 样式，文档主标题不会出现在目录中

- 目录后跟分页符，与正文分隔

## 排版标准

### 字体与字号

| 元素       | 字体          | 字号        | 样式      |
| -------- | ----------- | --------- | ------- |
| 文档标题     | 黑体 (SimHei) | 16pt (三号) | 加粗，居中   |
| H1（一级标题） | 黑体 (SimHei) | 16pt (三号) | 加粗，左对齐  |
| H2（二级标题） | 黑体 (SimHei) | 12pt (小四) | 加粗，左对齐  |
| H3（三级标题） | 黑体 (SimHei) | 11pt      | 加粗，左对齐  |
| 正文       | 宋体 (SimSun) | 12pt (小四) | 常规，两端对齐 |
| 目录标题     | 黑体 (SimHei) | 16pt (三号) | 加粗，居中   |
| 一级条目     | 黑体 (SimHei) | 14pt (四号) | 常规      |
| 二级条目     | 宋体 (SimSun) | 12pt (小四) | 常规      |
| 表格单元格    | 宋体 (SimSun) | 9-10.5pt  | 常规      |
| 表格框图     | 宋体 (SimSun) | 10-14pt   | 标题加粗    |

### 页面设置

- 纸张大小：A4（11907 × 16840 twips）

- 页边距：四周 2.54cm（1440 twips）

- 行距：1.5倍

- 首行缩进：24pt = 2个中文字符

- 表格对齐：居中

## 引用上标

正文中的参考文献引用标记自动渲染为上标。适用于 `add_body()` 和 `add_item_para()` 函数。

### 支持的引用格式

| 格式 | 示例                 | 渲染效果              |
| -- | ------------------ | ----------------- |
| 单个 | `[1]`              | <sup>\[1]</sup>   |
| 多个 | `[1,2]` 或 `[1, 2]` | <sup>\[1,2]</sup> |
| 范围 | `[1-3]`            | <sup>\[1-3]</sup> |

### 工作原理

`add_body` 函数使用组合正则解析正文，同时检测 `**bold**` 片段和 `[n]` 引用标记。引用标记渲染为上标 run：

```python
# 引用上标通过以下方式设置：
run.font.superscript = True
```

### 用法示例

直接在正文中包含 `[n]` 标记即可，无需特殊语法：

```python
add_body(doc, "如文献[1]所述，该方法在工业场景中表现出色。后续研究[2,3]进一步验证了这一结论。")
```

### 内部辅助函数

- `_add_runs_with_formatting(p, text)` — 被 `add_body()` 和 `add_item_para()` 共用，解析 `**bold**` 和 `[n]` 引用

## OMML 数学公式指南

数学公式由 `mathHelpers.py` 直接构建 OMML（Office Math Markup Language）XML，经 `build_docx.py` 的 `add_eq_para()` / `add_body_with_math()` 插入文档。生成的是 Word 原生可编辑公式，与手动插入的公式完全一致。

### 基本元素

```python
from mathHelpers import r, sub, sup, frac, sumOp, func, paren, bracket, math, inlineMath

# 纯数学文本
r("E_total")

# 下标：E_loss
sub("E", "loss")

# 上标：x²
sup("x", "2")

# 分式：a/b
frac([r("a")], [r("b")])

# 求和：Σ_i（不要空上标！）
sumOp([r("i")], [sub("y", "i")])  # 返回列表，自动展平

# 函数：log(p_i)
func("log", [sub("p", "i")])

# 圆括号：(a+b)
paren([r("a+b")])

# 方括号：[a+b]
bracket([r("a+b")])

# 块级公式（居中，独立行）— 传给 add_eq_para()
math([sub("E", "loss"), r(" = -"), sumOp([r("i")], [sub("y","i"), r(" log("), sub("p","i"), r(")")])])

# 行内公式（与文本混排）— 传给 add_body_with_math()
inlineMath([sub("y", "i")])
```

### 插入公式

**块级公式（居中独立行）：**

```python
from build_docx import add_eq_para
from mathHelpers import r, sub, sumOp, func, math

eq = math([
    sub("L", "LLM"), r(" = - "),
    sumOp([r("i")], [sub("y", "i"), func("log", [sub("p", "i")])]),
])
add_eq_para(doc, eq)
```

**行内公式（与正文混排）：**

```python
from build_docx import add_body_with_math
from mathHelpers import sub, inlineMath

add_body_with_math(doc, [
    ("text", "其中，"),
    ("math", inlineMath([sub("L", "LLM")])),
    ("text", "为语言模型损失项，其计算涉及交叉熵。"),
])
```

### 预定义公式（formulas.py）

`formulas.py` 提供 eq1\~eq4 常用公式模板（总损失、交叉熵、对比学习 InfoNCE、LoRA 分解），可直接使用或作为新增公式的参考模式：

```python
from formulas import eq1, eq2, eq3, eq4
add_eq_para(doc, eq2)   # 交叉熵损失
```

### 关键：求和符号

**问题**：n 元运算符（`m:nary`）使用空上标会在 Word 中渲染一个不可见的上标框，导致文件看起来损坏或显示异常。

**解决方案**：`sumOp()` 已用下标结构（`m:sSub`）配合 Unicode `∑`（U+2211）实现，无空上标问题。**始终使用** **`sumOp()`，切勿手写** **`m:nary`** **加空上标**。

### 关键：列表展平

`sumOp()` 返回列表 `[sumSymbol, *body]`。所有接收 children 的函数（`math()`、`frac()`、`func()`、`paren()`、`bracket()`、`inlineMath()`）内部已调用 `_flat()` 自动展平嵌套列表。

### 关键：XML 转义

`r()` 自动对文本进行 XML 转义（`&`、`<`、`>`），公式中可安全使用这些字符。

### 希腊字母与特殊字符

在 `r()` 调用中直接使用 Unicode：

| 字符         | Unicode  | 用途              |
| ---------- | -------- | --------------- |
| λ (lambda) | `\u03bb` | 系数 / 权重参数       |
| τ (tau)    | `\u03c4` | 温度参数            |
| α (alpha)  | `\u03b1` | 权重系数            |
| β (beta)   | `\u03b2` | 权重系数            |
| Σ (sigma)  | `\u2211` | 求和（在 sumOp 中使用） |
| − (减号)     | `\u2212` | 数学减号            |
| × (乘号)     | `\u00d7` | 乘法              |
| · (点号)     | `\u00b7` | 点积              |
| ∈ (属于)     | `\u2208` | 集合隶属            |
| ↓ (下箭头)    | `\u2193` | 流程图连接符          |
| → (右箭头)    | `\u2192` | 横向流程            |

## 代码块 / 数据表 / 单元格公式

### 代码块（add_code_block）

命令行、源码片段等代码内容使用等宽字体 + 浅灰底纹渲染，**不要**将代码放入普通正文段落：

```python
add_code_block(doc, "python validate_docx.py output.docx --verbose")

add_code_block(doc, """from build_docx import setup_document
doc = setup_document()""")
```

### 数据表（add_data_table）

参数表、检查项表、功能矩阵等**数据型表格**使用 `add_data_table`（灰色表头、固定列宽），配合 `add_table_caption` 自动编号；流程图框图仍用 `add_box` / `add_arrow_row` 系列：

```python
add_table_caption(doc, "排版标准")
add_data_table(doc,
    ["元素", "字体", "字号"],
    [
        ("文档标题", "黑体", "16pt"),
        ("正文", "宋体", "12pt"),
    ],
    col_widths=[2.2, 1.8, 1.8], font_size=9.5)
```

注意：`col_widths` 总宽度建议 ≤ 14.5cm（A4 减去页边距后的版心宽度）；数据行末列左对齐（描述列），其余居中。

### 单元格内公式（add_math_to_cell）

在表格单元格中插入行内 OMML 公式（如"函数调用 vs 渲染效果"对照表）：

```python
from mathHelpers import sub, inlineMath
table = add_data_table(doc, ["元素", "调用", "效果"], rows=[...], col_widths=[2.0, 6.0, 5.0])
add_math_to_cell(table.cell(1, 2), inlineMath([sub("y", "pred")]))
```

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

1. `add_fig_caption(doc, "")` — 图前空间距段落
2. `add_box(doc, "text")` — 单个框
3. `add_arrow_down(doc)` — 连接符
4. `add_multi_line_box(doc, ["line1", "line2"])` — 多行框
5. `add_multi_col_table(doc, [...])` — 并行项行
6. `add_fig_caption(doc, "描述")` — 图后标题
7. `add_note(doc, "annotation text")` — 图下斜体注释

## 工作流程

1. **阅读**用户提供的源文本/内容
2. **复制模板脚本**从本 skill 目录的 `scripts/` 到工作目录：

   - 无公式文档：仅 `build_docx.py`

   - 含公式文档：`build_docx.py` + `mathHelpers.py` + `formulas.py`

   - **不要复制** **`scripts/optional/`**（可选功能专用）
3. **修改**模板的 CONTENT 区段为实际文档内容
4. **运行脚本** — 输出到用户工作区文件夹
5. **验证** — 使用纯 Python 标准库验证工具检查文档完整性：

   ```bash
   python validate_docx.py output.docx
   # 或详细输出模式
   python validate_docx.py output.docx --verbose
   ```

   验证工具执行 5 项检查：ZIP 完整性、XML 格式良好性、文件引用完整性、内容类型声明、空白保留。所有检查通过退出码为 0，否则为 1。
6. **分享**文件，使用 `computer://` 链接

## 常见陷阱

1. **禁止使用 emoji** — 文档中仅使用 Unicode 符号
2. **文件锁定**：如果 Word 已打开文件，保存会静默失败 — 使用不同文件名
3. **禁止手写 m:nary 空上标**：始终使用 mathHelpers.py 中的 `sumOp()` — 它使用 `m:sSub` 配合 Unicode ∑，避免空上标渲染问题
4. **始终展平列表**：所有数学辅助函数使用 `_flat()` — 新增函数时需对所有 children 参数应用 `_flat()`
5. **字体东亚设置**：python-docx 中必须设置 `rFonts.set(qn('w:eastAsia'), font_name)` 才能正确渲染中文字体
6. **表格列宽**：必须同时设置 `table.columns[i].width` 和每个 cell 的 width，否则 Word autofit 会重排列宽
7. **表格底纹**：使用 `WD_FILL_PATTERN.CLEAR` 类型底纹而非其他填充模式，否则部分 Word 版本渲染异常
8. **列表符号**：使用 Word 内置编号样式，切勿在文本中直接写 `•` 字符
9. **脚本级变量**：`_fig_counter` / `_tbl_counter` 在 `setup_document()` 中自动重置。单进程多次生成文档时，只要重新调用 `setup_document()` 即可从 图1/表1 开始

## 文档验证（validate\_docx.py）

`validate_docx.py` 是纯 Python 标准库实现的轻量验证脚本，无外部依赖。用于在生成 .docx 后快速检查文档结构完整性。

### 验证项

| # | 检查项       | 说明                                                           | 常见触发原因                   |
| - | --------- | ------------------------------------------------------------ | ------------------------ |
| 1 | ZIP 完整性   | 所有 ZIP 条目均可解压                                                | 文件写入不完整、磁盘空间不足           |
| 2 | XML 格式良好性 | 所有 .xml/.rels 文件可被 `ElementTree` 解析                          | 标签未闭合、OMML 字符串拼接错误       |
| 3 | 文件引用完整性   | .rels 中每个 Target 在包内存在                                       | 删除部件未同步更新关系文件            |
| 4 | 内容类型声明    | word/ 下的 XML 文件在 `[Content_Types].xml` 中有 Override 或 Default | 手动添加部件未注册内容类型            |
| 5 | 空白保留      | 含首尾空白的 `w:t` 元素必须有 `xml:space="preserve"`                    | python-docx 未设置 space 属性 |

### 用法

```bash
# 基本验证（仅显示 FAILED 项）
python validate_docx.py output.docx

# 详细模式（显示每项 PASSED）
python validate_docx.py output.docx --verbose
```

### 退出码

- `0` — 全部检查通过

- `1` — 一项或多项检查失败

### 编程式调用

```python
from validate_docx import validate_docx

if validate_docx("output.docx", verbose=True):
    print("文档验证通过")
else:
    print("文档存在问题，请检查上方输出")
```

## 安全处理规范

处理来自不可信来源的 .docx 文件时，必须遵循以下安全规范：

### 安全解压（safe\_extract）

`build_docx.py` 中的 `safe_extract()` 函数在解压 .docx ZIP 包时提供两层防护：

1. **符号链接拒绝**：检测 `external_attr` 中的 `S_ISLNK` 标志，拒绝包含符号链接的条目
2. **路径遍历防护**：解析每个条目的目标路径，验证其不超出目标目录边界

```python
from build_docx import safe_extract, rezip
import zipfile

# 安全解压现有 .docx
with zipfile.ZipFile("input.docx", "r") as zf:
    safe_extract(zf, "./unpacked")

# 编辑 XML 后重新打包
rezip("./unpacked", "output.docx")
```

### 安全重新打包（rezip）

`rezip()` 函数确保重新打包的 .docx 符合 OOXML 规范：

1. **`[Content_Types].xml`** **首位存储**：使用 `ZIP_STORED`（无压缩）将该文件放在 ZIP 包首位
2. **原子写入**：先写入临时文件，再通过 `os.replace()` 原子替换目标文件
3. **权限保留**：继承原文件权限（若存在），否则按 umask 创建

### 编辑现有文档工作流

```
解压（safe_extract）→ 合并碎片 run（可选）→ 编辑 XML → 重新打包（rezip）→ 验证（validate_docx）
```

> **注意**：编辑 XML 时保持原始格式，不要 pretty-print。`xml.etree.ElementTree` 的 `tostring()` 默认不保留原始缩进。

## 文件验证清单

交付 .docx 文件前：

- [ ] `validate_docx.py` 全部 5 项检查通过

- [ ] 文件大小 > 10KB（空文档约 3KB）

- [ ] 内容中无 emoji 字符

- [ ] 所有数学符号渲染为 OMML（非纯文本）

- [ ] 图表使用表格框模式（非图片）

- [ ] 表格有边框和标题底纹（D9D9D9，灰色）

- [ ] 所有公式使用 `math()` / `inlineMath()`，非原始文本加下标字符

- [ ] 标题颜色为黑色（非 Word 默认蓝色）

- [ ] 图表编号从 图1/表1 开始（`setup_document()` 自动重置）

***

## 可选功能详解

以下章节**默认不使用**。仅当用户明确要求对应功能时才阅读和执行。所有可选依赖（LibreOffice、lxml、pandoc 等）按需安装，默认不检查。

### 可选功能 A：PDF 渲染验证

**用途**：将生成的 .docx 转换为 PDF 进行视觉检查，确认排版正确。

**依赖**：LibreOffice（`soffice` 命令行）

**脚本**：`scripts/optional/office/soffice.py`（跨平台 LibreOffice 调用封装）

```bash
# 转换为 PDF
python scripts/optional/office/soffice.py convert output.docx --outdir ./pdf_out
```

### 可选功能 B：XSD 模式验证

**用途**：对 .docx 内的 XML 部件执行 OOXML XSD 深度模式验证，比轻量验证更严格。

**依赖**：lxml（`pip install lxml`）

**脚本**：`scripts/optional/office/validate.py`

```bash
python scripts/optional/office/validate.py output.docx
```

### 可选功能 C：编辑现有 docx

**用途**：修改已有 .docx 文件内容（配合 `safe_extract` / `rezip` 使用）。

**前置步骤**：`scripts/optional/merge_runs.py` 合并碎片 run，使编辑更可靠

**工作流**：解压（safe\_extract）→ 合并 run → 编辑 XML → 重打包（rezip）→ 验证

### 可选功能 D：追踪修订 / 批注

**用途**：在文档中插入追踪修订或批注。

**脚本**：

- `scripts/optional/accept_changes.py` — 接受所有追踪修订（LibreOffice 宏）

- `scripts/optional/comment.py` — 批注管理（6 文件交叉链接系统）

- `scripts/optional/templates/` — 批注 XML 模板（comments.xml、commentsExtended.xml、commentsExtensible.xml、commentsIds.xml、people.xml）

