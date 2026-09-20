---

name: docx-formatter
description: 将文本或 Markdown 内容排版为专业的中文 Word (.docx) 文档，支持 OMML 数学公式、表格框图、标准排版、引用上标和轻量完整性验证。当用户要求把 Markdown 或纯文本转成 Word 文档、生成带公式的 Word 文档、用数学方程格式化文档、在 Word 中创建流程图样式图表、将参考文献引用渲染为上标、或排版/格式化文档为 Word 时调用。
---

# 中文 DOCX 排版工具

本工具用于将文本或 Markdown 内容排版为专业的中文 Word (.docx) 文档。最常见的用法是把 AI 生成的 Markdown 正文（报告、说明、方案、论文、讲义等）按学术与工程规范排版为 .docx，同时借助库函数补充 Markdown 无法表达的构件：数学公式、表格框图、引用上标、封面与目录。核心功能：

1. **Markdown 排版**（把 Markdown 或纯文本正文按中文文档规范排版为 .docx）
2. **标准排版**（黑体标题/宋体正文、标准页边距、行距）
3. **OMML 数学公式**（原生 Word 公式：行内公式和块级公式，直接生成 OMML XML，无中间库依赖）
4. **表格框图**（使用 Word 表格 + Unicode 箭头绘制流程图，不使用 emoji 或图片）
5. **引用上标**（正文中的 \[1] 等引用标记自动渲染为上标）
6. **轻量验证**（纯 stdlib 的 ZIP/XML 完整性检查，生成后默认执行）

## 如何开始

本工具为**函数库**，不含可执行入口：`docx_layout_kit.py` 无 `main()` 函数，直接运行不产生任何输出。Agent 生成文档的方式是**编写调用脚本**，导入库函数后按文档结构依次调用。

**Agent 生成文档前，先阅读 `scripts/_build_readme_docx.py`。** 该文件是本工具的标准骨架，从 `setup_document()` 到 `doc.save()` 的完整调用顺序（封面 → 目录 → 标题 → 正文 → 图表 → 参考文献 → 保存）均在其中示范；它同时用于生成项目根目录的 README.docx，因而始终与库保持同步。

**参考文件：**

| 文件 | 用途 |
| -- | -- |
| `scripts/_build_readme_docx.py` | 标准骨架范例：import 段示范应导入的函数，正文示范各函数的调用顺序，文件头附使用说明 |
| `SKILL.md`（本文档） | 各函数的完整参数说明、排版数值、规则与陷阱 |
| `scripts/docx_layout_kit.py` | 函数库本体与 `PRESETS` 预设值，用于确认实现细节 |

**标准流程：**

1. 阅读 `scripts/_build_readme_docx.py`，理解 import 段与正文的调用顺序
2. 复制 `scripts/docx_layout_kit.py` 到工作目录；文档含公式时，再复制 `scripts/omml_math_kit.py` 与 `scripts/formula_templates.py`
3. 新建调用脚本，按范例顺序写入实际内容
4. 运行脚本生成文档，经 `scripts/docx_validator.py` 验证后交付

**顺序规则**（范例文件头有同样提示）：

1. `set_preset()` 须在 `setup_document()` 之前调用
2. `add_cover_page()` 须在所有其它内容之前调用
3. 分节决定页码起点：有目录时 `add_toc()` 已自动分节；有封面但无目录时须手动调用 `start_body(doc)`；两者皆无则无需处理

## 可选功能 — 默认跳过

以下功能**默认不执行、不检查环境、不安装依赖**。仅当用户**明确要求**时才启用，详见文末「可选功能详解」章节：

| 可选功能      | 触发条件（仅用户明确要求）             | 所在脚本                                                                        |
| --------- | ------------------------- | --------------------------------------------------------------------------- |
| PDF 渲染验证  | 用户要求转 PDF 做视觉检查           | `scripts/optional/office/word2pdf.py`（Windows Word）、`scripts/optional/office/soffice.py`（LibreOffice） |
| XSD 模式验证  | 用户要求 OOXML 标准深度验证         | `scripts/optional/office/validate.py`                                       |
| 编辑现有 docx | 用户要求修改已有 .docx 文件         | `scripts/optional/merge_runs.py` + `docx_layout_kit.py` 内 `safe_extract`/`rezip` |
| 追踪修订 / 批注 | 用户要求 redlining / comments | `scripts/optional/accept_changes.py`、`comment.py`                           |

> **规则**：默认工作流只有「生成 → 轻量验证 → 交付」。遇到可选功能请求时，先参考对应章节确认依赖（LibreOffice、lxml、pandoc 等均按需安装），再执行。

## 调用时机

- 用户要求把 Markdown 或纯文本内容排版、转换、导出为 Word 文档（.docx）

- 用户要求把 AI 生成的一段正文（报告、方案、说明、论文、讲义）做成 Word

- 用户要求生成带数学公式的 Word 文档（.docx）

- 用户要求用方程、下标、分式、求和符号格式化文档

- 用户要求在 Word 中创建流程图样式图表（使用表格 + 箭头）

- 用户要求将文档中的参考文献引用渲染为上标

- 用户要求"排版"或"格式化"文档为 Word

- 用户要求页码从正文开始（封面、目录不编号）

- 用户要求制作文档封面（居中大标题 + 副标题 + 单位/日期落款）

- 用户要求"使用与之前相同的排版"生成 Word 文档

- 用户明确要求编辑现有 docx、追踪修订、批注、PDF/XSD 验证 → 调用本 skill 并启用对应可选功能

## 前置条件

- 用于中间脚本的工作目录（临时文件夹）

- 用于最终交付物的输出目录（用户的工作区文件夹）

## 环境配置

本工具仅需**一个运行时**：Python 3.8+ 与 python-docx。数学公式由 `omml_math_kit.py` 直接生成 OMML XML，无需 Node.js 或其他任何外部依赖。

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
| `parse_xml()` 失败                              | OMML 字符串格式错误     | 确认使用 omml_math_kit.py 的构建函数     |

## 文件结构

```
skill 根目录（SKILL.md 所在目录）/
├── SKILL.md                     # 本文件
├── README.md                    # 项目说明
├── README.docx                  # 排版效果样例（由 scripts/_build_readme_docx.py 生成）
├── README.pdf                   # README.docx 的 PDF 渲染结果（由 scripts/optional/office/word2pdf.py 生成）
├── assets/
│   └── effect-overview.png      # README 开头的技能效果示意图
├── LICENSE
└── scripts/
    ├── docx_layout_kit.py     # Python 文档构建库（核心；含 safe_extract/rezip 供可选编辑功能使用）
    ├── omml_math_kit.py       # OMML 数学元素构建器（核心，纯 stdlib）
    ├── formula_templates.py   # 公式定义模板（核心）
    ├── docx_validator.py      # 轻量验证脚本（核心，纯 stdlib，默认执行）
    ├── _build_readme_docx.py  # ★ 标准骨架范例 —— 生成文档前先阅读
    └── optional/              # ★ 可选功能脚本 — 默认不使用、不复制到工作目录
        ├── merge_runs.py      # 合并碎片 run（编辑现有文档前置步骤）
        ├── accept_changes.py  # 接受所有追踪修订（LibreOffice 宏）
        ├── comment.py         # 批注管理（6 文件交叉链接系统）
        ├── templates/         # 批注 XML 模板
        │   ├── comments.xml
        │   ├── commentsExtended.xml
        │   ├── commentsExtensible.xml
        │   ├── commentsIds.xml
        │   └── people.xml
        └── office/            # 验证与转换工具（可选）
            ├── word2pdf.py    # Windows Word（COM）渲染 PDF
            ├── soffice.py     # LibreOffice 跨平台调用
            ├── validate.py    # XSD 模式验证入口
            ├── helpers/       # 通用辅助函数（safe_extract, rezip, opc_target 等）
            ├── schemas/       # OOXML XSD 模式文件
            └── validators/    # 验证器（docx / redlining / pptx）
```

## 核心函数（docx_layout_kit.py）

**需复制到工作目录的脚本文件：**

- `scripts/docx_layout_kit.py` — 包含所有排版辅助函数的构建库（**无 `main()` 入口**，由调用脚本 import 使用）

- `scripts/omml_math_kit.py` — OMML 数学构建器（仅文档含公式时需要）

- `scripts/formula_templates.py` — 公式定义示例（仅文档含公式时需要）

> **调用方式见 `scripts/_build_readme_docx.py`** —— 那份范例展示了下面这些函数在真实文档中的完整串联顺序。不要直接运行 `docx_layout_kit.py`（它没有入口），也不要试图修改它来放入内容；**内容应写在你自己的调用脚本里**。

| 函数                                                                                                  | 用途                                                                                                             |
| --------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| `setup_document()`                                                                                  | 创建文档，上下 2.54cm 左右 3.18cm 页边距，页脚居中页码，宋体默认字体                                                |
| `set_preset(name)`                                                                                  | 整套切换排版预设（标题/正文样式、中西文字体分离、图表标题与图注字号、页边距、目录样式）；**须在** **`setup_document()`** **之前调用**；默认 `set_preset('default')` |
| `add_cover_page(doc, title, subtitle, org, date)`                                                   | **按需**封面：居中标题块（黑体 26pt 大标题 + 20pt 副标题 + 宋体 15pt 单位/日期），用顶部留白定位，留白高度按版心高比例自动换算；封面独立成一节且不显示页码。**必须在写任何其它内容之前调用**（见「封面」一节）            |
| `add_title(doc, text)`                                                                              | 黑体 三号(16pt) 居中加粗，1.5倍行距，段前=12pt，段后=12pt（读预设）                                                                   |
| `add_h1(doc, text)`                                                                                 | 黑体 三号(16pt) 左对齐加粗，1.5倍行距，段前=24pt，段后=6pt                                                                       |
| `add_h2(doc, text)`                                                                                 | 黑体 小四(12pt) 左对齐加粗，1.5倍行距，段前=12pt，段后=6pt                                                                       |
| `add_h3(doc, text)`                                                                                 | 黑体 小四(12pt) 左对齐加粗，1.5倍行距，段前=12pt，段后=6pt                                                                        |
| `add_body(doc, text)`                                                                               | 宋体 小四(12pt) 两端对齐，首行缩进2字符，1.5倍行距。支持 `**bold**` 和 `[n]` 引用上标                                                     |
| `add_item_para(doc, label, text)`                                                                   | 加粗标签 + 正文，首行缩进。正文支持 `**bold**` 和 `[n]` 引用上标                                                                    |
| `add_eq_para(doc, math_xml)`                                                                        | 居中块级公式段落（接收 omml_math_kit 的 OMML XML 字符串）                                                                        |
| `add_body_with_math(doc, parts)`                                                                    | 正文与行内公式混排段落（parts 为 `("text", str)` / `("math", xml)` 列表）                                                      |
| `add_code_block(doc, code, font_size)`                                                              | 代码块：Consolas 等宽字体（默认9pt），浅灰底纹（F2F2F2），逐行段落，左缩进                                                                 |
| `add_data_table(doc, headers, rows, col_widths, font_size)`                                         | 数据表：灰色表头（D9D9D9 黑体加粗居中），数据行宋体，末列左对齐其余居中，固定列宽                                                                   |
| `add_math_to_cell(cell, omml_xml)`                                                                  | 向表格单元格插入行内 OMML 公式（居中），配合 `inlineMath()` 用于公式对照表等场景                                                            |
| `add_box(doc, text, width_cm, font_size)`                                                           | 单个居中框，用于流程图（默认宽度14cm，含单元格边距）                                                                                   |
| `add_multi_line_box(doc, lines, width_cm, font_size)`                                               | 多行居中框，用于流程图（一个框内多行文字）                                                                                          |
| `add_multi_col_table(doc, cells, col_width_cm, font_size)`                                          | 并排框行（默认总宽14cm，含单元格边距）                                                                                          |
| `add_arrow_down(doc)`                                                                               | 居中 ↓ 箭头（16pt），连接上下框                                                                                            |
| `add_arrow_row(doc, left_text, right_text, total_cm, arrow_cm, font_size, left_shade, right_shade)` | 横向"过程 → 产出"行：3列表格（左\|箭头\|右），左右等宽，固定布局，箭头恒居中。融合中间竖线，可选底纹（默认全白，左灰右白传 `left_shade='D9D9D9'`）                      |
| `add_separator_note(doc, text)`                                                                     | 居中虚线分隔注释（如 '----- 前处理止于此处 -----')                                                                              |
| `add_fig_caption(doc, text)`                                                                        | 图标题，位于图下方（自动："图N　描述"，标签加粗，全角空格分隔）                                                                 |
| `add_table_caption(doc, text)`                                                                      | 表标题，位于表上方（自动："表N　描述"，标签加粗，全角空格分隔）                                                                 |
| `reset_counters()`                                                                                  | 重置图表计数器为零（由 `setup_document()` 自动调用）                                                                           |
| `add_note(doc, text)`                                                                               | 居中斜体注释（9pt），位于图下方                                                                                           |
| `add_bibliography(doc, entries, title='参考文献')`                                                  | 参考文献表：黑体三号居中标题（Heading 1，进入目录）+ GB/T 7714-2015 顺序编码制条目，宋体五号，悬挂缩进，[n] 自动编号与正文引用上标对应 |
| `add_toc(doc, title, levels, auto_update)`                                                          | **按需**目录（黑体三号居中标题 + TOC域，levels '1-3'；`auto_update` 默认 `True`，见「目录（TOC）」一节）。末尾自动调用 `start_body()` 分节 |
| `start_body(doc)`                                                                                   | **按需**正文分节：把此前内容（封面、目录）划为不显示页码的前置节，正文另起一节并从 1 开始编号。`add_toc()` 末尾已自动调用；无目录文档需在正文前手动调用；重复调用为幂等空操作。若末节仍是空的（封面节刚以分节符结束），会直接复用该节而不新建，避免空白页（见「页码」一节） |
| `set_table_border(table)`                                                                           | 设置所有边框为单线黑色                                                                                                    |
| `set_cell_shading(cell, color_hex)`                                                                 | 设置单元格背景色（如 'D9D9D9' 灰色）                                                                                        |
| `safe_extract(zf, dest)`                                                                            | 安全解压 .docx ZIP 包，防止路径遍历和符号链接攻击（可选功能用）                                                                          |
| `rezip(src_dir, out_path)`                                                                          | 将目录重新打包为 .docx 文件，确保 `[Content_Types].xml` 首位存储（可选功能用）                                                         |

## 文档结构与函数映射

生成文档时，请根据结构元素选择正确的函数：

| 场景              | 示例                      | Python 函数                                           | 格式                       | 进入目录              |
| --------------- | ----------------------- | --------------------------------------------------- | ------------------------ | ----------------- |
| 封面页             | 项目名 + 技术方案 + 单位 + 日期     | `add_cover_page(doc, title, subtitle, org, date)`   | 黑体 26pt 居中，顶部留白定位；独立成节且不显示页码 | 否                 |
| 文档主标题           | `技术交底书`                 | `add_title(doc, text)`                              | 黑体 16pt, 居中, 加粗          | 否                 |
| 一级章节（一、二、）      | `一、发明名称`                | `add_h1(doc, text)`                                 | 黑体 16pt, 左对齐, 加粗         | **是 (Heading 1)** |
| 二级编号标题          | `3.1 现有技术...`           | `add_h2(doc, text)`                                 | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 二级步骤标题          | `步骤1：构建...`             | `add_h2(doc, text)`                                 | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 二级括号标题          | `（一）系统总体架构`             | `add_h2(doc, text)`                                 | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 2)** |
| 三级标题            | `6.1 xxx方法`             | `add_h3(doc, text)`                                 | 黑体 12pt, 左对齐, 加粗         | **是 (Heading 3)** |
| **编号分点（加粗标签）**  | `（1）定义核心概念...：内容`       | `add_item_para(doc, "（1）定义...：", "内容")`             | 宋体 12pt, 加粗标签 + 正文, 首行缩进 | <br />            |
| **字母子步骤（加粗标签）** | `a. 任务分解：内容`            | `add_item_para(doc, "a. 任务分解：", "内容")`              | 宋体 12pt, 加粗标签 + 正文, 首行缩进 | <br />            |
| 普通正文段落          | 技术描述段落                  | `add_body(doc, text)`                               | 宋体 12pt, 两端对齐, 缩进2字符     | <br />            |
| 块级数学公式          | `L = Σ ...`             | `add_eq_para(doc, math([...]))`                     | 居中，1.5倍行距                | <br />            |
| 行内公式混排          | `其中 L_LLM 为...`         | `add_body_with_math(doc, parts)`                    | 宋体 12pt, 两端对齐            | <br />            |
| 代码块             | 命令/源码片段                 | `add_code_block(doc, code)`                         | Consolas 9pt, 灰底, 无首行缩进  | <br />            |
| 数据表             | 参数表 / 检查项表 / 功能矩阵       | `add_data_table(doc, headers, rows, col_widths)`    | 灰色表头, 固定列宽               | <br />            |
| 分层架构图           | 五层系统架构图                 | `add_layered_architecture(doc, layers, col_widths)` | 表格框图堆叠, 首行加粗层名, 无连接线     | <br />            |
| 表格单元格内公式        | 公式对照表第三列                | `add_math_to_cell(cell, inlineMath([...]))`         | OMML 居中                  | <br />            |
| 图标题             | `系统架构图` → 自动 "图1　系统架构图" | `add_fig_caption(doc, text)`                        | 宋体 10.5pt, 标签加粗, 居中, 图下方 | <br />            |
| 表标题             | `参数对比` → 自动 "表1　参数对比"   | `add_table_caption(doc, text)`                      | 宋体 10.5pt, 标签加粗, 居中, 表上方 | <br />            |
| 图注（斜体）          | `（核心：...）`              | `add_note(doc, text)`                               | 宋体 9pt, 居中, 斜体           | <br />            |
| 正文分节（页码起算）      | 封面/目录之后                 | `start_body(doc)`                                   | 分节符（下一页），前置部分无页码，正文从 1 起 | 否                 |

### 关键规则

1. **（1）/（2）/（3）分点**：label 部分加粗（含冒号），正文部分正常，使用 `add_item_para`，首行缩进2字符
2. **a. / b. / c. 子步骤**：同上格式，label 加粗（含冒号），使用 `add_item_para`
3. **步骤1：/ 步骤2：**：作为二级标题处理，使用 `add_h2`，黑体小四加粗左对齐
4. **（一）/（二）**：作为二级标题处理，使用 `add_h2`
5. **3.1 / 3.2**：作为二级标题处理，使用 `add_h2`
6. **正文中的** **`**bold**`**：`add_body` 自动解析加粗标记
7. **正文中的** **`[n]`** **引用**：`add_body` 自动渲染为上标
8. **Word 目录（TOC）支持**：`add_h1`/`add_h2`/`add_h3` 使用 Word 内置 Heading 样式。使用 `add_toc` 按需插入目录域（TOC field），Word 打开文档时提示更新域，选择"是"即生成目录条目。`add_title` 不使用 Heading 样式，不进入目录
9. **标题字符格式写在样式上，不在 run 上**：H1–H3 的字体、字号、加粗、黑色全部由 `_setup_heading_styles()`（`setup_document()` 内自动调用）写到 Word 内置 Heading 1/2/3 样式上，标题 run 不带任何直接字符格式。这样既覆盖了模板默认的蓝色标题（清除 `color` 的 `themeColor` 引用）与主题字体引用，又让 Word 更新目录域时无法把标题字体搬进条目，目录条目因此严格按 TOC 样式呈现。`add_h1`/`add_h2`/`add_h3`/`add_bibliography` 都遵循此约定，**不要在调用脚本里给标题 run 直接设字体**
10. **图表自动编号**：`add_fig_caption` 和 `add_table_caption` 自动递增编号（图N / 表N），无需手动填写编号。传空字符串 `""` 给 `add_fig_caption` 可生成纯间距段落（不编号）。图标题在图**下方**，表标题在表**上方**
11. **页码从正文开始**：封面与目录不显示页码，正文第一页为 1。含目录时由 `add_toc()` 自动分节；有封面但无目录时需手动调用 `start_body(doc)`；无封面无目录的单节文档无需任何处理
12. **封面独立成节**：`add_cover_page` 必须以「下一页」分节符结束封面节，因此它**必须在写任何其它内容之前调用**，且用了它就不必再调 `add_title`。分节链上 `add_toc()` 与 `start_body()` 会自动识别并复用分节符留下的空节，不会产生空白页

## 行距与标题间距

所有间距值对齐 **NJUThesis** LaTeX 模板（`linespread = 1.625`，等效 Word 1.5倍行距）。

### 间距参数

| 元素     | 行距     | 段前   | 段后   | 说明                                               |
| ------ | ------ | ---- | ---- | ------------------------------------------------ |
| 文档标题   | 1.5倍   | 24pt | 18pt | 居中，黑体 26pt（一号）                                 |
| H1（章）  | 1.5倍   | 24pt | 6pt  | 左对齐，黑体 16pt。北大研究生学位论文写作指南标准             |
| H2（节）  | 1.5倍   | 12pt | 6pt  | 左对齐，黑体 12pt。北大研究生学位论文写作指南标准             |
| H3（小节） | 1.5倍   | 12pt | 6pt  | 左对齐，黑体 12pt。北大研究生学位论文写作指南标准             |
| 正文段落   | 1.5倍   | 0    | 7pt  | 两端对齐，宋体 12pt，首行缩进2字符                             |
| 分点段落   | 1.5倍   | 0    | 7pt  | 加粗标签 + 正文，首行缩进                                   |
| 块级公式   | 1.5倍   | 6pt  | 6pt  | 居中                                               |
| 目录·章标题行 | 固定20pt | 6pt  | 0    | TOC 1 样式，独立于正文间距                               |
| 目录·节标题行 | 固定20pt | 0    | 0    | TOC 2/3 样式，独立于正文间距                             |

### NJUThesis 行距背景

NJUThesis 全局使用 `linespread = 1.625`。计算方式：LaTeX 默认行距倍数为 1.2，Word 默认为 1.3，要求 1.5倍 Word 行距，故 `1.5 × (1.3 / 1.2) = 1.625`。在 Word 中设置 1.5倍行距即可达到相同视觉效果。

### 实现说明

- `pf.line_spacing = 1.5` 设置 Word 1.5倍行距。`pf.space_before` / `pf.space_after` 控制段落间距。

- 标题前的空段落（`doc.add_paragraph()`）已移除，间距改由 `space_before` / `space_after` 精确控制，结构更干净。

## 排版预设（PRESETS / set\_preset）

排版参数集中在 `docx_layout_kit.py` 顶部的 `PRESETS` 字典中（嵌套分组），不在各 `add_*` 函数内散落硬编码。

**分组结构**：

| 分组            | 键                                                                                                       | 说明                                  |
| ------------- | ------------------------------------------------------------------------------------------------------- | ----------------------------------- |
| `west`        | `separate` / `body` / `head`                                                                            | 中西文字体分离开关及西文正文/标题字体（默认关闭）           |
| `title`\~`h3` | `font` / `size` / `bold` / `color` / `line_spacing` / `space_before` / `space_after`                    | 标题样式                                |
| `body`        | `font` / `size` / `line_spacing` / `first_line_indent` / `space_before` / `space_after`                 | 正文；间距为 `None` 表示不设置、沿用样式继承（默认不设段间距） |
| `item`        | `space_before` / `space_after`                                                                          | 分点段落间距                              |
| `caption`     | `size`                                                                                                  | 图/表标题字号                             |
| `note`        | `size`                                                                                                  | 图注字号                                |
| `page`        | `page_width` / `page_height` / `margin_top` / `margin_bottom` / `margin_left` / `margin_right` / `header_distance` / `footer_distance` | 纸张尺寸（cm）与页边距、页眉页脚距离（cm）         |
| `toc`         | `title_font` / `title_size` / `line_spacing` / `levels`                                                  | 目录标题与条目：`levels` 按层级给出各条目的字体/字号/段前/段后/缩进（汉字符） |
| `cover`       | `top_padding_ratio` / `spacer_height` / `title_*` / `subtitle_*` / `org_*` / `date_*`                     | 封面：顶部留白占版心高的比例、留白空段单段高（pt）、标题块各元素字体/字号/间距（pt） |

**切换方式**：`set_preset('name')` 整套切换，`add_*` 函数读取当前激活预设。默认激活 `default`（数值与「行距与标题间距」一节完全一致），无需显式调用。**须在** **`setup_document()`** **之前调用**：页面边距与 Normal 默认字体在 `setup_document()` 时读取预设。

**中西文分离**：将预设 `west.separate` 设为 `True` 后，西文字符使用 `west.body` / `west.head` 指定的字体（如 Times New Roman），中文保持各元素字体不变。

**不覆盖**（按约定保持函数内硬编码，避免过度抽象）：代码块、数据表、框图、箭头、公式段。

**新增预设**：向 `PRESETS` 添加一份同键结构的嵌套字典即可，例如紧凑版或西文版。

## 分层架构图（add\_layered\_architecture）

用表格框图堆叠表达多层架构（如"采集层 → 预处理层 → 三个并列模块 → 应用层"），配合 `add_fig_caption` 生成图标题。

```python
add_layered_architecture(doc, [
    ["采集层：传感器阵列 / 信号调理 / 抗混叠滤波"],
    ["预处理层：去直流 / 滤波 / 加窗"],
    [["变换模块", "快速傅里叶变换"],
     ["分析模块", "频带能量 / 谱质心"],
     ["决策模块", "模式匹配 / 阈值判断"]],
    ["应用层：频谱显示 / 报告输出"],
], col_widths=[4.0, 5.0, 5.0])
```

- 元素为 `list[str]` → 全宽层框（首行加粗层名，其余居中内容行）

- 元素为 `list[list[str]]` → 一行并列框（每个框首行加粗层名）

- 整图总宽**只有一个来源**：传 `col_widths` 时**其和即总宽**，全宽层与并列层共用同一个总宽，上下框体必然等宽；只传 `width_cm`（默认 `TABLE_WIDTH_CM` = 14cm）时它即总宽，并列层按总宽均分

- 层间自动插入小空段（技术上必须：无间隔段落时 Word 会把相邻表格合并为一张表）

- 第一版**不画连接线/箭头**；需要时可在层间手动插入 `add_arrow_down`

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

- "过程 → 产出"步骤（如 `① 时域采样 → 离散信号 x(n)`）

- 需要用底纹区分层级的流程图（如预处理层 → 决策层，左灰右白）

- 任何需要横向箭头连接不同长度文字框的流程图

### 用法

以下示例展示两种底纹模式：前两步地位相等（全白），后两步层级转换（左灰右白）。

```python
# 前两步：同属预处理层级，地位相等 → 全白
add_arrow_row(doc, "① 时域采样", "离散信号 x(n)")
add_arrow_down(doc)
add_arrow_row(doc, "② 加窗处理", "加窗信号")
add_arrow_down(doc)

# 后两步：从预处理层进入决策层，层级不一 → 左灰右白
add_arrow_row(doc, "③ 离散傅里叶变换", "频谱 X(k)", left_shade='D9D9D9')
add_arrow_down(doc)
add_arrow_row(doc, "④ 频谱分析", "特征向量", left_shade='D9D9D9')
add_arrow_down(doc)

add_separator_note(doc, "----- 信号预处理止于此处，后续由分类器承接 -----")
add_arrow_down(doc)
add_box(doc, "分类决策：特征归一化 / 模式匹配 / 结果输出")
add_fig_caption(doc, "信号处理与分类流程图")
```

### 实现原理

1. **3 列表格**：左框 | 箭头列 | 右框
2. **左右等宽**：`side = (total - arrow) / 2`，布局常数，不依赖内容
3. **固定布局**（`tblLayout=fixed`）：防止 Word autofit 重排窄箭头列
4. **融合中间竖线**（`insideV=none`）：去掉左/箭头/右之间的竖线，外框 + insideH 保留
5. **可选底纹**：`left_shade` / `right_shade` 默认 None（全白），层级不同时左灰右白

## 封面（add\_cover\_page）— 按需生成

封面**默认不生成**。仅在文档需要封面时调用 `add_cover_page()`。

### 版式

居中标题块，用**顶部留白**定位（不是垂直居中）：标题块距版心顶端一段按比例算出的距离，下方留白自然延伸。默认比例 0.264，A4 下留白 6.50cm，大标题段起点距页顶约 9.04cm（页面高度 30.4%）。

| 元素  | 字体          | 字号   | 段后      |
| --- | ----------- | ---- | ------- |
| 大标题 | 黑体 (SimHei) | 26pt | 8pt     |
| 副标题 | 黑体 (SimHei) | 20pt | 60pt    |
| 单位  | 宋体 (SimSun) | 15pt | 0       |
| 日期  | 宋体 (SimSun) | 15pt | 段前 8pt  |

`subtitle` / `org` / `date` 均可省略，省略则该行不生成；`title` 必填。

### 留白如何自动适配

留白高度**不写死行数**，而是按当前预设的版心高现算：

```
版心高   = 纸张高 − 上边距 − 下边距        （预设 page 组）
目标留白 = cover.top_padding_ratio × 版心高
空段数 N = 取整(目标留白 ÷ cover.spacer_height)
余数     = 目标留白 − N × spacer_height    （挂到大标题段的段前）
```

留白空段使用**固定行距**（`lineRule="exact"`），每段高度严格等于 `spacer_height`，故总留白 = N × spacer_height + 余数 = 目标留白，误差远小于 1 磅。余数之所以挂到大标题段的段前，是因为该段前面已有 N 个空段、不在页首，可避开 Word「页首段落抑制段前间距」的行为。

换纸张（A4/A5/Letter）、改页边距或整套切换预设时，留白随之等比换算，封面视觉位置保持一致，无需再调参。

**默认 `top_padding_ratio = 0.264`**，A4 下留白 6.50cm、大标题段起点距页顶约 9.04cm（页面高度 30.4%）。比例是**固定值**而非按内容垂直居中，故行数少的封面（只给大标题 + 副标题）上缘位置不变、下方留白自然更大。

### 用法

```python
doc = setup_document()

add_cover_page(doc, "文档标题",
               subtitle="技术方案",
               org="某某科技有限公司",
               date="2026年9月")

add_toc(doc)                     # 目录；add_toc 末尾自动分节，页码从正文起算
add_h1(doc, "一、概述")
```

```python
# 有封面但不生成目录：封面之后手动分节
add_cover_page(doc, "文档标题", subtitle="技术方案")
start_body(doc)
add_h1(doc, "一、概述")
```

### 注意事项

- **必须在写任何其它内容之前调用**：它以「下一页」分节符结束封面节，之后的内容自然落到新的一页
- 用了封面就**不必**再调 `add_title`（`add_title` 用于不放封面时置于正文首行的文档主标题）
- 封面节不显示页码；其后的目录节同样不显示，正文从 1 起（见「页码」一节）
- 留白比例上限为 0.5：超过半个版心会把标题块挤到第二页
- 留白由**空段落**构成。Word 与 WPS 正常渲染；若交付链路中经过会剥离空段落的 docx→HTML 转换器，留白可能丢失

## 目录（TOC）— 按需生成

目录**默认不生成**。仅在文档需要目录时调用 `add_toc()`。

### 格式规范

| 元素     | 对应标题 | 字体          | 字号        | 行距     | 段前  | 段后 | 缩进      |
| ------ | ---- | ----------- | --------- | ------ | --- | -- | ------- |
| 目录标题   | —    | 黑体 (SimHei) | 16pt (三号) | —      | —   | —  | 居中      |
| 章标题行   | H1   | 黑体 (SimHei) | 12pt (小四) | 固定20pt | 6pt | 0  | 居左，不缩进  |
| 一级节标题行 | H2   | 宋体 (SimSun) | 12pt (小四) | 固定20pt | 0   | 0  | 缩进1个汉字符 |
| 二级节标题行 | H3   | 宋体 (SimSun) | 12pt (小四) | 固定20pt | 0   | 0  | 缩进2个汉字符 |

缩进以「汉字符」为单位写入：`w:ind` 同时给出 `w:leftChars`（100 = 1 个汉字符）与 `w:left`（twips 兜底，1 个汉字符宽 = 字号 × 20，小四 12pt → 240 twips）。

### 工作原理

1. **标题样式**：`add_h1` / `add_h2` / `add_h3` 使用 Word 内置 Heading 1/2/3 样式，TOC 域可识别。样式本身的字符格式由 `_setup_heading_styles()` 在 `setup_document()` 内写入，标题 run 不带直接字符格式——这是目录条目能按 TOC 样式呈现字体（而非跟随标题黑体）的前提
2. **目录条目样式**：TOC 1 / TOC 2 / TOC 3 三级样式按上表预定义：`_setup_toc_styles(doc)` 在 `add_toc()` 内部调用，逐级配置字体、字号、固定行距、段前后与缩进；样式不带 `customStyle`，`basedOn` 指向 Normal
3. **域生成**：目录是 Word 域（`TOC \o "1-3" \h \z \u`）。条目由 Word 在更新域时计算生成，生成前显示占位文字
4. **层级**：默认 `'1-3'`，包含 H1、H2、H3 三级标题
5. **打开时更新**：默认 `auto_update=True`，在 `word/settings.xml` 写入 `w:updateFields` 开关，Word 打开文档时提示更新域，确认后生成目录条目；`auto_update=False` 不写入该开关，仅在目录上右键 → "更新域"时手动生成

### 用法

```python
doc = setup_document()
add_title(doc, "文档标题")
add_toc(doc)                      # 默认收 1-3 级：Word 打开时提示更新域，选"是"生成条目
# add_toc(doc, levels='1-2')      # 只收 H1 + H2
# add_toc(doc, auto_update=False)  # 改为完全手动：仅右键"更新域"时生成条目
add_h1(doc, "一、概述")
add_body(doc, "正文内容...")
```

### 注意事项

- Word 打开文档时提示"是否更新此文档中的域"：选择"是"生成目录条目；选择"否"则保留占位文字，可随时在目录上右键 → "更新域"手动生成

- 更新前显示占位文字"（请在 Word 中右键此处选择"更新域"以生成目录）"

- 打开时更新由 `auto_update` 参数控制（默认开启），设为 `False` 后目录仅能手动更新

- `add_title` 不使用 Heading 样式，文档主标题不会出现在目录中

- 目录末尾是**分节符（下一页）**，与正文分隔并触发正文页码重新起算；不要再另加分页符

- **条目字体由标题样式的写法决定**：Word 更新目录域时，会把标题 run 上的字体（run 级直接格式）写进目录条目，其优先级高于 TOC 样式。本 skill 的标题一律**不在 run 上设字体**，H1–H3 的字体/字号/加粗/颜色全部写在 Heading 1/2/3 样式上（见 `_setup_heading_styles()`），标题 run 不带任何字符格式，因此 Word 无从覆盖，条目按 TOC 样式呈现：章标题行为黑体、一级/二级节标题行为宋体，无需在 Word 中手工调整
- **不要在调用脚本里给标题 run 设字体**：若改用 `set_run_font` 等方式给标题文字直接设字体，Word 更新域时会把该字体搬进条目，覆盖 TOC 样式对宋体的声明——这正是本 skill 把标题字符格式放在样式层的原因

## 页码 — 从正文开始编号

封面页与目录页**不显示页码**，页码从正文第一页起算为 1。该行为由 `start_body(doc)` 实现，它把此前内容划为前置节、正文另起一节，并重置页码计数。

### 何时需要调用

| 文档结构            | 是否需要调用 `start_body()`                     |
| --------------- | ---------------------------------------- |
| 无封面、无目录（单节文档）   | 不需要，`setup_document()` 已让页码从第 1 页起连续编号     |
| 有目录（无论有无封面）     | 不需要，`add_toc()` 末尾已自动调用 `start_body()` |
| 有封面但无目录         | **需要**，在正文第一个标题前手动调用 `start_body(doc)`  |

封面由 `add_cover_page()` 生成，它自己就把封面结束成一个独立节；有目录时 `add_toc()` 接在封面之后，形成「封面节 → 目录节 → 正文节」三节结构（封面、目录两节均不显示页码）。

### 用法

```python
# 场景一：封面 + 目录（add_cover_page 结束封面节，add_toc 末尾自动分节）
doc = setup_document()
add_cover_page(doc, "文档标题", subtitle="技术方案")
add_toc(doc)
add_h1(doc, "一、概述")          # 正文从这里开始，页码为 1

# 场景二：封面 + 正文，不生成目录（需手动分节）
doc = setup_document()
add_cover_page(doc, "文档标题", subtitle="技术方案")
start_body(doc)                  # 封面节到此结束，正文另起一页
add_h1(doc, "一、概述")          # 页码为 1

# 场景三：不放封面，用主标题开篇 + 目录
doc = setup_document()
add_title(doc, "文档标题")
add_toc(doc)
add_h1(doc, "一、概述")          # 页码为 1
```

### 工作原理

1. **前置节**：封面与目录各自成一节（`add_cover_page()` 结束封面节，`add_toc()` 结束目录节）。`start_body()` 会清空**所有**前置节的页脚，因此封面页与目录页都不显示任何页码
2. **正文节**：新建一节（分节符类型为"下一页"，因此兼作封面/目录与正文之间的分页），页面参数（纸张、页边距、页眉页脚距离）沿用预设
3. **页码重置**：正文节写入 `w:pgNumType`（`w:start` 取预设 `page.body_page_start`，默认 1），页脚居中放置 PAGE 域，五号 Times New Roman
4. **幂等**：重复调用 `start_body()` 不会重复分节，第二次起为空操作
5. **不产生空白页**：分节符本身已换页，调用后直接写正文首个标题即可。`add_toc()` 在末节为空时（紧跟封面分节符）会跳过自己的前导分页符；`start_body()` 在末节仍为空时会直接复用该节，不再新建

## 排版标准

### 字体与字号

| 元素       | 字体          | 字号        | 样式      |
| -------- | ----------- | --------- | ------- |
| 文档标题     | 黑体 (SimHei) | 16pt (三号) | 加粗，居中   |
| H1（一级标题） | 黑体 (SimHei) | 16pt (三号) | 加粗，左对齐  |
| H2（二级标题） | 黑体 (SimHei) | 12pt (小四) | 加粗，左对齐  |
| H3（三级标题） | 黑体 (SimHei) | 12pt (小四) | 加粗，左对齐  |
| 正文       | 宋体 (SimSun) | 12pt (小四) | 常规，两端对齐 |
| 目录标题     | 黑体 (SimHei) | 16pt (三号) | 加粗，居中   |
| 目录·章标题行   | 黑体 (SimHei) | 12pt (小四) | 常规，居左   |
| 目录·一级节标题行 | 宋体 (SimSun) | 12pt (小四) | 常规，缩进1字 |
| 目录·二级节标题行 | 宋体 (SimSun) | 12pt (小四) | 常规，缩进2字 |
| 表格单元格    | 宋体 (SimSun) | 9-10.5pt  | 常规      |
| 表格框图     | 宋体 (SimSun) | 10-14pt   | 标题加粗    |

### 页面设置

- 纸张大小：A4（21 × 29.7 cm，由预设 `page.page_width` / `page.page_height` 给定；左右边距 3.18cm 时版心宽 14.64cm）

- 页边距：上下 2.54cm，左右 3.18cm（对齐 NJUThesis hmargin=3.18cm）

- 行距：1.5倍

- 首行缩进：24pt = 2个中文字符

- 页码：页脚居中，五号 Times New Roman，PAGE 域自动编号（对齐 NJUThesis plain 页脚样式）

- 页码起算：单节文档从第 1 页起连续编号；含封面或目录的文档由 `start_body()` 分节，前置部分不显示页码，正文节从 `page.body_page_start`（默认 1）起编号（见「页码」一节）

- 封面留白：顶部留白 = `cover.top_padding_ratio` × 版心高（版心高 = 纸张高 − 上下页边距 ≈ 697.9pt），换算为固定行距空段与大标题段前余数（见「封面」一节）

- 中西文字体分离：正文西文 Times New Roman 衬线，标题西文 Arial 无衬线（对齐 NJUThesis 正文衬线、标题 \sffamily）

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

## 参考文献表（add_bibliography）

生成 GB/T 7714-2015 顺序编码制参考文献表，格式对齐 NJUThesis LaTeX 模板。正文引用上标 `[n]` 与条目序号一一对应：`add_body` 中的 `[1]` 指向 `entries[0]`，以此类推。

**排版**：

| 元素     | 格式                                       |
| ------- | ---------------------------------------- |
| 章节标题   | 黑体 三号(16pt) 加粗居中，Heading 1 样式，进入目录    |
| 条目     | 宋体 五号(10.5pt)，两端对齐，1.5倍行距，悬挂缩进2字符，段后3pt |
| 序号     | `[1]`、`[2]`... 按传入顺序自动生成，不需手写            |

**条目格式**（GB/T 7714-2015 各文献类型，传入时不含序号）：

| 类型   | 格式                                        |
| ---- | ----------------------------------------- |
| 期刊   | 作者. 题名[J]. 刊名, 年, 卷(期): 页码               |
| 图书   | 作者. 书名[M]. 出版地: 出版社, 年                    |
| 会议   | 作者. 题名[C]//会议名. 出版地: 出版社, 年: 页码           |
| 学位论文 | 作者. 题名[D]. 城市: 学校, 年                     |
| 电子资源 | 作者. 题名[EB/OL]. (更新日期)[引用日期]. 访问路径         |

**用法**：

```python
add_bibliography(doc, [
    "张三, 李四. 示例期刊论文题名[J]. 示例期刊, 2020, 12(3): 45-56",
    "王五. 示例图书题名[M]. 示例市: 示例出版社, 2019",
    "赵六. 示例学位论文题名[D]. 示例市: 示例大学, 2021",
])
```

**规则**：

1. 条目序号自动生成，正文 `[n]` 引用与条目顺序必须一致
2. 多作者用逗号分隔，三位以上作者取前三名加 "等" 或 "et al"
3. 中英文作者混排时保持原语言，不强制转换
4. 标题默认为"参考文献"，可通过 `title` 参数自定义

## OMML 数学公式指南

数学公式由 `omml_math_kit.py` 直接构建 OMML（Office Math Markup Language）XML，经 `docx_layout_kit.py` 的 `add_eq_para()` / `add_body_with_math()` 插入文档。生成的是 Word 原生可编辑公式，与手动插入的公式完全一致。

### 基本元素

```python
from omml_math_kit import r, sub, sup, frac, sumOp, func, paren, bracket, math, inlineMath

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
from docx_layout_kit import add_eq_para
from omml_math_kit import r, sub, sumOp, func, math

eq = math([
    sub("L", "LLM"), r(" = - "),
    sumOp([r("i")], [sub("y", "i"), func("log", [sub("p", "i")])]),
])
add_eq_para(doc, eq)
```

**行内公式（与正文混排）：**

```python
from docx_layout_kit import add_body_with_math
from omml_math_kit import sub, inlineMath

add_body_with_math(doc, [
    ("text", "其中，"),
    ("math", inlineMath([sub("L", "LLM")])),
    ("text", "为语言模型损失项，其计算涉及交叉熵。"),
])
```

### 预定义公式（formula_templates.py）

`formula_templates.py` 提供 eq1\~eq4 常用公式模板（总损失、交叉熵、对比学习 InfoNCE、LoRA 分解），可直接使用或作为新增公式的参考模式：

```python
from formula_templates import eq1, eq2, eq3, eq4
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

### 代码块（add\_code\_block）

命令行、源码片段等代码内容使用等宽字体 + 浅灰底纹渲染，**不要**将代码放入普通正文段落：

```python
add_code_block(doc, "python docx_validator.py output.docx --verbose")

add_code_block(doc, """from docx_layout_kit import setup_document
doc = setup_document()""")
```

### 数据表（add\_data\_table）

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

注意：`col_widths` 之和即表格总宽，建议 ≤ 14.64cm（A4 21cm 减左右各 3.18cm 的版心宽）；数据行末列左对齐（描述列），其余居中。

### 单元格内公式（add\_math\_to\_cell）

在表格单元格中插入行内 OMML 公式（如"函数调用 vs 渲染效果"对照表）：

```python
from omml_math_kit import sub, inlineMath
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

**宽度**：所有表格（单框 / 多行框 / 并行框 / 数据表 / 箭头行）都由内部 `_new_table()` 建立，**表总宽 = 传入列宽之和**——一次性写入 `w:tblW`（dxa）、`w:tblGrid`、每个 `w:tcW`，并锁定 `tblLayout=fixed`。调用方只提供列宽、不要自己设宽；同一张图内各行的列宽之和保持相等，框体才会左右对齐。默认总宽见常量 `TABLE_WIDTH_CM`（14cm）。

## 工作流程

1. **阅读**用户提供的源文本/内容
2. **阅读范例** `scripts/_build_readme_docx.py` —— 这是生成文档的标准骨架，理解其 import 段与调用顺序（封面 → 目录 → 标题 → 正文 → 图表 → 参考文献 → 保存）
3. **复制构建库**从本 skill 目录的 `scripts/` 到工作目录：

   - 无公式文档：仅 `docx_layout_kit.py`

   - 含公式文档：`docx_layout_kit.py` + `omml_math_kit.py` + `formula_templates.py`

   - **不要复制** **`scripts/optional/`**（可选功能专用）
4. **编写调用脚本** — 在**新建的脚本**中 `import` 需要的 `add_*` 函数，照范例的顺序拼出内容

   - `docx_layout_kit.py` 是**无入口的库**，不要直接运行它，也不要改它的内容；**文档内容写在你自己的脚本里**
   - 注意「如何开始」一节列出的三条顺序规则
5. **运行脚本** — 输出到用户工作区文件夹
6. **验证** — 使用纯 Python 标准库验证工具检查文档完整性：

   ```bash
   python docx_validator.py output.docx
   # 或详细输出模式
   python docx_validator.py output.docx --verbose
   ```

   验证工具执行 5 项检查：ZIP 完整性、XML 格式良好性、文件引用完整性、内容类型声明、空白保留。所有检查通过退出码为 0，否则为 1。
7. **分享**文件，使用 `computer://` 链接

## 常见陷阱

1. **禁止使用 emoji** — 文档中仅使用 Unicode 符号
2. **文件锁定**：如果 Word 已打开文件，保存会静默失败 — 使用不同文件名
3. **禁止手写 m:nary 空上标**：始终使用 omml_math_kit.py 中的 `sumOp()` — 它使用 `m:sSub` 配合 Unicode ∑，避免空上标渲染问题
4. **始终展平列表**：所有数学辅助函数使用 `_flat()` — 新增函数时需对所有 children 参数应用 `_flat()`
5. **字体东亚设置**：python-docx 中必须设置 `rFonts.set(qn('w:eastAsia'), font_name)` 才能正确渲染中文字体
6. **表格列宽**：所有表格（框图 / 数据表 / 箭头行）一律经内部 `_new_table()` 建立——它把 `w:tblW`（dxa 总宽）、`w:tblGrid/gridCol`、每个 `w:tcW` 从**同一份列宽**一次写入并把布局锁为 fixed，宽度只有一个来源；默认总宽取常量 `TABLE_WIDTH_CM`（14cm）。**不要绕开它用 `doc.add_table()` 自建表格**，否则宽度会出现第二个口径，字多的那格被 Word autofit 撑宽，同一张图上下框体就会不等宽
7. **表格底纹**：使用 `WD_FILL_PATTERN.CLEAR` 类型底纹而非其他填充模式，否则部分 Word 版本渲染异常
8. **列表符号**：使用 Word 内置编号样式，切勿在文本中直接写 `•` 字符
9. **脚本级变量**：`_fig_counter` / `_tbl_counter` 在 `setup_document()` 中自动重置。单进程多次生成文档时，只要重新调用 `setup_document()` 即可从 图1/表1 开始

## 文档验证（docx_validator.py）

`docx_validator.py` 是纯 Python 标准库实现的轻量验证脚本，无外部依赖。用于在生成 .docx 后快速检查文档结构完整性。

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
python docx_validator.py output.docx

# 详细模式（显示每项 PASSED）
python docx_validator.py output.docx --verbose
```

### 退出码

- `0` — 全部检查通过

- `1` — 一项或多项检查失败

### 编程式调用

```python
from docx_validator import validate_docx

if validate_docx("output.docx", verbose=True):
    print("文档验证通过")
else:
    print("文档存在问题，请检查上方输出")
```

## 安全处理规范

处理来自不可信来源的 .docx 文件时，必须遵循以下安全规范：

### 安全解压（safe\_extract）

`docx_layout_kit.py` 中的 `safe_extract()` 函数在解压 .docx ZIP 包时提供两层防护：

1. **符号链接拒绝**：检测 `external_attr` 中的 `S_ISLNK` 标志，拒绝包含符号链接的条目
2. **路径遍历防护**：解析每个条目的目标路径，验证其不超出目标目录边界

```python
from docx_layout_kit import safe_extract, rezip
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
解压（safe_extract）→ 合并碎片 run（可选）→ 编辑 XML → 重新打包（rezip）→ 验证（docx_validator）
```

> **注意**：编辑 XML 时保持原始格式，不要 pretty-print。`xml.etree.ElementTree` 的 `tostring()` 默认不保留原始缩进。

## 文件验证清单

交付 .docx 文件前：

- [ ] `docx_validator.py` 全部 5 项检查通过

- [ ] 文件大小 > 10KB（空文档约 3KB）

- [ ] 内容中无 emoji 字符

- [ ] 所有数学符号渲染为 OMML（非纯文本）

- [ ] 图表使用表格框模式（非图片）

- [ ] 表格有边框和标题底纹（D9D9D9，灰色）

- [ ] 所有公式使用 `math()` / `inlineMath()`，非原始文本加下标字符

- [ ] 标题颜色为黑色（非 Word 默认蓝色），且字符格式写在 Heading 样式中、标题 run 不带直接字符格式（否则更新目录域后条目字体会跟随标题）

- [ ] 图表编号从 图1/表1 开始（`setup_document()` 自动重置）

- [ ] 含封面时：封面节与目录节均不显示页码，正文首页为 1，且封面与目录之间无空白页

***

## 可选功能详解

以下章节**默认不使用**。仅当用户明确要求对应功能时才阅读和执行。所有可选依赖（Windows 上渲染 PDF 所需的 pywin32、LibreOffice、lxml、pandoc 等）按需安装，默认不检查。

### 可选功能 A：PDF 渲染验证

**用途**：将生成的 .docx 转换为 PDF 进行视觉检查，确认排版正确。

**两条路径，按平台择一使用**：

| 路径                    | 依赖                                 | 适用平台    | 脚本                                        |
| --------------------- | ---------------------------------- | ------- | ----------------------------------------- |
| Windows Word（Windows 首选） | Microsoft Word + `pip install pywin32` | Windows | `scripts/optional/office/word2pdf.py`     |
| LibreOffice           | LibreOffice（`soffice` 命令行）          | 跨平台     | `scripts/optional/office/soffice.py`      |

**Windows Word 路径**（Windows 上本机装有 Word 时的首选）：调用 Word 自身的排版引擎，中文字体、OMML 公式、目录域与页脚页码分节均与在 Word 中打开时一致，较 LibreOffice 更贴近真实 Word 渲染效果。导出前先更新域，因此目录不会残留占位文字。每次使用**独立**的 Word 实例、文档以**只读**方式打开，导出后关闭文档并退出该实例，不干扰用户已打开的 Word 窗口，也不残留后台 `WINWORD.EXE` 进程。

```bash
# 默认输出到源文件同目录同名（output.docx → output.pdf）
python scripts/optional/office/word2pdf.py output.docx

# 指定输出文件 / 输出目录
python scripts/optional/office/word2pdf.py output.docx -o dist/output.pdf
python scripts/optional/office/word2pdf.py output.docx --outdir ./pdf_out

# 导出 PDF/A-1b
python scripts/optional/office/word2pdf.py output.docx --pdfa
```

编程式调用：

```python
from word2pdf import convert_docx_to_pdf

convert_docx_to_pdf("output.docx", "output.pdf")
```

**LibreOffice 路径**：非 Windows 平台，或 Windows 上未安装 Word 时使用。

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

