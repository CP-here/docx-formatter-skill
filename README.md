# Word文档一键排版 Skill：docx-formatter 

生成专业排版的 Word (.docx) 文档，支持 OMML 数学公式、引用上标、表格框图流程图、自动编号和文档验证。

## 项目简介

本项目包含两个独立的 Skill，可单独安装使用：

| Skill | 定位 | 适用场景 |
|-------|------|----------|
| **docx-formatter** | 轻量版 | 生成带公式/流程图的 Word 文档，文档完整性验证 |
| **docx-formatter-plus** | 增强版 | 在轻量版基础上增加 PDF 渲染验证、XSD 模式验证、编辑现有文档、追踪修订和批注 |

两个 Skill 均采用**双引擎架构**：

- **引擎 A（Python + python-docx）**：文本、表格、流程图，始终需要
- **引擎 B（Node.js + docx 库 v8.5.0+）**：OMML 数学公式，按需启用

## 核心特性

### 1. OMML 数学公式（引擎 B）

通过 `mathHelpers.js` 构建 Office 原生数学元素，Word 中可二次编辑，非图片或纯文本。

**解决的问题**：`MathSum` 使用空 `superScript: []` 会渲染不可见上标框导致文件损坏，本项目改用 `MathSubScript` + Unicode ∑ 彻底规避。

```javascript
const { r, sub, sup, frac, sumOp, func, math, inlineMath } = require("./mathHelpers");

// 下标：y_pred
sub("y", "pred")

// 上标：x³
sup("x", "3")

// 分式：η/N
frac([r("\u03b7")], [r("N")])

// 求和：Σ_k a_k（安全写法，无空上标）
sumOp([r("k")], [sub("a", "k")])

// 块级公式（居中，独立行）
math([r("f(t) = "), frac([sub("a", "0")], [r("2")]), r(" + "),
  sumOp([r("k")], [sub("a", "k"), func("cos", [r("k"), r("\u03c9t")])])])

// 行内公式（与文本混排在段落中）
inlineMath([sub("a", "k")])
```

渲染效果（Word 中显示为原生可编辑公式）：

> f(t) = a₀/2 + Σ_k a_k cos(kωt)

### 2. 引用上标（双引擎）

正文中的参考文献引用标记 `[1]`、`[1,2]`、`[1-3]` 自动渲染为上标，无需手动处理。

```python
# Python
add_body(doc, "该方法在实验中表现出色[1]。后续研究[2,3]进一步验证了这一结论。")
```

```javascript
// Node.js
para("该方法在实验中表现出色[1]。后续研究[2,3]进一步验证了这一结论。")
```

渲染效果：正文中的 `[1]` 和 `[2,3]` 显示为上标字符。

### 3. 表格框图流程图（双引擎）

使用 Word 表格 + Unicode 箭头绘制流程图，不依赖图片或 emoji。支持单框、多行框、并排框、横向箭头行和虚线分隔符。

**横向箭头行**（过程 → 产出）是本项目的创新设计：

- 3 列表格（左框 | → | 右框），左右等宽，箭头恒居中
- **融合中间竖线**（`insideV=none`），左右单元格视觉连通
- **可选底纹**：左右地位相等时全白，层级不同时左灰右白

```python
# Python：前两步全白，后两步左灰右白
add_arrow_row(doc, "① 时域采样", "离散信号 x(n)")
add_arrow_down(doc)
add_arrow_row(doc, "② 加窗处理", "加窗信号")
add_arrow_down(doc)
add_arrow_row(doc, "③ 离散傅里叶变换", "频谱 X(k)", left_shade='D9D9D9')
add_arrow_down(doc)
add_arrow_row(doc, "④ 频谱分析", "特征向量", left_shade='D9D9D9')
add_arrow_down(doc)
add_separator_note(doc, "----- 信号预处理止于此处，后续由分类器承接 -----")
add_arrow_down(doc)
add_box(doc, "分类决策：特征归一化 / 模式匹配 / 结果输出 / 置信度评估")
add_fig_caption(doc, "信号处理与分类流程图")
```

```javascript
// Node.js
children.push(arrowRow("① 时域采样", "离散信号 x(n)"));
children.push(arrowDown());
children.push(arrowRow("② 加窗处理", "加窗信号"));
children.push(arrowDown());
children.push(arrowRow("③ 离散傅里叶变换", "频谱 X(k)", { leftShade: "D9D9D9" }));
children.push(arrowDown());
children.push(arrowRow("④ 频谱分析", "特征向量", { leftShade: "D9D9D9" }));
children.push(arrowDown());
children.push(separatorNote("----- 信号预处理止于此处，后续由分类器承接 -----"));
children.push(arrowDown());
children.push(box("分类决策：特征归一化 / 模式匹配 / 结果输出 / 置信度评估"));
children.push(figCaption("信号处理与分类流程图"));
```

渲染效果（Word 表格形式）：

```
┌─────────────┬───┬─────────────┐
│ ① 时域采样  │ → │ 离散信号x(n)│  ← 全白
└─────────────┴───┴─────────────┘
         ↓
┌─────────────┬───┬─────────────┐
│ ② 加窗处理  │ → │  加窗信号   │  ← 全白
└─────────────┴───┴─────────────┘
         ↓
┌──────────────┬───┬───────────┐
│▓③ 离散傅里叶▓│ → │ 频谱 X(k) │  ← 左灰右白
└──────────────┴───┴───────────┘
         ↓
┌─────────────┬───┬─────────────┐
│▓④ 频谱分析▓│ → │  特征向量   │  ← 左灰右白
└─────────────┴───┴─────────────┘
    ----- 信号预处理止于此处 -----
         ↓
┌──────────────────────────────────┐
│ 分类决策：特征归一化 / 模式匹配...│
└──────────────────────────────────┘
          图1 信号处理与分类流程图
```

### 4. 图表自动编号（双引擎）

图标题和表标题自动递增编号（图1, 图2... / 表1, 表2...），计数器在文档初始化时自动重置。

```python
add_box(doc, "系统架构图")
add_fig_caption(doc, "系统架构")    # → "图1 系统架构"

add_table_caption(doc, "参数对比")  # → "表1 参数对比"
```

### 5. 目录（TOC）— 按需生成

使用 Word 内置 Heading 样式 + TOC 域，支持 Ctrl+点击超链接跳转。标题颜色显式覆盖为黑色（避免 Word 模板默认蓝色）。目录条目样式对齐 NJUThesis LaTeX 模板。

```python
doc = setup_document()
add_title(doc, "文档标题")
add_toc(doc)                    # 在标题后、正文前插入目录
add_h1(doc, "一、概述")
add_body(doc, "正文内容...")
```

### 6. 文档验证（docx-formatter 起）

纯 Python 标准库实现的轻量验证工具，无外部依赖，执行 5 项结构检查：

| 检查项 | 说明 |
|--------|------|
| ZIP 完整性 | 所有 ZIP 条目均可解压 |
| XML 格式良好性 | 所有 .xml/.rels 文件可被 ElementTree 解析 |
| 文件引用完整性 | .rels 中每个 Target 在包内存在 |
| 内容类型声明 | word/ 下的 XML 文件在 [Content_Types].xml 中有声明 |
| 空白保留 | 含首尾空白的 w:t 元素必须有 xml:space="preserve" |

```bash
python validate_docx.py output.docx --verbose
```

### 7. 安全处理规范

`safe_extract()` 函数在解压 .docx 时防止路径遍历和符号链接攻击，`rezip()` 确保重新打包的文件符合 OOXML 规范（`[Content_Types].xml` 首位存储 + 原子写入）。

## 排版标准

所有间距值对齐 NJUThesis LaTeX 模板（`linespread = 1.625`，等效 Word 1.5 倍行距）。

| 元素 | 字体 | 字号 | 样式 | 段前 | 段后 |
|---------|------|------|-------|-------------|-------------|
| 文档标题 | 黑体 | 16pt | 加粗，居中 | 12pt | 12pt |
| H1（章） | 黑体 | 16pt | 加粗，左对齐 | 10pt | 24pt |
| H2（节） | 黑体 | 12pt | 加粗，左对齐 | 18pt | 12pt |
| H3（小节） | 黑体 | 11pt | 加粗，左对齐 | 14pt | 8pt |
| 正文 | 宋体 | 12pt | 两端对齐，首行缩进2字符 | 0 | 7pt |
| 块级公式 | — | — | 居中 | 6pt | 6pt |
| 图标题 | 宋体 | 10.5pt | 居中，图下方 | 6pt | 12pt |
| 表标题 | 宋体 | 10.5pt | 居中，表上方 | 12pt | 6pt |

## 文件结构

### docx-formatter（轻量版）

```
docx-formatter/
├── SKILL.md                     # Skill 指令文件
└── scripts/
    ├── build_docx.js            # Node.js 文档构建模板（含 OMML 数学公式）
    ├── build_docx.py            # Python 文档构建模板（文本 + 图表 + safe_extract/rezip）
    ├── mathHelpers.js           # OMML 数学元素构建器
    ├── formulas.js              # 公式定义模板
    └── validate_docx.py         # 轻量验证脚本（纯 stdlib，无外部依赖）
```

### docx-formatter-plus（增强版）

```
docx-formatter-plus/
├── SKILL.md                          # Skill 指令文件
└── scripts/
    ├── build_docx.js                 # Node.js 文档构建模板
    ├── build_docx.py                 # Python 文档构建模板
    ├── mathHelpers.js                # OMML 数学元素构建器
    ├── formulas.js                   # 公式定义模板
    ├── merge_runs.py                 # 合并碎片 run（编辑现有文档前置步骤）
    ├── accept_changes.py             # 接受所有追踪修订
    ├── comment.py                    # 批注管理（6 文件交叉链接系统）
    ├── templates/                    # 批注 XML 模板
    │   ├── comments.xml
    │   ├── commentsExtended.xml
    │   ├── commentsExtensible.xml
    │   ├── commentsIds.xml
    │   └── people.xml
    └── office/                       # 验证与转换工具
        ├── soffice.py                # LibreOffice 跨平台调用
        ├── validate.py               # XSD 模式验证入口
        ├── helpers/                  # 通用辅助函数包
        ├── schemas/                  # OOXML XSD 模式文件
        └── validators/               # 验证器（docx/pptx/redlining）
```

## 环境要求

### 必需（引擎 A：Python）

- Python 3.8+
- python-docx (`pip install python-docx`)

### 按需（引擎 B：Node.js，数学公式时需要）

- Node.js 18+
- docx 库 v8.5.0+ (`npm install docx`)

### 可选（仅 docx-formatter-plus）

- LibreOffice（PDF 渲染验证）
- lxml（XSD 模式验证）

## 安装方式

### 通过客户端安装

将 skill ZIP 上传至 Codex 或 任意AI客户端对应的 Skill 安装位置即可。Agent 会自动读取 `SKILL.md` 中的 frontmatter 和指令。

### 手动安装

将 skill 文件夹放入项目的 `./skills/` 目录：

```
你的项目/
└── ./skills/
    ├── docx-formatter/
    │   ├── SKILL.md
    │   └── scripts/
    └── docx-formatter-plus/
        ├── SKILL.md
        └── scripts/
```

## 使用方式

安装后，在对话中直接描述需求即可触发 Skill：

- "帮我生成一份带公式的 Word 文档"
- "排版一份技术报告为 Word"
- "用数学方程格式化这段文档"
- "在 Word 中创建流程图"
- "验证这个 docx 文件结构"

Agent 会根据 `SKILL.md` 中的触发条件自动选择合适的引擎和函数。

## 完整样例

以下样例展示一个包含公式、引用上标、流程图和目录的完整文档生成过程：

**Python（引擎 A）：**
```python
from build_docx import *

doc = setup_document()
add_title(doc, "实验报告")
add_toc(doc)

add_h1(doc, "一、实验目的")
add_body(doc, "本实验旨在验证傅里叶分析方法的有效性。如文献[1]所述，**傅里叶分析**在多种信号处理任务中表现优异。后续研究[2,3]进一步扩展了其应用范围。")

add_h1(doc, "二、方法")
add_h2(doc, "2.1 傅里叶级数定义")
add_body(doc, "周期信号的傅里叶级数展开定义为")
# 注意：数学公式需要引擎 B（Node.js）

add_h1(doc, "三、实验流程")
add_arrow_row(doc, "① 时域采样", "离散信号 x(n)")
add_arrow_down(doc)
add_arrow_row(doc, "② 加窗处理", "加窗信号", left_shade='D9D9D9')
add_arrow_down(doc)
add_arrow_row(doc, "③ 离散傅里叶变换", "频谱 X(k)", left_shade='D9D9D9')
add_separator_note(doc, "----- 预处理止于此处 -----")
add_arrow_down(doc)
add_box(doc, "频谱分析：峰值检测 / 频带能量 / 谱质心提取")
add_fig_caption(doc, "信号处理流程图")

doc.save("report.docx")
```

**Node.js（引擎 B，含数学公式）：**
```javascript
const { title, h1, h2, para, eqPara, arrowRow, arrowDown,
        box, figCaption, separatorNote, toc, pageBreak } = require("./build_docx");
const { sub, sup, frac, sumOp, func, math, r } = require("./mathHelpers");

children.push(title("实验报告"));
children.push(...toc());

children.push(h1("一、实验目的"));
children.push(para("本实验旨在验证傅里叶分析方法的有效性。如文献[1]所述，**傅里叶分析**在多种信号处理任务中表现优异。后续研究[2,3]进一步扩展了其应用范围。"));

children.push(h1("二、方法"));
children.push(h2("2.1 傅里叶级数定义"));
children.push(para("周期信号的傅里叶级数展开定义为："));
// 块级公式：f(t) = a₀/2 + Σ_k a_k cos(kωt)
children.push(eqPara(math([
  r("f(t) = "), frac([sub("a", "0")], [r("2")]), r(" + "),
  sumOp([r("k")], [sub("a", "k"), func("cos", [r("k"), r("\u03c9t")])])
])));
// 块级公式：DFT X(k) = Σ_n x(n)·e^{-j2πkn/N}
children.push(eqPara(math([
  r("X(k) = "),
  sumOp([r("n")], [
    r("x(n) \u00b7 "),
    sup(r("e"), r("\u2212j2\u03c0kn/N")),
  ]),
])));

children.push(h1("三、实验流程"));
children.push(arrowRow("① 时域采样", "离散信号 x(n)"));
children.push(arrowDown());
children.push(arrowRow("② 加窗处理", "加窗信号", { leftShade: "D9D9D9" }));
children.push(arrowDown());
children.push(arrowRow("③ 离散傅里叶变换", "频谱 X(k)", { leftShade: "D9D9D9" }));
children.push(separatorNote("----- 预处理止于此处 -----"));
children.push(arrowDown());
children.push(box("频谱分析：峰值检测 / 频带能量 / 谱质心提取"));
children.push(figCaption("信号处理流程图"));
```

## 技术亮点

### 求和符号安全渲染

`MathSum` 使用空 `superScript: []` 会在 Word 中渲染一个不可见的上标框，导致文件看起来损坏。本项目改用 `MathSubScript` 配合 Unicode ∑（U+2211），彻底规避此问题。

```javascript
// 错误 — 渲染空上标，可能导致文件损坏
new MathSum({ subScript: [r("i")], superScript: [], children: [...] })

// 正确 — 无上标，干净渲染
new MathSubScript({ children: [r("∑")], subScript: [r("i")] })
```

### 数组展平防 XML 损坏

`sumOp()` 返回数组 `[sumSymbol, ...body]`，如果不展平，嵌套数组会导致 XML 损坏，.docx 文件无法打开。所有接收 children 的函数均通过 `flat()` 处理。

### 横向箭头行融合边框

通过 `tblBorders` 的 `insideV=none` 去掉左/箭头/右之间的竖线，外框和 `insideH` 保留，实现左右单元格视觉连通的效果。同时通过 `tblLayout=fixed` 锁定列宽，防止 Word autofit 重排窄箭头列。

### TOC 样式覆盖

Word 更新 TOC 域时会套用模板内置的 TOC 样式（默认加粗）。本项目在样式定义中移除 `customStyle="1"` 属性并设置 `basedOn="Normal"`，确保 Word 使用自定义格式而非模板默认值。

## 版本对比

| 特性 | docx-formatter | docx-formatter-plus |
|------|:-:|:-:|
| 标准排版（标题/正文/缩进/行距） | ✓ | ✓ |
| OMML 数学公式 | ✓ | ✓ |
| 引用上标 | ✓ | ✓ |
| 表格框图流程图 | ✓ | ✓ |
| 图表自动编号 | ✓ | ✓ |
| 目录（TOC） | ✓ | ✓ |
| 轻量文档验证（纯 stdlib） | ✓ | ✓ |
| 安全解压/重打包 | ✓ | ✓ |
| PDF 渲染验证（LibreOffice） | — | ✓ |
| XSD 模式验证 | — | ✓ |
| 编辑现有文档 | — | ✓ |
| 追踪修订 | — | ✓ |
| 批注管理 | — | ✓ |

## 许可证

本项目可自由使用和修改。

## 致谢

- 排版标准对齐 [NJUThesis](https://github.com/nju-lug/NJUThesis) LaTeX 模板
- 增强版验证与编辑功能参考 [Anthropic 官方 docx skill](https://github.com/anthropics/skills/tree/main/skills/docx)
