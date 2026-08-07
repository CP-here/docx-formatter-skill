/**
 * build_docx.js
 * Node.js template for generating Word (.docx) documents
 * with OMML math formulas and table-box diagrams.
 *
 * PREREQUISITE: npm install docx  (v8.5.0+)
 *
 * USAGE:
 *   1. Copy this file + mathHelpers.js + formulas.js to working directory
 *   2. Modify the CONTENT section below with your document text
 *   3. Modify OUTPUT_PATH to the desired output location
 *   4. Run: node build_docx.js
 *
 * This template demonstrates all key patterns:
 *   - Document title (centered) and heading styles (h1/h2, left-aligned)
 *   - Body paragraphs with first-line indent
 *   - Item paragraphs (bold label + body text)
 *   - Block math formulas (centered)
 *   - Inline math formulas (mixed with text)
 *   - Table-box diagrams (single box, two-column box, multi-cell row)
 *   - Arrow connectors (↓)
 *   - Figure captions
 *   - Page breaks
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  Table, TableRow, TableCell, WidthType, VerticalAlign, PageBreak,
  TableOfContents, TableLayoutType, BorderStyle,
} = require("docx");
const { eq1, eq2, eq3, eq4 } = require("./formulas");
const { r, sub, math, inlineMath } = require("./mathHelpers");

// ============================================================
// CONFIGURATION
// ============================================================
const FONT = "SimSun";        // 宋体 (body)
const HFONT = "SimHei";       // 黑体 (headings)
const OUTPUT_PATH = "OUTPUT.docx";  // TODO: Set output path

// ============================================================
// HELPER FUNCTIONS — Reusable building blocks
// ============================================================

/** Plain text run */
function t(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: 21,  // 10.5pt = 小四
    bold: !!opts.bold,
    italics: !!opts.italics,
  });
}

/**
 * Parse text into runs, handling **bold** and [n] citation superscripts.
 * Returns an array of TextRun objects.
 *
 * Citation pattern matches: [1], [1,2], [1-3], [1, 2, 3], etc.
 * These are rendered as superscript (e.g. for academic references).
 */
function parseRichText(text) {
  const runs = [];
  // Combined regex: **bold** OR [citation]
  const pattern = /(\*\*.*?\*\*|\[\d[\d,\-\s]*\])/g;
  let lastIndex = 0;
  let match;
  while ((match = pattern.exec(text)) !== null) {
    // Add preceding plain text
    if (match.index > lastIndex) {
      runs.push(t(text.slice(lastIndex, match.index)));
    }
    const token = match[0];
    if (token.startsWith('**') && token.endsWith('**')) {
      // Bold text
      runs.push(t(token.slice(2, -2), { bold: true }));
    } else {
      // Citation marker → superscript
      runs.push(new TextRun({
        text: token,
        font: FONT,
        size: 21,
        superScript: true,
      }));
    }
    lastIndex = match.index + token.length;
  }
  // Add remaining plain text
  if (lastIndex < text.length) {
    runs.push(t(text.slice(lastIndex)));
  }
  return runs;
}

/** Body paragraph (justified, first-line indent, 1.5x line spacing).
 *  Supports **bold** and [n] citation superscripts.
 *  Line spacing: 360/240 = 1.5x (aligned with NJUThesis linespread=1.625). */
function para(text, opts = {}) {
  const { spacingAfter = 140, align = AlignmentType.JUSTIFIED, indent = 480 } = opts;
  return new Paragraph({
    alignment: align,
    indent: indent ? { firstLine: indent } : undefined,
    spacing: { after: spacingAfter, line: 360 },
    children: parseRichText(text),
  });
}

/** Document title (黑体 三号 16pt, bold, centered) — for document main title only.
 *  Spacing: before=12pt, after=12pt, line=1.5x (aligned with NJUThesis). */
function title(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240, line: 360 },
    children: [new TextRun({ text, bold: true, font: HFONT, size: 32 })],  // 16pt = 32 half-points
  });
}

/** H1 heading (黑体 三号 16pt, bold, black, left-aligned) — for section headings.
 *  Uses Word built-in Heading 1 style for TOC compatibility.
 *  Spacing: before=10pt, after=24pt, line=1.5x (NJUThesis chapter, after reduced for Word). */
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 480, line: 360 },
    children: [new TextRun({ text, bold: true, font: HFONT, size: 32, color: "000000" })],  // 16pt = 32 half-points
  });
}

/** H2 heading (黑体 小四 12pt, bold, black) — for sub-section headings.
 *  Uses Word built-in Heading 2 style for TOC compatibility.
 *  Spacing: before=18pt, after=12pt, line=1.5x (aligned with NJUThesis section). */
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.LEFT,
    spacing: { before: 360, after: 240, line: 360 },
    children: [new TextRun({ text, bold: true, font: HFONT, size: 24, color: "000000" })],  // 12pt = 24 half-points
  });
}

/** Item paragraph: bold label + normal body text, first-line indent.
 *  Body text supports **bold** and [n] citation superscripts.
 *  Line spacing: 1.5x (aligned with NJUThesis linespread=1.625). */
function itemPara(label, text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { after: 140, line: 360 },
    children: [
      new TextRun({ text: label, bold: true, font: FONT, size: 21 }),
      ...parseRichText(text),
    ],
  });
}

/** Centered paragraph for block math formula.
 *  Line spacing: 1.5x (aligned with NJUThesis linespread=1.625). */
function eqPara(mathObj) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: 360 },
    children: [mathObj],
  });
}

/** Page break */
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ----- Table-box diagram helpers -----

/** Single centered box (for flowchart steps) */
function box(text, opts = {}) {
  const { width = 6000, fontSize = 20 } = opts;  // width in DXA (twentieths of a point)
  return new Table({
    width: { size: width, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: width, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 100, bottom: 100, left: 100, right: 100 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, font: FONT, size: fontSize, bold: true })],
        })],
      })],
    })],
  });
}

/** Multi-line centered box (for flowchart steps with multiple lines) */
function multiLineBox(lines, opts = {}) {
  const { width = 6000, fontSize = 20 } = opts;
  const cellChildren = lines.map((line, i) => new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: i < lines.length - 1 ? 30 : 0 },
    children: [new TextRun({ text: line, font: FONT, size: fontSize, bold: true })],
  }));
  return new Table({
    width: { size: width, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: width, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 100, bottom: 100, left: 100, right: 100 },
        children: cellChildren,
      })],
    })],
  });
}

/** Down arrow connector */
function arrowDown() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: "\u2193", font: FONT, size: 24, bold: true })],  // ↓
  });
}

// CM to DXA conversion (1 cm = 566.93 DXA)
const _CM_TO_DXA = 566.93;

/** Horizontal "process → output" row with arrow centered in the table.
 *  Uses a 3-column table (left | arrow | right) with equal-width left/right
 *  columns and fixed layout, so the arrow is always geometrically centered
 *  regardless of text length. The middle vertical border is removed (fused),
 *  making left and right cells appear seamlessly connected.
 *
 *  Shading logic:
 *    - leftShade / rightShade default undefined (white background)
 *    - Equal status (both undefined): both cells white
 *    - Different hierarchy: pass leftShade: "D9D9D9" for left gray, right white
 *
 *  @param {string} leftText  - text for the left box (process step)
 *  @param {string} rightText - text for the right box (output/result)
 *  @param {object} opts
 *  @param {number} opts.totalCm  - total table width in cm (default 14.0)
 *  @param {number} opts.arrowCm  - arrow column width in cm (default 2.0)
 *  @param {number} opts.fontSize - font size in half-points (default 21 = 10.5pt)
 *  @param {boolean} opts.leftBold  - left text bold (default true)
 *  @param {boolean} opts.rightBold - right text bold (default false)
 *  @param {string} opts.leftShade  - left cell bg color hex, or undefined for white
 *  @param {string} opts.rightShade - right cell bg color hex, or undefined for white
 *  @returns {Table} */
function arrowRow(leftText, rightText, opts = {}) {
  const {
    totalCm = 14.0,
    arrowCm = 2.0,
    fontSize = 21,        // 10.5pt = 21 half-points
    leftBold = true,
    rightBold = false,
    leftShade,
    rightShade,
  } = opts;

  const sideCm = (totalCm - arrowCm) / 2.0;
  const sideDxa = Math.round(sideCm * _CM_TO_DXA);
  const arrowDxa = Math.round(arrowCm * _CM_TO_DXA);

  const mkCell = (text, widthDxa, bold, shade) => new TableCell({
    width: { size: widthDxa, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
    shading: shade ? { fill: shade } : undefined,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, font: FONT, size: fontSize, bold })],
    })],
  });

  const arrowCell = new TableCell({
    width: { size: arrowDxa, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "\u2192", font: FONT, size: 28, bold: true })],  // →
    })],
  });

  return new Table({
    alignment: AlignmentType.CENTER,
    width: { size: Math.round(totalCm * _CM_TO_DXA), type: WidthType.DXA },
    columnWidths: [sideDxa, arrowDxa, sideDxa],
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "auto" },
      insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
    },
    rows: [new TableRow({
      children: [
        mkCell(leftText, sideDxa, leftBold, leftShade),
        arrowCell,
        mkCell(rightText, sideDxa, rightBold, rightShade),
      ],
    })],
  });
}

/** Centered dashed separator line with note text.
 *  Used in flowcharts to mark a boundary (e.g. '----- AI 介入止于此处 -----').
 *  @param {string} text - separator text
 *  @returns {Paragraph} */
function separatorNote(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 21 })],
  });
}

// ============================================================
// FIGURE/TABLE AUTO-NUMBERING
// ============================================================

let _figCounter = 0;
let _tblCounter = 0;

/** Reset figure and table counters to zero.
 *  Called automatically at the start of the CONTENT section below.
 *  Manual call only needed if generating multiple documents in one process. */
function resetCounters() {
  _figCounter = 0;
  _tblCounter = 0;
}

/** Figure caption (centered, below figure). Auto-increments figure number.
 *  Format: "图N 描述文字" (N starts at 1, auto-increments).
 *  Pass empty string "" for spacing-only caption (no number, no text). */
function figCaption(text) {
  let runs = [];
  if (text) {
    _figCounter++;
    runs = [new TextRun({ text: `图${_figCounter} ${text}`, font: FONT, size: 21 })];
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 300 },
    children: runs,
  });
}

/** Table caption (centered, above table). Auto-increments table number.
 *  Format: "表N 描述文字" (N starts at 1, auto-increments). */
function tableCaption(text) {
  _tblCounter++;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text: `表${_tblCounter} ${text}`, font: FONT, size: 21 })],
  });
}

/** Figure note (centered, italic, smaller font) */
function note(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 18, italics: true })],
  });
}

/**
 * Table of Contents (TOC) — on-demand, only call when document needs a TOC.
 * Returns an array of Paragraph/TableOfContents elements.
 *
 * Formatting (aligned with NJUThesis):
 *   - Title: 黑体 三号(16pt) 居中 加粗
 *   - TOC 1 (一级): 黑体 四号(14pt) 不加粗
 *   - TOC 2 (二级): 宋体 小四(12pt) 不加粗
 *   - Line spacing: 固定22磅
 *   - Levels: default '1-2' (H1 + H2 only)
 *
 * TOC entry styles (TOC1/TOC2) are pre-defined in the Document constructor.
 * After opening in Word, right-click the TOC area and select "更新域" to generate.
 *
 * @param {string} title - TOC title text (default '目  录')
 * @param {string} levels - heading levels to include (default '1-2' for H1+H2)
 * @returns {Array} Array of elements to push into children
 */
function toc(title = "目  录", levels = "1-2") {
  return [
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 240, line: 360 },
      children: [new TextRun({ text: title, bold: true, font: HFONT, size: 32, color: "000000" })],
    }),
    new TableOfContents("目录", {
      hyperlink: true,
      headingStyleRange: levels,
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

/** Two-column box (side-by-side comparison) */
function twoColBox(leftTitle, leftLines, rightTitle, rightLines, opts = {}) {
  const { width = 6200, fs = 17 } = opts;
  const half = Math.floor(width / 2);
  const mkCell = (title, lines) => new TableCell({
    width: { size: half, type: WidthType.DXA },
    margins: { top: 90, bottom: 90, left: 80, right: 80 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({ text: title, bold: true, font: FONT, size: fs })],
      }),
      ...lines.map(l => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 30 },
        children: [new TextRun({ text: l, font: FONT, size: fs - 1 })],
      })),
    ],
  });
  return new Table({
    alignment: AlignmentType.CENTER,
    width: { size: width, type: WidthType.DXA },
    columnWidths: [half, half],
    rows: [new TableRow({ children: [mkCell(leftTitle, leftLines), mkCell(rightTitle, rightLines)] })],
  });
}

/**
 * Multi-cell horizontal row (for showing parallel items).
 * cells: [{ title: "string", desc: "string" }, ...]
 */
function multiCellRow(cells, opts = {}) {
  const { cellWidth = 1300, fontSize = 15 } = opts;
  const totalWidth = cellWidth * cells.length;
  const tableCells = cells.map(c => new TableCell({
    width: { size: cellWidth, type: WidthType.DXA },
    margins: { top: 70, bottom: 70, left: 40, right: 40 },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 30 },
        children: [new TextRun({ text: c.title, bold: true, font: FONT, size: fontSize })],
      }),
      ...(c.desc ? [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: c.desc, font: FONT, size: fontSize - 1 })],
      })] : []),
    ],
  }));
  return new Table({
    alignment: AlignmentType.CENTER,
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: Array(cells.length).fill(cellWidth),
    rows: [new TableRow({ children: tableCells })],
  });
}

// ============================================================
// CONTENT — Modify this section for your document
// ============================================================
const children = [];
resetCounters();  // Auto-reset figure/table counters for each new document

// Example: Document structure with formulas and diagrams
// --- Replace with your actual content ---

children.push(
  title("文档标题（居中）"),
);

children.push(
  h1("一、概述"),
  para("在此填写文档概述内容"),
);

children.push(
  h1("二、技术背景"),
  para("在此填写技术背景描述"),
);

// Example: Block math formula
children.push(h1("三、方法描述"));
children.push(para("总损失函数为："));
children.push(eqPara(eq1));

// Example: Inline math formula mixed with text
children.push(new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  indent: { firstLine: 480 },
  spacing: { after: 140, line: 360 },
  children: [
    new TextRun({ text: "其中，", font: FONT, size: 21 }),
    inlineMath([sub("L", "LLM")]),
    new TextRun({ text: "为损失项，计算公式为：", font: FONT, size: 21 }),
  ],
}));
children.push(eqPara(eq2));

// Example: Table-box diagram (flowchart)
children.push(figCaption(""));
children.push(box("第一层（描述内容）", { fontSize: 18 }));
children.push(arrowDown());
children.push(box("第二层（描述内容）", { fontSize: 18 }));
children.push(arrowDown());
children.push(twoColBox(
  "左侧标题", ["条目1", "条目2"],
  "右侧标题", ["条目1", "条目2"],
));
children.push(figCaption("图1 示意图"));

// --- End of example content ---

// ============================================================
// DOCUMENT ASSEMBLY & EXPORT
// ============================================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    // TOC entry styles (aligned with NJUThesis LaTeX template)
    // TOC 1 (一级): 黑体 四号(14pt) 不加粗, 固定行距22磅
    // TOC 2 (二级): 宋体 小四(12pt) 不加粗, 固定行距22磅
    paragraphStyles: [
      {
        id: "TOC1",
        name: "TOC 1",
        basedOn: "Normal",
        next: "Normal",
        run: { font: HFONT, size: 28, bold: false },  // 黑体 14pt = 28 half-points
        paragraph: { spacing: { line: 440, lineRule: "exact" } },  // 22pt = 440 twips
      },
      {
        id: "TOC2",
        name: "TOC 2",
        basedOn: "Normal",
        next: "Normal",
        run: { font: FONT, size: 24, bold: false },   // 宋体 12pt = 24 half-points
        paragraph: { spacing: { line: 440, lineRule: "exact" } },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        // A4: 11907 × 16840 twips; margins: 2.54cm = 1440 twips
        size: { width: 11907, height: 16840 },
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync(OUTPUT_PATH, buf);
  console.log("done: " + OUTPUT_PATH);
});
