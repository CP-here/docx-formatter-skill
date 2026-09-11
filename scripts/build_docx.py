# -*- coding: utf-8 -*-
"""
build_docx.py
Python template for generating Word (.docx) documents.
Use this engine when NO OMML math formulas are needed (only text, tables, diagrams).

PREREQUISITE: pip install python-docx

USAGE:
  1. Copy this file to working directory
  2. Modify the main() function with your content
  3. Modify OUTPUT_PATH
  4. Run: python build_docx.py

Features:
  - Typography presets: PRESETS + set_preset('name') whole-set switching —
    headings/body fonts & spacing, 中西文字体分离 (west fonts), caption/note
    sizes, page margins, TOC styles
  - 黑体三号(16pt) Title centered, H1/H2/H3 left-aligned
  - 宋体小四(12pt) body, first-line indent 2 chars, 1.5 line spacing
  - Table-box diagrams: add_box(), add_multi_line_box(), add_multi_col_table(), add_arrow_down()
  - Layered architecture diagrams: add_layered_architecture(doc, layers)
  - Single width source: every table is built by _new_table(), which writes
    w:tblW + w:tblGrid + w:tcW from one column-width list
  - Cell margins (tcMar) for all table-box elements
  - Figure notes: add_note() for italic annotations below figures
  - Table Grid style with header shading (D9D9D9, gray)
  - **bold** text formatting support in body text
"""

from docx import Document
from docx.shared import Pt, Cm, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import re
import zipfile


# ============================================================
# TYPOGRAPHY PRESETS — 排版参数集中定义（嵌套分组）
# ============================================================
# 覆盖分组：
#   - west        中西文字体分离（开关 + 西文正文/标题字体）
#   - title~h3    标题字体/字号/加粗/颜色/行距/段前后
#   - body/item   正文与分点段落（字体/字号/行距/缩进/段前后；None=不设置）
#   - caption/note 图表标题与图注字号
#   - page        纸张尺寸 + 页边距 + 页眉/页脚距离
#   - toc         目录标题/条目样式与固定行距
# 不覆盖：代码块、数据表、框图、箭头、公式段（保持函数内硬编码）。
# 切换方式：set_preset('name') 整套切换；须在 setup_document() 之前调用
#           （页面边距与 Normal 默认字体在 setup_document() 时读取预设）。
# 新增预设：向 PRESETS 添加一份同键结构的嵌套字典即可。

PRESETS = {
    'default': {
        'west': {
            'separate': True,            # 中西文分离（对齐 NJUThesis：正文西文衬线、标题西文无衬线）
            'body': 'Times New Roman',    # 西文正文字体（衬线）
            'head': 'Arial',             # 西文标题字体（无衬线，对齐 NJUThesis \sffamily）
        },
        'title': {'font': '黑体', 'size': 26, 'bold': True, 'color': None,
                  'line_spacing': 1.5, 'space_before': 24, 'space_after': 18},
        'h1':    {'font': '黑体', 'size': 16, 'bold': True, 'color': (0, 0, 0),
                  'line_spacing': 1.5, 'space_before': 24, 'space_after': 6},
        'h2':    {'font': '黑体', 'size': 12, 'bold': True, 'color': (0, 0, 0),
                  'line_spacing': 1.5, 'space_before': 12, 'space_after': 6},
        'h3':    {'font': '黑体', 'size': 12, 'bold': True, 'color': (0, 0, 0),
                  'line_spacing': 1.5, 'space_before': 12, 'space_after': 6},
        'body':  {'font': '宋体', 'size': 12, 'line_spacing': 1.5,
                  'first_line_indent': 24,
                  # None = 不设置，沿用样式继承（原实现正文不另设段间距）
                  'space_before': None, 'space_after': None},
        'item':  {'space_before': None, 'space_after': 7},
        'caption': {'size': 10.5, 'label_bold': True},   # 图/表标题字号与标签加粗（对齐 NJUThesis njucap）
        'note':  {'size': 9},         # 图注字号（斜体小字）
        'page': {
            'page_width': 21.0, 'page_height': 29.7,      # A4（21 × 29.7 cm）
            'margin_top': 2.54, 'margin_bottom': 2.54,
            'margin_left': 3.18, 'margin_right': 3.18,
            'header_distance': 1.27, 'footer_distance': 1.27,
        },
        'toc': {
            'title_font': '黑体', 'title_size': 16,
            'toc1_font': '黑体', 'toc1_size': 14,
            'toc2_font': '宋体', 'toc2_size': 12,
            'line_spacing': 22,      # 目录固定行距（pt）
        },
    },
}

_ACTIVE_PRESET = PRESETS['default']


def set_preset(name):
    """Switch the active typography preset (whole-set switching).

    Args:
        name: key in PRESETS, e.g. set_preset('default').
    Raises:
        KeyError: if the preset name is unknown.

    Note: call BEFORE setup_document() — page margins and the Normal
    default font are read from the preset when setup_document() runs.
    """
    global _ACTIVE_PRESET
    if name not in PRESETS:
        raise KeyError(f"unknown preset: {name!r} (available: {', '.join(PRESETS)})")
    _ACTIVE_PRESET = PRESETS[name]


def _west_font(kind):
    """Western font for 'body'/'head' when 中西文分离 is enabled, else None."""
    w = _ACTIVE_PRESET.get('west')
    if w and w.get('separate'):
        return w.get(kind)
    return None


def _preset_color(rgb_tuple):
    """Convert a preset (r, g, b) tuple to RGBColor, or None."""
    return RGBColor(*rgb_tuple) if rgb_tuple else None


# ============================================================
# FONT & STYLE HELPERS
# ============================================================

def set_run_font(run, font_name, size, bold=False, color=None, west_font=None):
    """Set font name, size, bold for a run, including East Asian font.

    west_font: optional Western font for ascii/hAnsi. When given, Western
    characters render in west_font while East Asian characters keep
    font_name (中西文分离).
    """
    run.font.name = west_font if west_font else font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


def add_title(doc, text):
    """Document title (for document main title only), centered, preset-driven.
    Default preset: 黑体 三号(16pt) 居中 加粗, 1.5x line spacing,
    space before/after 12pt (aligned with NJUThesis linespread=1.625).
    """
    s = _ACTIVE_PRESET['title']
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    run = p.add_run(text)
    set_run_font(run, s['font'], s['size'], bold=s['bold'],
                 color=_preset_color(s['color']), west_font=_west_font('head'))


def add_h1(doc, text):
    """H1 (section heading), left-aligned, preset-driven.
    Uses Word built-in Heading 1 style for TOC compatibility.
    Default preset: 黑体 16pt bold black, 1.5x, before=10pt, after=24pt
    (NJUThesis chapter, after reduced for Word).
    """
    s = _ACTIVE_PRESET['h1']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    run = p.add_run(text)
    set_run_font(run, s['font'], s['size'], bold=s['bold'],
                 color=_preset_color(s['color']), west_font=_west_font('head'))


def add_h2(doc, text):
    """H2, left-aligned, preset-driven.
    Uses Word built-in Heading 2 style for TOC compatibility.
    Default preset: 黑体 12pt bold black, 1.5x, before=18pt, after=12pt
    (aligned with NJUThesis section).
    """
    s = _ACTIVE_PRESET['h2']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 2']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    run = p.add_run(text)
    set_run_font(run, s['font'], s['size'], bold=s['bold'],
                 color=_preset_color(s['color']), west_font=_west_font('head'))


def add_h3(doc, text):
    """H3, left-aligned, preset-driven.
    Uses Word built-in Heading 3 style for TOC compatibility.
    Default preset: 黑体 11pt bold black, 1.5x, before=14pt, after=8pt
    (aligned with NJUThesis subsection).
    """
    s = _ACTIVE_PRESET['h3']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 3']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    run = p.add_run(text)
    set_run_font(run, s['font'], s['size'], bold=s['bold'],
                 color=_preset_color(s['color']), west_font=_west_font('head'))


def add_body(doc, text):
    """Body paragraph, preset-driven (default: 宋体 12pt, justified,
    first-line indent 2 chars, 1.5x line spacing).

    Supports:
      - **bold** segments
      - [n] citation markers → rendered as superscript (e.g. [1], [1,2], [1-3])
    """
    s = _ACTIVE_PRESET['body']
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = s['line_spacing']
    pf.first_line_indent = Pt(s['first_line_indent'])
    if s['space_before'] is not None:
        pf.space_before = Pt(s['space_before'])
    if s['space_after'] is not None:
        pf.space_after = Pt(s['space_after'])
    _add_runs_with_formatting(p, text)


def add_item_para(doc, label, text):
    """Item paragraph: bold label + normal body, first-line indent.

    Body text supports **bold** and [n] citation superscripts.
    Default preset: body style + space_after 7pt.
    """
    s = _ACTIVE_PRESET['body']
    it = _ACTIVE_PRESET['item']
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = s['line_spacing']
    pf.first_line_indent = Pt(s['first_line_indent'])
    if it['space_before'] is not None:
        pf.space_before = Pt(it['space_before'])
    pf.space_after = Pt(it['space_after'])
    run = p.add_run(label)
    set_run_font(run, s['font'], s['size'], bold=True,
                 west_font=_west_font('body'))
    _add_runs_with_formatting(p, text)


def _add_runs_with_formatting(p, text):
    """Add runs to paragraph p, parsing **bold** and [n] citation superscripts.

    Citation pattern matches: [1], [1,2], [1-3], [1, 2, 3], etc.
    Fonts read from the active body preset.
    """
    s = _ACTIVE_PRESET['body']
    # Combined regex: **bold** OR [citation]
    # Citation: [digits, possibly with commas/hyphens/spaces]
    pattern = r'(\*\*.*?\*\*|\[\d[\d,\-\s]*\])'
    parts = re.split(pattern, text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            # Bold text
            run = p.add_run(part[2:-2])
            set_run_font(run, s['font'], s['size'], bold=True,
                         west_font=_west_font('body'))
        elif re.match(r'^\[\d[\d,\-\s]*\]$', part):
            # Citation marker → superscript
            run = p.add_run(part)
            set_run_font(run, s['font'], s['size'], bold=False,
                         west_font=_west_font('body'))
            run.font.superscript = True
        else:
            run = p.add_run(part)
            set_run_font(run, s['font'], s['size'], bold=False,
                         west_font=_west_font('body'))


# ============================================================
# OMML MATH HELPERS — formulas via mathHelpers.py
# ============================================================

_M_NS_DECL = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def _insert_omml(p, omml_xml):
    """Parse an OMML XML string (from mathHelpers) and append it to paragraph p.

    mathHelpers.py is imported lazily so build_docx.py stays standalone
    (copyable without mathHelpers.py) for math-free documents.
    """
    from docx.oxml import parse_xml  # noqa: local import, same package as OxmlElement
    if "xmlns:m=" not in omml_xml:
        omml_xml = omml_xml.replace("<m:oMath>", "<m:oMath %s>" % _M_NS_DECL, 1)
    p._p.append(parse_xml(omml_xml))


def add_eq_para(doc, math_xml):
    """Centered block math formula paragraph.

    Args:
        math_xml: OMML XML string from mathHelpers.math(), e.g.
            from mathHelpers import r, sub, sumOp, func, math
            eq = math([sub("L", "LLM"), r(" = - "),
                       sumOp([r("i")], [sub("y", "i")])])
            add_eq_para(doc, eq)
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    _insert_omml(p, math_xml)
    return p


def add_body_with_math(doc, parts):
    """Body paragraph mixing text and inline math (行内公式与正文混排).

    Args:
        parts: list of ("text", str) or ("math", omml_xml) tuples, in order.
            Text parts support **bold** and [n] citation superscripts.

    Example:
        from mathHelpers import sub, inlineMath
        add_body_with_math(doc, [
            ("text", "其中，"),
            ("math", inlineMath([sub("L", "LLM")])),
            ("text", "为语言模型损失项。"),
        ])
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    s = _ACTIVE_PRESET['body']
    pf.line_spacing = s['line_spacing']
    pf.first_line_indent = Pt(s['first_line_indent'])
    if s['space_before'] is not None:
        pf.space_before = Pt(s['space_before'])
    if s['space_after'] is not None:
        pf.space_after = Pt(s['space_after'])
    for kind, val in parts:
        if kind == "math":
            _insert_omml(p, val)
        else:
            _add_runs_with_formatting(p, val)
    return p


# ============================================================
# CODE BLOCK / DATA TABLE / MATH-IN-CELL HELPERS
# ============================================================

def add_code_block(doc, code, font_size=9):
    """Code block: monospace (Consolas) lines with light-gray paragraph shading.

    Args:
        code: source code string; each line becomes one paragraph.
        font_size: code font size in pt (default 9).
    """
    for line in code.split("\n"):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(18)
        pf.line_spacing = 1.15
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F2F2F2')
        p._p.get_or_add_pPr().insert_element_before(
            shd, 'w:tabs', 'w:spacing', 'w:ind', 'w:jc')
        run = p.add_run(line if line.strip() else " ")
        run.font.name = 'Consolas'
        run.font.size = Pt(font_size)
        rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
        rfonts.set(qn('w:eastAsia'), '宋体')
    return doc


def add_data_table(doc, headers, rows, col_widths, font_size=9.5):
    """Data table with gray (D9D9D9) header row and fixed column widths.

    Header row: 黑体 bold, centered. Data rows: 宋体, centered except the
    last column (left-aligned, typically the description column).

    Args:
        headers: list of header strings.
        rows: list of row lists (same length as headers).
        col_widths: column widths in cm, same length as headers; their sum is
                    the table's total width (keep it ≤ the 版心宽).
        font_size: cell font size in pt (default 9.5).
    """
    table = _new_table(doc, 1 + len(rows), col_widths, margins=60)
    # Header
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        set_cell_shading(cell, 'D9D9D9')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        set_run_font(p.add_run(h), '黑体', font_size, bold=True)
    # Data rows
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            p = cell.paragraphs[0]
            p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if j < len(row) - 1
                           else WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.first_line_indent = Pt(0)
            set_run_font(p.add_run(str(val)), '宋体', font_size, bold=False)
    return table


def add_math_to_cell(cell, omml_xml):
    """Insert an inline OMML formula (mathHelpers.inlineMath) into a table
    cell, centered. The cell's first paragraph is used."""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    _insert_omml(p, omml_xml)
    return p


# ============================================================
# TABLE HELPERS
# ============================================================

def set_table_border(table):
    """Set all borders (top/left/bottom/right/insideH/insideV) to single black."""
    tbl = table._element
    tblPr = tbl.tblPr if tbl.tblPr is not None else tbl.makeelement(qn('w:tblPr'), {})
    borders = tblPr.makeelement(qn('w:tblBorders'), {})
    for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = borders.makeelement(qn(f'w:{edge}'), {
            qn('w:val'): 'single',
            qn('w:sz'): '4',
            qn('w:space'): '0',
            qn('w:color'): '000000'
        })
        borders.append(border)
    tblPr.insert_element_before(
        borders, 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook')


def set_table_border_no_insideV(table):
    """Set outer borders + insideH to single black, but insideV to none.

    This fuses the middle vertical border between adjacent cells, making
    left and right cells appear seamlessly connected while keeping
    individual outer borders intact.
    """
    tbl = table._element
    tblPr = tbl.tblPr if tbl.tblPr is not None else tbl.makeelement(qn('w:tblPr'), {})
    # Remove existing borders if any
    old_borders = tblPr.find(qn('w:tblBorders'))
    if old_borders is not None:
        tblPr.remove(old_borders)
    borders = OxmlElement('w:tblBorders')
    for edge, val in [('top', 'single'), ('left', 'single'), ('bottom', 'single'),
                       ('right', 'single'), ('insideH', 'single'), ('insideV', 'none')]:
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        borders.append(border)
    tblPr.insert_element_before(
        borders, 'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook')


def set_cell_shading(cell, color_hex):
    """Set cell background color (e.g., 'D9D9D9' for gray)."""
    shading = cell._element.get_or_add_tcPr().makeelement(qn('w:shd'), {
        qn('w:val'): 'clear',
        qn('w:color'): 'auto',
        qn('w:fill'): color_hex,
    })
    cell._element.get_or_add_tcPr().insert_element_before(
        shading, 'w:noWrap', 'w:tcMar', 'w:textDirection', 'w:tcFitText',
        'w:vAlign', 'w:hideMark')


# ============================================================
# TABLE-BOX DIAGRAM HELPERS (for flowcharts without images)
# ============================================================

# 框图/表格总宽的唯一来源（cm）。所有框图函数的默认宽度都取它；
# 单张表格的列宽之和即该表总宽，由 _new_table() 一次性写入
# w:tblW + w:tblGrid + w:tcW 三处，宽度不存在第二个口径。
# A4 + 左右边距各 3.18cm 的版心宽为 14.64cm，可按需传入更大值。
TABLE_WIDTH_CM = 14.0


def _set_cell_margins(cell, margin_dxa=80):
    """Set cell internal margins (top/bottom/left/right) in dxa units."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.makeelement(qn('w:tcMar'), {})
    # CT_TcMar order: top, start, left, bottom, end, right
    for side in ['top', 'left', 'bottom', 'right']:
        mar = tcMar.makeelement(qn(f'w:{side}'), {qn('w:w'): str(margin_dxa), qn('w:type'): 'dxa'})
        tcMar.append(mar)
    tcPr.insert_element_before(
        tcMar, 'w:textDirection', 'w:tcFitText', 'w:vAlign', 'w:hideMark')


def _set_table_width(table, total_twips):
    """Write the table's total width into w:tblW as an explicit dxa length.

    `total_twips` is the exact sum of the columns' w:gridCol twips, so w:tblW,
    w:tblGrid and w:tcW agree down to the twip. python-docx creates every table
    with ``<w:tblW w:type="auto" w:w="0"/>`` — "let the renderer decide", which
    is the autofit switch — so replacing it with a dxa length pins the total
    width independently of the grid and of any renderer's autofit.
    """
    tblPr = table._element.tblPr
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert_element_before(
            tblW, 'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders',
            'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook')
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(int(total_twips)))


def _new_table(doc, row_count, col_widths_cm, margins=60, borders=None):
    """Create a table whose width has exactly ONE source: `col_widths_cm`.

    The three places a table stores its geometry — the total width (w:tblW),
    the column grid (w:tblGrid/gridCol) and each cell width (w:tcW) — are all
    written from the same list, and the layout is locked to fixed. Callers only
    fill in cell content; they must not set widths themselves.

    Args:
        doc: the Document to add the table to.
        row_count: number of rows to create.
        col_widths_cm: list of column widths in cm; their sum is the table width.
        margins: cell internal margins in dxa (default 60).
        borders: border helper applied to the table (default set_table_border).
    """
    borders = borders or set_table_border
    # 总宽 = 列宽之和，只取整一次；各列按四舍五入取 twips，末列用差值补齐，
    # 再以 Twips() 原样写回 gridCol 与 tcW。于是 ΣgridCol == ΣtcW == tblW
    # 严格相等，且同一总宽无论分成几列，表宽都是同一个数——全宽层与并列层
    # 不可能差出 1 twip（不要用 Cm() 反算，它在 twips→cm→EMU 时会掉半个 twip）。
    total_twips = Cm(sum(col_widths_cm)).twips
    col_twips = [Cm(w).twips for w in col_widths_cm]
    col_twips[-1] = total_twips - sum(col_twips[:-1])
    table = doc.add_table(rows=row_count, cols=len(col_widths_cm))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    _set_table_width(table, total_twips)
    borders(table)
    _set_table_fixed_layout(table)
    for j, tw in enumerate(col_twips):
        table.columns[j].width = Twips(tw)      # → w:tblGrid/w:gridCol
        for row in table.rows:
            cell = row.cells[j]
            cell.width = Twips(tw)              # → w:tcW
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, margins)
    return table


def add_box(doc, text, width_cm=TABLE_WIDTH_CM, font_size=10.5, bold=True):
    """Single centered box for flowchart steps."""
    table = _new_table(doc, 1, [width_cm], margins=80)
    p = table.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p.add_run(text), '宋体', font_size, bold=bold)
    return table


def add_multi_line_box(doc, lines, width_cm=TABLE_WIDTH_CM, font_size=10.5):
    """Multi-line centered box for flowchart steps.

    lines: list of strings, each rendered as a separate centered line.
    """
    table = _new_table(doc, 1, [width_cm], margins=80)
    cell = table.cell(0, 0)
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        set_run_font(p.add_run(line), '宋体', font_size, bold=True)
    return table


def add_multi_col_table(doc, cells, col_width_cm=None, font_size=10.5):
    """
    Side-by-side boxes in a single row.
    cells: list of strings or list of (title, desc) tuples.
    col_width_cm: list of column widths in cm; None distributes
                  TABLE_WIDTH_CM evenly across the columns.
    """
    if col_width_cm is None:
        col_width_cm = [TABLE_WIDTH_CM / len(cells)] * len(cells)

    table = _new_table(doc, 1, col_width_cm, margins=60)
    for i, cell_data in enumerate(cells):
        cell = table.cell(0, i)

        if isinstance(cell_data, tuple):
            title, desc = cell_data
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(title)
            set_run_font(run, '宋体', font_size, bold=True)
            if desc:
                p2 = cell.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.first_line_indent = Pt(0)
                run = p2.add_run(desc)
                set_run_font(run, '宋体', font_size - 1, bold=False)
        else:
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(cell_data)
            set_run_font(run, '宋体', font_size, bold=True)

    return table


def add_arrow_down(doc):
    """Centered down arrow (↓) between flowchart boxes."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run('\u2193')  # ↓
    set_run_font(run, '宋体', 16, bold=True)


def add_arrow_horizontal(doc, text='\u2192'):
    """Horizontal arrow (→) for inline flow."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(text)
    set_run_font(run, '宋体', 16, bold=True)


def _set_table_fixed_layout(table):
    """Lock table to fixed column widths (prevents Word autofit from reflowing)."""
    # python-docx ordered setter: inserts before w:tblLook automatically
    layout = table._element.tblPr.get_or_add_tblLayout()
    layout.set(qn('w:type'), 'fixed')


def add_arrow_row(doc, left_text, right_text, total_cm=TABLE_WIDTH_CM, arrow_cm=2.0,
                  font_size=10.5, left_bold=True, right_bold=False,
                  left_shade=None, right_shade=None):
    """Horizontal "process → output" row with arrow centered in the table.

    Uses a 3-column table (left | arrow | right) with equal-width left/right
    columns and fixed layout, so the arrow is always geometrically centered
    regardless of text length. The middle vertical border is removed (fused),
    making left and right cells appear seamlessly connected.

    Shading logic:
      - left_shade / right_shade default None (white background)
      - Equal status (both None): both cells white
      - Different hierarchy: pass left_shade='D9D9D9' for left gray, right white

    Args:
        left_text: text for the left box (process step).
        right_text: text for the right box (output/result).
        total_cm: total table width in cm (default TABLE_WIDTH_CM).
        arrow_cm: width of the arrow column in cm (default 2.0).
        font_size: font size for left/right text (default 10.5).
        left_bold: whether left text is bold (default True).
        right_bold: whether right text is bold (default False).
        left_shade: hex color for left cell background, or None for white.
        right_shade: hex color for right cell background, or None for white.
    """
    side_cm = (total_cm - arrow_cm) / 2.0

    table = _new_table(doc, 1, [side_cm, arrow_cm, side_cm], margins=80,
                       borders=set_table_border_no_insideV)

    # Left box (process)
    c0 = table.cell(0, 0)
    if left_shade:
        set_cell_shading(c0, left_shade)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p0.add_run(left_text), '宋体', font_size, bold=left_bold)

    # Arrow (centered in its own column; equal side widths ⇒ table center)
    p1 = table.cell(0, 1).paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p1.add_run('\u2192'), '宋体', 14, bold=True)

    # Right box (output)
    c2 = table.cell(0, 2)
    if right_shade:
        set_cell_shading(c2, right_shade)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p2.add_run(right_text), '宋体', font_size, bold=right_bold)

    return table


def add_separator_note(doc, text):
    """Centered dashed separator line with note text (e.g. '----- AI 介入止于此处 -----').

    Used in flowcharts to mark a boundary between AI and existing systems.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    set_run_font(run, '宋体', 10.5, bold=False)


# ============================================================
# LAYERED ARCHITECTURE DIAGRAM (table-box based, no connectors)
# ============================================================

def _add_layer_spacer(doc, size_pt=6):
    """Small empty paragraph between diagram layers.

    Technically required: without a separating paragraph Word merges two
    adjacent tables into one.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    run = p.add_run('')
    run.font.size = Pt(size_pt)


def _add_arch_box(doc, lines, width_cm, font_size):
    """Single full-width layer box; lines[0] bold (layer name), rest plain."""
    table = _new_table(doc, 1, [width_cm], margins=80)
    cell = table.cell(0, 0)
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        set_run_font(p.add_run(line), '宋体', font_size, bold=(i == 0))
    return table


def _add_arch_parallel(doc, boxes, total_cm, col_widths, font_size):
    """Row of parallel layer boxes; each box's lines[0] bold (layer name).

    Column widths sum to `total_cm` either directly (col_widths given) or by
    even split — the caller passes the diagram's single total width, so this
    row can never end up wider or narrower than the full-width layers.
    """
    if col_widths is None:
        col_widths = [total_cm / len(boxes)] * len(boxes)
    table = _new_table(doc, 1, col_widths, margins=60)
    for i, lines in enumerate(boxes):
        cell = table.cell(0, i)
        for k, line in enumerate(lines):
            p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            set_run_font(p.add_run(line), '宋体', font_size, bold=(k == 0))
    return table


def add_layered_architecture(doc, layers, width_cm=TABLE_WIDTH_CM, col_widths=None,
                             font_size=10.5):
    """Layered architecture diagram built from table boxes (v1: no connectors).

    Each element of `layers` is one diagram layer, in top-to-bottom order.
    Two forms are accepted:

      - list[str]: full-width box. lines[0] is the bold layer-name line;
        remaining lines are plain centered content lines.
      - list[list[str]]: row of parallel boxes. Each inner list is one box
        (lines[0] bold layer name, rest plain lines).

    Total width has ONE source for the whole diagram, resolved once here and
    handed to every layer, so all boxes in the figure are exactly as wide as
    each other:

      - `col_widths` given → its SUM is the diagram width (the full-width
        layers take that same total, the parallel row uses the split as is);
      - only `width_cm` given → that is the width; the parallel row splits it
        evenly.

    A small spacer paragraph is inserted between layers (prevents Word from
    merging adjacent tables). Connector lines/arrows are intentionally
    omitted in v1 per design decision.

    Args:
        layers: list of layers, top to bottom. Example::

            add_layered_architecture(doc, [
                ["应用层（前端 · B/S）", "菜单栏 / 数据上传 / 参数配置"],
                [["预处理模块", "高清栅格化"],
                 ["配置模块", "参数模板库"]],
                ["数据接入层：上传 → 校验 → 归档"],
            ], col_widths=[4, 4, 6])

        (a 2-string list is a full-width box; a list of lists is a parallel row)

        width_cm: total width of the diagram in cm (default TABLE_WIDTH_CM),
            ignored when `col_widths` is given.
        col_widths: column widths of the parallel row in cm; their sum defines
            the diagram's total width.
        font_size: box font size in pt (default 10.5).
    """
    total_cm = round(sum(col_widths), 3) if col_widths else width_cm
    for idx, layer in enumerate(layers):
        if idx > 0:
            _add_layer_spacer(doc)
        if layer and isinstance(layer[0], (list, tuple)):
            _add_arch_parallel(doc, layer, total_cm, col_widths, font_size)
        else:
            _add_arch_box(doc, layer, total_cm, font_size)


# ============================================================
# FIGURE/TABLE AUTO-NUMBERING
# ============================================================

_fig_counter = 0
_tbl_counter = 0


def reset_counters():
    """Reset figure and table counters to zero.
    Called automatically by setup_document(); manual call only needed
    if generating multiple documents without re-calling setup_document().
    """
    global _fig_counter, _tbl_counter
    _fig_counter = 0
    _tbl_counter = 0


def add_fig_caption(doc, text):
    """Figure caption (centered, below figure). Auto-increments figure number.

    Format: "图N 描述文字" (N starts at 1, auto-increments).
    Use empty string "" for spacing-only caption (no number, no text).
    """
    global _fig_counter
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    if text:
        _fig_counter += 1
        # 对齐 NJUThesis njucap：标签加粗，标签与描述间全角空格（labelsep=quad）
        cap = _ACTIVE_PRESET['caption']
        label_run = p.add_run(f"图{_fig_counter}")
        set_run_font(label_run, '宋体', cap['size'], bold=cap.get('label_bold', False))
        sep_run = p.add_run("　")
        set_run_font(sep_run, '宋体', cap['size'])
        text_run = p.add_run(text)
        set_run_font(text_run, '宋体', cap['size'])


def add_table_caption(doc, text):
    """Table caption (centered, above table). Auto-increments table number.

    Format: "表N 描述文字" (N starts at 1, auto-increments).
    """
    global _tbl_counter
    _tbl_counter += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    # 对齐 NJUThesis njucap：标签加粗，标签与描述间全角空格（labelsep=quad）
    cap = _ACTIVE_PRESET['caption']
    label_run = p.add_run(f"表{_tbl_counter}")
    set_run_font(label_run, '宋体', cap['size'], bold=cap.get('label_bold', False))
    sep_run = p.add_run("　")
    set_run_font(sep_run, '宋体', cap['size'])
    text_run = p.add_run(text)
    set_run_font(text_run, '宋体', cap['size'])


def add_note(doc, text):
    """Figure note (centered, italic, smaller font)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, '宋体', _ACTIVE_PRESET['note']['size'], bold=False)
    run.italic = True


def add_bibliography(doc, entries, title='参考文献'):
    """Bibliography section, GB/T 7714-2015 numeric style (aligned with NJUThesis).

    Title: 黑体 三号(16pt) 加粗 居中, uses Heading 1 style so it enters TOC
    (like NJUThesis unnumbered 参考文献 chapter).
    Entries: 宋体 五号(10.5pt), justified, 1.5x line spacing, hanging indent
    (2 chars) so wrapped lines align after the [n] label.

    entries: list of GB/T 7714-2015 formatted strings WITHOUT the leading
    [n] label — labels [1], [2], ... auto-numbered in given order,
    matching in-text citation superscripts rendered by add_body.
    Entry examples:
      期刊: 作者. 题名[J]. 刊名, 年, 卷(期): 页码
      图书: 作者. 书名[M]. 出版地: 出版社, 年
      会议: 作者. 题名[C]//会议名. 出版地: 出版社, 年: 页码
      学位论文: 作者. 题名[D]. 城市: 学校, 年
      电子资源: 作者. 题名[EB/OL]. (更新日期)[引用日期]. 访问路径
    """
    s = _ACTIVE_PRESET['h1']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    run = p.add_run(title)
    set_run_font(run, s['font'], s['size'], bold=s['bold'],
                 color=_preset_color(s['color']), west_font=_west_font('head'))
    for i, entry in enumerate(entries, 1):
        ep = doc.add_paragraph()
        ep.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ep.paragraph_format.first_line_indent = Pt(-21)
        ep.paragraph_format.left_indent = Pt(21)
        ep.paragraph_format.line_spacing = 1.5
        ep.paragraph_format.space_after = Pt(3)
        erun = ep.add_run(f"[{i}] {entry}")
        set_run_font(erun, '宋体', 10.5, bold=False, west_font=_west_font('body'))


def _setup_toc_styles(doc):
    """Pre-define TOC entry styles so Word applies correct formatting when updating the TOC field.

    Values come from the active preset's 'toc' section. Default preset,
    aligned with NJUThesis LaTeX template:
    - TOC 1 (一级目录条目): 黑体 四号(14pt) 不加粗, 固定行距22磅
    - TOC 2 (二级目录条目): 宋体 小四(12pt) 不加粗, 固定行距22磅

    Key: must NOT carry customStyle="1", otherwise Word ignores these
    styles when updating the TOC field and falls back to the built-in
    template defaults (which bold TOC 2).
    """
    t = _ACTIVE_PRESET['toc']
    # (level, font_name, size_pt, bold)
    toc_specs = [
        (1, t['toc1_font'], t['toc1_size'], False),
        (2, t['toc2_font'], t['toc2_size'], False),
    ]
    for level, font_name, size_pt, bold in toc_specs:
        style_name = f'TOC {level}'
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

        # --- Remove customStyle="1" so Word treats this as a built-in TOC style ---
        style.element.attrib.pop(qn('w:customStyle'), None)

        # --- Ensure basedOn Normal (cut inheritance from template defaults) ---
        existing_basedon = style.element.find(qn('w:basedOn'))
        if existing_basedon is None:
            based_on = OxmlElement('w:basedOn')
            based_on.set(qn('w:val'), 'Normal')
            # Insert basedOn right after <w:name>, before <w:pPr>/<w:rPr>
            name_el = style.element.find(qn('w:name'))
            if name_el is not None:
                name_el.addnext(based_on)
            else:
                style.element.insert(0, based_on)

        # --- Font ---
        style.font.name = font_name
        style.font.size = Pt(size_pt)
        style.font.bold = bold

        # --- rPr: East Asian font + bCs + szCs ---
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn('w:eastAsia'), font_name)

        # bCs (complex script bold): must match bold setting
        for old_bcs in rpr.findall(qn('w:bCs')):
            rpr.remove(old_bcs)
        bcs = OxmlElement('w:bCs')
        bcs.set(qn('w:val'), '1' if bold else '0')
        rpr.append(bcs)

        # szCs (complex script size): size_pt * 2 half-points
        for old_szcs in rpr.findall(qn('w:szCs')):
            rpr.remove(old_szcs)
        szcs = OxmlElement('w:szCs')
        szcs.set(qn('w:val'), str(size_pt * 2))
        rpr.append(szcs)

        # --- Line spacing: fixed (from preset, in pt) ---
        pf = style.paragraph_format
        pf.line_spacing = Pt(t['line_spacing'])


# CT_Settings is an xsd:sequence — w:updateFields must sit between
# w:alwaysMergeEmptyNamespace and w:hdrShapeDefaults. python-docx emits
# w:compat right after w:savePreviewPicture, so inserting before the first
# successor element present keeps the element order schema-valid.
_UPDATE_FIELDS_SUCCESSORS = (
    'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars',
    'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping',
    'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade',
    'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary',
    'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator',
)


def _enable_update_fields_on_open(doc):
    """Write <w:updateFields w:val="true"/> into word/settings.xml.

    Applications that support field calculation (e.g. Word) recalculate all
    fields when the document opens, so TOC entries are generated without a
    manual right-click. Word asks for confirmation on open.

    Idempotent — inserting twice leaves a single element.
    """
    settings_el = doc.settings.element
    if settings_el.find(qn('w:updateFields')) is not None:
        return

    update_fields = OxmlElement('w:updateFields')
    update_fields.set(qn('w:val'), 'true')

    for child in settings_el:
        if child.tag.split('}')[-1] in _UPDATE_FIELDS_SUCCESSORS:
            child.addprevious(update_fields)
            return
    settings_el.append(update_fields)


def add_toc(doc, title='目  录', levels='1-2', auto_update=True):
    """Insert a Word Table of Contents field (TOC). On-demand only.

    Formatting:
      - Title: 黑体 三号(16pt) 居中 加粗
      - Entries (aligned with NJUThesis):
        - TOC 1 (一级): 黑体 四号(14pt) 不加粗
        - TOC 2 (二级): 宋体 小四(12pt) 不加粗
      - Line spacing: 固定22磅
      - Levels: default '1-2' (H1 + H2 only)

    Field update:
      - auto_update=True (default): writes <w:updateFields w:val="true"/> into
        word/settings.xml. Word prompts to update fields when the document
        opens; accepting generates the TOC entries.
      - auto_update=False: the TOC field updates only on demand — right-click
        the TOC area in Word and select "更新域".

    Args:
        title: TOC title text (default '目  录').
        levels: heading levels to include, e.g. '1-2' for H1+H2 only.
        auto_update: request a field update when the document is opened.
    """
    if auto_update:
        _enable_update_fields_on_open(doc)

    # Pre-configure TOC entry styles
    _setup_toc_styles(doc)

    # Page break before TOC
    doc.add_page_break()

    # TOC title (font/size from the active preset's 'toc' section)
    t = _ACTIVE_PRESET['toc']
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(title)
    set_run_font(run, t['title_font'], t['title_size'], bold=True, color=RGBColor(0, 0, 0))

    # TOC field: TOC \o "1-2" \h \z \u
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Pt(0)

    # begin field
    run_begin = paragraph.add_run()
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    run_begin._element.append(fldChar_begin)

    # field instruction
    run_instr = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = f' TOC \\o "{levels}" \\h \\z \\u '
    run_instr._element.append(instrText)

    # separate field
    run_sep = paragraph.add_run()
    fldChar_sep = OxmlElement('w:fldChar')
    fldChar_sep.set(qn('w:fldCharType'), 'separate')
    run_sep._element.append(fldChar_sep)

    # placeholder text
    run_placeholder = paragraph.add_run('（请在 Word 中右键此处选择"更新域"以生成目录）')
    set_run_font(run_placeholder, '宋体', 12, bold=False)

    # end field
    run_end = paragraph.add_run()
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run_end._element.append(fldChar_end)

    # page break after TOC
    doc.add_page_break()


# ============================================================
# DOCUMENT SETUP
# ============================================================

def setup_document():
    """Create a Document with page setup from the active preset.

    Reads the 'page' section (paper size, margins, header/footer distance)
    and the 'body' section (Normal default font) of the active preset. Call
    set_preset() before this function when switching presets.
    Also resets figure/table counters so each document starts from 图1/表1.
    """
    reset_counters()
    doc = Document()

    # The python-docx default template emits <w:zoom w:val="bestFit"/> without
    # w:percent, which its schema copy marks as required. Write a cached 100%
    # so the settings part validates; w:val still governs the zoom type, and
    # the spec treats percent as ignored whenever w:val is not "none".
    zoom = doc.settings.element.find(qn('w:zoom'))
    if zoom is not None and zoom.get(qn('w:percent')) is None:
        zoom.set(qn('w:percent'), '100')

    pg = _ACTIVE_PRESET['page']
    for section in doc.sections:
        section.page_width = Cm(pg['page_width'])
        section.page_height = Cm(pg['page_height'])
        section.top_margin = Cm(pg['margin_top'])
        section.bottom_margin = Cm(pg['margin_bottom'])
        section.left_margin = Cm(pg['margin_left'])
        section.right_margin = Cm(pg['margin_right'])
        section.header_distance = Cm(pg['header_distance'])
        section.footer_distance = Cm(pg['footer_distance'])

    # Set default font from the body preset (west font for ascii when separated)
    b = _ACTIVE_PRESET['body']
    style = doc.styles['Normal']
    style.font.name = _west_font('body') or b['font']
    style.font.size = Pt(b['size'])
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), b['font'])

    _add_page_number_footer(doc)

    return doc


def _add_page_number_footer(doc):
    """Footer page number: centered, 五号(10.5pt) Times New Roman (NJUThesis style).

    Inserts a PAGE field into the default footer paragraph of every section,
    so page numbers render in Word after field update.
    """
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        # clear any default empty run
        for existing_run in list(p.runs):
            existing_run._element.getparent().remove(existing_run._element)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = ' PAGE '
        fld_sep = OxmlElement('w:fldChar')
        fld_sep.set(qn('w:fldCharType'), 'separate')
        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')
        run._element.extend([fld_begin, instr, fld_sep, fld_end])
        set_run_font(run, 'Times New Roman', 10.5)


# ============================================================
# MAIN — Modify this function for your document content
# ============================================================

def main():
    OUTPUT_PATH = r'OUTPUT.docx'  # TODO: Set output path

    doc = setup_document()

    # --- Example content (replace with your own) ---

    add_title(doc, "文档标题（居中）")

    add_h1(doc, "一、概述")
    add_body(doc, "在此填写文档概述内容")

    add_h1(doc, "二、技术背景")
    add_body(doc, "在此填写技术背景描述")

    # Example: Table-box diagram (flowchart)
    add_h1(doc, "三、系统架构")
    add_box(doc, "第一层（描述内容）", width_cm=12, font_size=10)
    add_arrow_down(doc)
    add_box(doc, "第二层（描述内容）", width_cm=12, font_size=10)
    add_arrow_down(doc)
    add_multi_col_table(doc, [
        ("模块A", "描述A"),
        ("模块B", "描述B"),
        ("模块C", "描述C"),
    ], col_width_cm=[4, 4, 4], font_size=9)
    add_arrow_down(doc)
    add_box(doc, "第三层（描述内容）", width_cm=12, font_size=10)
    add_fig_caption(doc, "图1 系统架构图")

    # --- End of example content ---

    doc.save(OUTPUT_PATH)
    print(f'Word文档已生成: {OUTPUT_PATH}')


# ============================================================
# SAFE EXTRACT / REZIP — for editing existing .docx files
# Uses only Python standard library (zipfile, stat, pathlib)
# ============================================================

def safe_extract(zf, dest):
    """Safely extract a .docx ZIP archive, preventing path traversal and symlinks.

    Args:
        zf: An open zipfile.ZipFile object.
        dest: Destination directory (str or Path).

    Raises:
        ValueError: If a symlink entry or path-traversal entry is detected.
    """
    import stat as _stat
    dest = Path(dest).resolve()
    dest.mkdir(parents=True, exist_ok=True)

    for member in zf.infolist():
        # Reject symlink entries (untrusted external documents)
        if _stat.S_ISLNK(member.external_attr >> 16):
            raise ValueError(f"symlink archive entry not allowed: {member.filename!r}")

        # Resolve and verify path stays within dest
        target = (dest / member.filename).resolve()
        try:
            target.relative_to(dest)
        except ValueError:
            raise ValueError(f"unsafe archive entry escapes destination: {member.filename!r}")

        zf.extract(member, dest)


def rezip(src_dir, out_path):
    """Repackage a directory into a .docx ZIP file.

    Ensures [Content_Types].xml is stored first (uncompressed) as required
    by the OOXML specification for optimal reader compatibility.

    Args:
        src_dir: Source directory (str or Path) containing the unpacked .docx.
        out_path: Output .docx file path (str or Path).
    """
    import os as _os
    import tempfile as _tempfile

    src_dir = Path(src_dir)
    out_path = Path(out_path)
    files = sorted(p for p in src_dir.rglob("*") if p.is_file())
    ct = src_dir / "[Content_Types].xml"

    fd, tmp_name = _tempfile.mkstemp(
        prefix=out_path.name + ".", suffix=".tmp", dir=str(out_path.parent)
    )
    tmp_out = Path(tmp_name)
    try:
        import os as _os
        with _os.fdopen(fd, "wb") as fh:
            with zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zf:
                if ct.exists():
                    zf.write(ct, ct.relative_to(src_dir), compress_type=zipfile.ZIP_STORED)
                for f in files:
                    if f == ct:
                        continue
                    zf.write(f, f.relative_to(src_dir))

        # Preserve permissions
        if out_path.exists():
            mode = out_path.stat().st_mode & 0o777
        else:
            umask = _os.umask(0)
            _os.umask(umask)
            mode = 0o666 & ~umask
        _os.chmod(tmp_out, mode)
        _os.replace(tmp_out, out_path)
    finally:
        if tmp_out.exists():
            tmp_out.unlink()


if __name__ == '__main__':
    main()
