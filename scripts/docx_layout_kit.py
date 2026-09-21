# -*- coding: utf-8 -*-
"""
docx_layout_kit.py
Word (.docx) document builder — a library of typography and layout functions.

This module is a LIBRARY, not a runnable script. It has no main() entry point:
you import the helpers and assemble your own document in your own script.

PREREQUISITE: pip install python-docx

USAGE — write your own script that imports from this module:
  1. Copy this file (and omml_math_kit.py / formula_templates.py if formulas are
     needed) to your working directory
  2. Create a script that imports the helpers and calls them in order
  3. Run your script

Example:

    from docx_layout_kit import setup_document, add_title, add_h1, add_body

    doc = setup_document()
    add_title(doc, "文档标题")
    add_h1(doc, "一、概述")
    add_body(doc, "正文内容...")
    doc.save("output.docx")

For a complete worked example, read _build_readme_docx.py in this same directory —
it is the skeleton to copy: cover, TOC, headings, body, math, tables, diagrams
and bibliography in the correct call order.

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
  - Page numbering from the body: start_body() splits off the front matter
    (cover + TOC, unnumbered) and restarts the body section at page 1
  - Cover page: add_cover_page() writes a centered title block positioned by
    top padding. The padding is derived from the preset's text-area height, so
    switching preset/paper/margins keeps the cover position proportional.
  - Heading fonts on the styles, not the runs: _setup_heading_styles() puts
    the H1–H3 font/size/bold/colour on the Heading 1/2/3 styles so the heading
    runs stay free of character formatting. Word copies heading-run formatting
    onto TOC entry runs and it outranks the TOC styles, so leaving the runs
    clean is what keeps TOC 2/3 entries on the declared 宋体.
  - **bold** text formatting support in body text
"""

from docx import Document
from docx.shared import Pt, Cm, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
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
#   - page        纸张尺寸 + 页边距 + 页眉/页脚距离 + 正文页码起始值
#   - toc         目录标题与条目样式（字体/字号/行距/段前后/缩进）
#   - cover       封面顶部留白比例 + 留白空段单段高 + 标题块各元素字号与间距
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
            'body_page_start': 1,    # 正文节页码起始值（前置节不显示页码）
        },
        'toc': {
            'title_font': '黑体', 'title_size': 16,
            'line_spacing': 20,      # 目录条目固定行距（pt）
            # 条目规格按层级排列：章标题行(TOC 1) / 一级节标题行(TOC 2) / 二级节标题行(TOC 3)
            # indent_chars 以「汉字符」为单位，0 = 居左不缩进
            'levels': [
                {'font': '黑体', 'size': 12, 'indent_chars': 0,
                 'space_before': 6, 'space_after': 0},
                {'font': '宋体', 'size': 12, 'indent_chars': 1,
                 'space_before': 0, 'space_after': 0},
                {'font': '宋体', 'size': 12, 'indent_chars': 2,
                 'space_before': 0, 'space_after': 0},
            ],
        },
        'cover': {
            # 顶部留白 = top_padding_ratio × 版心高（版心高 = 纸张高 − 上下页边距）。
            # 用比例而非固定磅值，换纸张/页边距/预设时封面视觉位置自动等比适配。
            # 0.264 的取值：A4 下留白 6.50cm，大标题段起点距页顶 256.25pt
            # （页面高的 30.4%）。
            'top_padding_ratio': 0.264,
            # 留白空段单段高（pt，固定行距）。空段高度严格等于此值，留白因此可精确换算
            'spacer_height': 22,
            # 标题块（居中）。org 不设间距（为 0），date 与 org 之间隔 date_space_before
            'title_font': '黑体', 'title_size': 26, 'title_space_after': 8,
            'subtitle_font': '黑体', 'subtitle_size': 20, 'subtitle_space_after': 60,
            'org_font': '宋体', 'org_size': 15,
            'date_font': '宋体', 'date_size': 15, 'date_space_before': 8,
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
    Default preset: 黑体 16pt bold black, 1.5x, before=24pt, after=6pt
    (NJUThesis chapter, after reduced for Word).

    Font/size/bold/colour are declared on the Heading 1 style by
    _setup_heading_styles(), deliberately NOT on the run: a heading run
    carrying its own w:rFonts would be copied onto every TOC 1 entry by Word
    and override the TOC styles.
    """
    s = _ACTIVE_PRESET['h1']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    p.add_run(text)


def add_h2(doc, text):
    """H2, left-aligned, preset-driven.
    Uses Word built-in Heading 2 style for TOC compatibility.
    Default preset: 黑体 12pt bold black, 1.5x, before=12pt, after=6pt
    (aligned with NJUThesis section).

    Character formatting lives on the Heading 2 style, not the run — see
    _setup_heading_styles() for why the TOC entry fonts depend on this.
    """
    s = _ACTIVE_PRESET['h2']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 2']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    p.add_run(text)


def add_h3(doc, text):
    """H3, left-aligned, preset-driven.
    Uses Word built-in Heading 3 style for TOC compatibility.
    Default preset: 黑体 12pt bold black, 1.5x, before=12pt, after=6pt
    (aligned with NJUThesis subsection).

    Character formatting lives on the Heading 3 style, not the run — see
    _setup_heading_styles() for why the TOC entry fonts depend on this.
    """
    s = _ACTIVE_PRESET['h3']
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 3']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = s['line_spacing']
    p.paragraph_format.space_before = Pt(s['space_before'])
    p.paragraph_format.space_after = Pt(s['space_after'])
    p.add_run(text)


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
# OMML MATH HELPERS — formulas via omml_math_kit.py
# ============================================================

_M_NS_DECL = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def _insert_omml(p, omml_xml):
    """Parse an OMML XML string (from omml_math_kit) and append it to paragraph p.

    omml_math_kit.py is imported lazily so docx_layout_kit.py stays standalone
    (copyable without omml_math_kit.py) for math-free documents.
    """
    from docx.oxml import parse_xml  # noqa: local import, same package as OxmlElement
    if "xmlns:m=" not in omml_xml:
        omml_xml = omml_xml.replace("<m:oMath>", "<m:oMath %s>" % _M_NS_DECL, 1)
    p._p.append(parse_xml(omml_xml))


def add_eq_para(doc, math_xml):
    """Centered block math formula paragraph.

    Args:
        math_xml: OMML XML string from omml_math_kit.math(), e.g.
            from omml_math_kit import r, sub, sumOp, func, math
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
        from omml_math_kit import sub, inlineMath
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
    """Insert an inline OMML formula (omml_math_kit.inlineMath) into a table
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
    """Centered dashed separator line with note text (e.g. '----- 预处理止于此处 -----').

    Used in flowcharts to mark the boundary between two stages.
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
                ["采集层：传感器阵列 / 信号调理 / 抗混叠滤波"],
                ["预处理层：去直流 / 滤波 / 加窗"],
                [["变换模块", "快速傅里叶变换"],
                 ["分析模块", "频带能量 / 谱质心"]],
                ["应用层：频谱显示 / 报告输出"],
            ], col_widths=[4, 5, 5])

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
    (like NJUThesis unnumbered 参考文献 chapter). Font comes from the style,
    not the run, so Word does not copy it onto the TOC entries.
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
    p.add_run(title)
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

    Values come from the active preset's 'toc' section, one spec per level.
    Default preset:
    - TOC 1 章标题行:     黑体 小四(12pt), 固定行距20磅, 段前6磅/段后0磅, 居左
    - TOC 2 一级节标题行: 宋体 小四(12pt), 固定行距20磅, 段前0磅/段后0磅, 缩进1个汉字符
    - TOC 3 二级节标题行: 宋体 小四(12pt), 固定行距20磅, 段前0磅/段后0磅, 缩进2个汉字符

    Key: must NOT carry customStyle="1", otherwise Word ignores these
    styles when updating the TOC field and falls back to the built-in
    template defaults (which bold TOC 2).
    """
    t = _ACTIVE_PRESET['toc']
    line_spacing_pt = t['line_spacing']

    for level, spec in enumerate(t['levels'], start=1):
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

        font_name = spec['font']
        size_pt = spec['size']

        # --- Font (目录条目不使用加粗) ---
        style.font.name = font_name
        style.font.size = Pt(size_pt)
        style.font.bold = False

        # --- rPr: East Asian font + bCs + szCs ---
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn('w:eastAsia'), font_name)

        # bCs (complex script bold): must match the bold setting above.
        # get_or_add_bCs() keeps it in schema order (right after <w:b>).
        bcs = rpr.get_or_add_bCs()
        bcs.set(qn('w:val'), '0')

        # szCs (complex script size): size_pt * 2 half-points.
        # CT_RPr has no szCs descriptor, so insert manually right after
        # <w:sz> to stay schema-valid.
        for old_szcs in rpr.findall(qn('w:szCs')):
            rpr.remove(old_szcs)
        szcs = OxmlElement('w:szCs')
        szcs.set(qn('w:val'), str(size_pt * 2))
        sz_el = rpr.find(qn('w:sz'))
        if sz_el is not None:
            sz_el.addnext(szcs)
        else:
            rpr.append(szcs)

        # --- 段落格式：固定行距 + 段前/段后 ---
        pf = style.paragraph_format
        pf.line_spacing = Pt(line_spacing_pt)
        pf.space_before = Pt(spec['space_before'])
        pf.space_after = Pt(spec['space_after'])

        # --- 缩进：以「汉字符」为单位 ---
        # w:leftChars 用 Word 的字符单位（1 个汉字符 = 100），w:left 用 twips
        # 兜底（1 个汉字符宽 = 字号 × 20 twips，小四 12pt → 240 twips）。
        # 同时清掉可能继承来的悬挂/首行/右缩进，使条目左端只由本设置决定。
        ind = style.element.get_or_add_pPr().get_or_add_ind()
        for attr in ('w:hanging', 'w:hangingChars', 'w:firstLine',
                     'w:firstLineChars', 'w:right', 'w:rightChars'):
            ind.attrib.pop(qn(attr), None)
        indent_chars = spec['indent_chars']
        ind.set(qn('w:left'), str(round(size_pt * indent_chars * 20)))
        ind.set(qn('w:leftChars'), str(indent_chars * 100))


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


def add_toc(doc, title='目  录', levels='1-3', auto_update=True):
    """Insert a Word Table of Contents field (TOC). On-demand only.

    Formatting (from the active preset's 'toc' section):
      - Title: 黑体 三号(16pt) 居中 加粗
      - TOC 1 章标题行:     黑体 小四(12pt), 段前6磅/段后0磅, 居左
      - TOC 2 一级节标题行: 宋体 小四(12pt), 段前0磅/段后0磅, 缩进1个汉字符
      - TOC 3 二级节标题行: 宋体 小四(12pt), 段前0磅/段后0磅, 缩进2个汉字符
      - Line spacing: 固定20磅
      - Levels: default '1-3' (H1 + H2 + H3)

    Field update:
      - auto_update=True (default): writes <w:updateFields w:val="true"/> into
        word/settings.xml. Word prompts to update fields when the document
        opens; accepting generates the TOC entries.
      - auto_update=False: the TOC field updates only on demand — right-click
        the TOC area in Word and select "更新域".

    Page numbering: the TOC ends with a section break (start_body()), so the
    front matter (cover page, TOC) carries no page number and the body starts a
    new section numbered from 1. When the TOC follows a cover page the cover and
    the TOC are two separate sections, both unnumbered.

    Args:
        title: TOC title text (default '目  录').
        levels: heading levels to include, e.g. '1-3' for H1+H2+H3.
        auto_update: request a field update when the document is opened.
    """
    if auto_update:
        _enable_update_fields_on_open(doc)

    # Pre-configure TOC entry styles
    _setup_toc_styles(doc)

    # Page break before the TOC — skipped when the TOC is already the first
    # content of a fresh section (a cover page ends with a next-page section
    # break), where an extra page break would only produce a blank page.
    if not _last_section_is_empty(doc):
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

    # section break after TOC — starts the body section, whose page
    # numbering restarts at 1 (front matter stays unnumbered)
    start_body(doc)


# ============================================================
# DOCUMENT SETUP
# ============================================================

def _apply_page_setup(section):
    """Apply the preset's paper size, margins and header/footer distance."""
    pg = _ACTIVE_PRESET['page']
    section.page_width = Cm(pg['page_width'])
    section.page_height = Cm(pg['page_height'])
    section.top_margin = Cm(pg['margin_top'])
    section.bottom_margin = Cm(pg['margin_bottom'])
    section.left_margin = Cm(pg['margin_left'])
    section.right_margin = Cm(pg['margin_right'])
    section.header_distance = Cm(pg['header_distance'])
    section.footer_distance = Cm(pg['footer_distance'])


def _setup_heading_styles(doc):
    """Declare the H1–H3 character formatting on the styles, not on the runs.

    Why this matters for the TOC: Word rebuilds the TOC field by comparing the
    heading run's *own* character formatting against the TOC 1/2/3 style, and
    writes every difference as direct formatting on the entry run — direct
    formatting outranks the style. A heading run carrying ``w:rFonts`` forces
    that font onto every entry of the level (a 黑体 heading turns 宋体 TOC 2/3
    entries back into 黑体). Leaving the heading runs free of character
    formatting keeps the entries on whatever the TOC styles declare.

    The python-docx template ships Heading 1/2/3 (and their linked
    Heading 1/2/3 Char styles) with theme font references and a theme colour
    (blue 365F91 / 4F81BD), so both are rewritten here to plain font names and
    the preset colour.
    """
    for level in (1, 2, 3):
        s = _ACTIVE_PRESET[f'h{level}']
        west = _west_font('head') or s['font']
        for style_name in (f'Heading {level}', f'Heading {level} Char'):
            try:
                style = doc.styles[style_name]
            except KeyError:
                continue

            style.font.size = Pt(s['size'])
            color = _preset_color(s['color'])
            if color is not None:
                style.font.color.rgb = color

            rpr = style.element.get_or_add_rPr()

            # Font: drop the theme references, pin every script to a real font.
            rfonts = rpr.get_or_add_rFonts()
            for attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                rfonts.attrib.pop(qn(f'w:{attr}'), None)
            rfonts.set(qn('w:ascii'), west)
            rfonts.set(qn('w:hAnsi'), west)
            rfonts.set(qn('w:eastAsia'), s['font'])

            # Colour: the template's w:color keeps a themeColor reference that
            # would win over w:val, so strip it or the headings stay blue.
            color_el = rpr.find(qn('w:color'))
            if color_el is not None:
                for attr in ('themeColor', 'themeShade', 'themeTint'):
                    color_el.attrib.pop(qn(f'w:{attr}'), None)

            # Bold: rewrite <w:b>/<w:bCs> right after <w:rFonts> so the
            # CT_RPr element order stays schema-valid.
            for tag in ('b', 'bCs'):
                for old in rpr.findall(qn(f'w:{tag}')):
                    rpr.remove(old)
            anchor = rfonts
            for tag in ('b', 'bCs'):
                el = OxmlElement(f'w:{tag}')
                el.set(qn('w:val'), '1' if s['bold'] else '0')
                anchor.addnext(el)
                anchor = el

            # Size for complex scripts must mirror w:sz, placed right after it.
            for old in rpr.findall(qn('w:szCs')):
                rpr.remove(old)
            szcs = OxmlElement('w:szCs')
            szcs.set(qn('w:val'), str(int(round(s['size'] * 2))))
            sz_el = rpr.find(qn('w:sz'))
            if sz_el is not None:
                sz_el.addnext(szcs)
            else:
                rpr.append(szcs)


def setup_document():
    """Create a Document with page setup from the active preset.

    Reads the 'page' section (paper size, margins, header/footer distance),
    the 'body' section (Normal default font) and the 'h1'/'h2'/'h3' sections
    (Heading 1/2/3 styles) of the active preset. Call set_preset() before this
    function when switching presets.
    Also resets figure/table counters so each document starts from 图1/表1.

    Page numbers are placed on every section from the start; numbering must not
    run through the front matter when one exists — add_cover_page() ends the
    cover as its own unnumbered section, and start_body() (called at the end of
    add_toc()) starts the numbered body section.
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

    for section in doc.sections:
        _apply_page_setup(section)
        _add_page_number_footer(section)

    # Set default font from the body preset (west font for ascii when separated)
    b = _ACTIVE_PRESET['body']
    style = doc.styles['Normal']
    style.font.name = _west_font('body') or b['font']
    style.font.size = Pt(b['size'])
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), b['font'])

    _setup_heading_styles(doc)

    return doc


# w:pgNumType belongs between w:lnNumType and w:cols in the CT_SectPr
# xsd:sequence; inserting before the first successor present keeps the
# element order schema-valid.
_PG_NUM_TYPE_SUCCESSORS = (
    'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote', 'w:titlePg',
    'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid',
    'w:printerSettings', 'w:sectPrChange',
)


def _set_page_number_start(section, start):
    """Restart the page counter of `section` at `start` (<w:pgNumType w:start>)."""
    sectPr = section._sectPr
    for old in sectPr.findall(qn('w:pgNumType')):
        sectPr.remove(old)
    pg_num_type = OxmlElement('w:pgNumType')
    pg_num_type.set(qn('w:start'), str(start))
    sectPr.insert_element_before(pg_num_type, *_PG_NUM_TYPE_SUCCESSORS)


def _last_section_is_empty(doc):
    """True when the last section of `doc` holds no content yet.

    add_section() ends the previous section by cloning its w:sectPr into a new
    (empty) w:p and hands the body-level sentinel w:sectPr to the new section,
    so right after a section break the last section has no paragraphs at all.
    Both add_toc() and start_body() consult this to reuse that empty section
    instead of splitting again — a second split would print a blank page.
    """
    children = list(doc.element.body.iterchildren())

    # The last section begins after the paragraph carrying the previous
    # section's w:sectPr (absent when the document never had a section break).
    start = 0
    for i, child in enumerate(children):
        if child.tag != qn('w:p'):
            continue
        pPr = child.find(qn('w:pPr'))
        if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
            start = i + 1

    return all(child.tag == qn('w:sectPr') for child in children[start:])


def _is_body_section(section):
    """True when `section` is a body section already created by start_body()."""
    pg_num_type = section._sectPr.find(qn('w:pgNumType'))
    return pg_num_type is not None and pg_num_type.get(qn('w:start')) is not None


def _clear_footer(footer):
    """Drop every run of `footer` so the section prints no page number."""
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        for run in list(p.runs):
            run._element.getparent().remove(run._element)


def start_body(doc):
    """Start the body section: no page number before it, numbering from 1 after.

    Every section written before this call (cover page, TOC) becomes front
    matter and carries no page number; the body becomes its own section whose
    footer holds a PAGE field and whose page counter restarts at the preset's
    ``page.body_page_start``.

    The body section starts on a new page, so it doubles as the page break
    between the front matter and the body — do not add a page break of your own.

    If the last section is still empty (a cover page ended with a section break
    just before), it is reused as the body section instead of adding another
    one, which would otherwise leave a blank page behind.

    Call it once, right before the first body heading. add_toc() already calls
    it at its end; a second call is a no-op (idempotent).

    Documents without a cover page or TOC need no call: their single section is
    numbered from 1 by setup_document().
    """
    if _is_body_section(doc.sections[-1]):
        return

    # Front matter never prints a page number — clear every section so far,
    # not just the last one (a cover page is its own section).
    for section in doc.sections:
        _clear_footer(section.footer)

    if _last_section_is_empty(doc):
        section = doc.sections[-1]
    else:
        # add_section() clones the previous sectPr (paper size, margins,
        # distances carry over) and drops its header/footer references, so the
        # new section starts out inheriting the (now empty) front-matter footer.
        section = doc.add_section(WD_SECTION.NEW_PAGE)

    _apply_page_setup(section)
    _set_page_number_start(section, _ACTIVE_PRESET['page']['body_page_start'])
    _add_page_number_footer(section)


def _add_page_number_footer(section):
    """Footer page number: centered, 五号(10.5pt) Times New Roman (NJUThesis style).

    Inserts a PAGE field into the default footer paragraph of `section`, so the
    page number renders in Word after field update.
    """
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
# COVER PAGE
# ============================================================

_CM_TO_PT = 72.0 / 2.54          # 1 cm = 28.3465 pt


def _cover_line(doc, text, font, size, bold=False,
                space_before=0, space_after=0, west_font=None):
    """One centered line of the cover title block (first-line indent cleared)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.line_spacing = 1.5
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    set_run_font(p.add_run(text), font, size, bold=bold, west_font=west_font)
    return p


def add_cover_page(doc, title, subtitle=None, org=None, date=None):
    """Write a cover page, then close it as its own (unnumbered) section.

    The title block is positioned by top padding (not by vertical centering),
    and the padding is computed from the active preset, so the cover keeps its
    visual position when the paper size, margins or preset change:

        版心高   = 纸张高 − 上边距 − 下边距        (preset 'page' 组)
        目标留白 = cover.top_padding_ratio × 版心高
        空段数 N = 取整(目标留白 ÷ cover.spacer_height)
        余数     = 目标留白 − N × cover.spacer_height   (挂到主标题段段前)

    The spacer paragraphs use fixed (exact) line spacing, so each one is
    exactly ``spacer_height`` tall and the total padding equals the target to
    well under a point. The remainder rides on the main title's space-before;
    the title is never the first paragraph of the page, which keeps Word from
    suppressing that space (Word drops space-before on a page-top paragraph).

    Call this first, before any other content: it clears the footer of the
    section it writes into (no page number) and ends that section with a
    next-page section break, so whatever follows starts on a new page. Follow
    it with add_toc() — which starts the numbered body section at its end — or
    with a direct start_body() call when the document has no TOC.

    Args:
        title: main title line (required).
        subtitle: optional second line, e.g. 技术方案.
        org: optional organization line, e.g. 某某科技有限公司.
        date: optional date line, e.g. 2026年9月.
    """
    cv = _ACTIVE_PRESET['cover']
    pg = _ACTIVE_PRESET['page']

    text_height = ((pg['page_height'] - pg['margin_top'] - pg['margin_bottom'])
                   * _CM_TO_PT)
    # 比例上限 0.5：超过半个版心的留白会把标题块挤到第二页
    padding = min(cv['top_padding_ratio'], 0.5) * text_height

    h = cv['spacer_height']
    n = int(padding // h)
    if n < 1:
        # 至少一个空段，保证标题段不落在页首（否则段前间距会被 Word 抑制）
        n, remainder = 1, 0.0
    else:
        remainder = padding - n * h

    for _ in range(n):
        p = doc.add_paragraph()
        # 空段居中：段落标记显示在页面水平中间，用户增删空段调节留白时便于定位
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.first_line_indent = Pt(0)
        pf.line_spacing = Pt(h)      # 固定行距 → 空段高度精确等于 h
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

    head_west = _west_font('head')
    body_west = _west_font('body')

    _cover_line(doc, title, cv['title_font'], cv['title_size'], bold=True,
                space_before=remainder, space_after=cv['title_space_after'],
                west_font=head_west)
    if subtitle is not None:
        _cover_line(doc, subtitle, cv['subtitle_font'], cv['subtitle_size'],
                    bold=True, space_after=cv['subtitle_space_after'],
                    west_font=head_west)
    if org is not None:
        _cover_line(doc, org, cv['org_font'], cv['org_size'],
                    west_font=body_west)
    if date is not None:
        _cover_line(doc, date, cv['date_font'], cv['date_size'],
                    space_before=cv['date_space_before'], west_font=body_west)

    # 封面不显示页码，并以「下一页」分节符结束封面节
    _clear_footer(doc.sections[-1].footer)
    doc.add_section(WD_SECTION.NEW_PAGE)


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
