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
  - 黑体三号(16pt) Title centered, H1/H2/H3 left-aligned
  - 宋体小四(12pt) body, first-line indent 2 chars, 1.5 line spacing
  - Table-box diagrams: add_box(), add_multi_line_box(), add_multi_col_table(), add_arrow_down()
  - Cell margins (tcMar) for all table-box elements
  - Figure notes: add_note() for italic annotations below figures
  - Table Grid style with header shading (D9D9D9, gray)
  - **bold** text formatting support in body text
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path
import re
import zipfile


# ============================================================
# FONT & STYLE HELPERS
# ============================================================

def set_run_font(run, font_name, size, bold=False, color=None):
    """Set font name, size, bold for a run, including East Asian font."""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


def add_title(doc, text):
    """Document title: 黑体 三号(16pt) 居中 加粗 (for document main title only).
    Line spacing: 1.5x (aligned with NJUThesis linespread=1.625).
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, '黑体', 16, bold=True)


def add_h1(doc, text):
    """H1: 黑体 三号(16pt) 左对齐 加粗 黑色 (for section headings).
    Uses Word built-in Heading 1 style for TOC compatibility.
    Spacing: before=10pt, after=24pt, line=1.5x (NJUThesis chapter, after reduced for Word).
    """
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(24)
    run = p.add_run(text)
    set_run_font(run, '黑体', 16, bold=True, color=RGBColor(0, 0, 0))


def add_h2(doc, text):
    """H2: 黑体 小四(12pt) 左对齐 加粗 黑色.
    Uses Word built-in Heading 2 style for TOC compatibility.
    Spacing: before=18pt, after=12pt, line=1.5x (aligned with NJUThesis section).
    """
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 2']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, '黑体', 12, bold=True, color=RGBColor(0, 0, 0))


def add_h3(doc, text):
    """H3: 黑体 11pt 左对齐 加粗 黑色.
    Uses Word built-in Heading 3 style for TOC compatibility.
    Spacing: before=14pt, after=8pt, line=1.5x (aligned with NJUThesis subsection).
    """
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 3']
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    set_run_font(run, '黑体', 11, bold=True, color=RGBColor(0, 0, 0))


def add_body(doc, text):
    """Body: 宋体 小四(12pt) 两端对齐 首行缩进2字符 1.5倍行距.

    Supports:
      - **bold** segments
      - [n] citation markers → rendered as superscript (e.g. [1], [1,2], [1-3])
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(24)
    _add_runs_with_formatting(p, text)


def add_item_para(doc, label, text):
    """Item paragraph: bold label + normal body, first-line indent.

    Body text supports **bold** and [n] citation superscripts.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(24)
    pf.space_after = Pt(7)
    run = p.add_run(label)
    set_run_font(run, '宋体', 12, bold=True)
    _add_runs_with_formatting(p, text)


def _add_runs_with_formatting(p, text):
    """Add runs to paragraph p, parsing **bold** and [n] citation superscripts.

    Citation pattern matches: [1], [1,2], [1-3], [1, 2, 3], etc.
    """
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
            set_run_font(run, '宋体', 12, bold=True)
        elif re.match(r'^\[\d[\d,\-\s]*\]$', part):
            # Citation marker → superscript
            run = p.add_run(part)
            set_run_font(run, '宋体', 12, bold=False)
            run.font.superscript = True
        else:
            run = p.add_run(part)
            set_run_font(run, '宋体', 12, bold=False)


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
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(24)
    for kind, val in parts:
        if kind == "math":
            _insert_omml(p, val)
        else:
            _add_runs_with_formatting(p, val)
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
    tblPr.append(borders)


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
    tblPr.append(borders)


def set_cell_shading(cell, color_hex):
    """Set cell background color (e.g., 'D9D9D9' for gray)."""
    shading = cell._element.get_or_add_tcPr().makeelement(qn('w:shd'), {
        qn('w:val'): 'clear',
        qn('w:color'): 'auto',
        qn('w:fill'): color_hex,
    })
    cell._element.get_or_add_tcPr().append(shading)


# ============================================================
# TABLE-BOX DIAGRAM HELPERS (for flowcharts without images)
# ============================================================

def _set_cell_margins(cell, margin_dxa=80):
    """Set cell internal margins (top/bottom/left/right) in dxa units."""
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.makeelement(qn('w:tcMar'), {})
    for side in ['top', 'bottom', 'left', 'right']:
        mar = tcMar.makeelement(qn(f'w:{side}'), {qn('w:w'): str(margin_dxa), qn('w:type'): 'dxa'})
        tcMar.append(mar)
    tcPr.append(tcMar)


def add_box(doc, text, width_cm=14, font_size=10.5, bold=True):
    """Single centered box for flowchart steps."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    cell = table.cell(0, 0)
    cell.width = Cm(width_cm)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_margins(cell, 80)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(text)
    set_run_font(run, '宋体', font_size, bold=bold)
    return table


def add_multi_line_box(doc, lines, width_cm=14, font_size=10.5):
    """Multi-line centered box for flowchart steps.

    lines: list of strings, each rendered as a separate centered line.
    """
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    cell = table.cell(0, 0)
    cell.width = Cm(width_cm)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_margins(cell, 80)

    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(line)
        set_run_font(run, '宋体', font_size, bold=True)

    set_table_border(table)
    return table


def add_multi_col_table(doc, cells, col_width_cm=None, font_size=10.5):
    """
    Side-by-side boxes in a single row.
    cells: list of strings or list of (title, desc) tuples.
    col_width_cm: list of column widths in cm, or None for equal distribution.
    """
    n = len(cells)
    table = doc.add_table(rows=1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    if col_width_cm is None:
        total_width = 14  # default total ~14cm
        col_width_cm = [total_width / n] * n

    for i, cell_data in enumerate(cells):
        cell = table.cell(0, i)
        cell.width = Cm(col_width_cm[i])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_margins(cell, 60)

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

    set_table_border(table)
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
    tblPr = table._element.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')


def add_arrow_row(doc, left_text, right_text, total_cm=14.0, arrow_cm=2.0,
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
        total_cm: total table width in cm (default 14.0).
        arrow_cm: width of the arrow column in cm (default 2.0).
        font_size: font size for left/right text (default 10.5).
        left_bold: whether left text is bold (default True).
        right_bold: whether right text is bold (default False).
        left_shade: hex color for left cell background, or None for white.
        right_shade: hex color for right cell background, or None for white.
    """
    side_cm = (total_cm - arrow_cm) / 2.0

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    set_table_border_no_insideV(table)
    _set_table_fixed_layout(table)

    # Left box (process)
    c0 = table.cell(0, 0)
    c0.width = Cm(side_cm)
    c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_margins(c0, 80)
    if left_shade:
        set_cell_shading(c0, left_shade)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p0.add_run(left_text), '宋体', font_size, bold=left_bold)

    # Arrow (centered in its own column; equal side widths ⇒ table center)
    c1 = table.cell(0, 1)
    c1.width = Cm(arrow_cm)
    c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p1.add_run('\u2192'), '宋体', 14, bold=True)

    # Right box (output)
    c2 = table.cell(0, 2)
    c2.width = Cm(side_cm)
    c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_cell_margins(c2, 80)
    if right_shade:
        set_cell_shading(c2, right_shade)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.first_line_indent = Pt(0)
    set_run_font(p2.add_run(right_text), '宋体', font_size, bold=right_bold)

    # Enforce widths on all rows (insurance against Word reflow)
    for r in table.rows:
        r.cells[0].width = Cm(side_cm)
        r.cells[1].width = Cm(arrow_cm)
        r.cells[2].width = Cm(side_cm)

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
        caption = f"图{_fig_counter} {text}"
        run = p.add_run(caption)
        set_run_font(run, '宋体', 10.5, bold=False)


def add_table_caption(doc, text):
    """Table caption (centered, above table). Auto-increments table number.

    Format: "表N 描述文字" (N starts at 1, auto-increments).
    """
    global _tbl_counter
    _tbl_counter += 1
    caption = f"表{_tbl_counter} {text}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(caption)
    set_run_font(run, '宋体', 10.5, bold=False)


def add_note(doc, text):
    """Figure note (centered, italic, smaller font)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    set_run_font(run, '宋体', 9, bold=False)
    run.italic = True


def _setup_toc_styles(doc):
    """Pre-define TOC entry styles so Word applies correct formatting when updating the TOC field.

    Format aligned with NJUThesis LaTeX template:
    - TOC 1 (一级目录条目): 黑体 四号(14pt) 不加粗, 固定行距22磅
    - TOC 2 (二级目录条目): 宋体 小四(12pt) 不加粗, 固定行距22磅

    Key: must NOT carry customStyle="1", otherwise Word ignores these
    styles when updating the TOC field and falls back to the built-in
    template defaults (which bold TOC 2).
    """
    # (level, font_name, size_pt, bold)
    toc_specs = [
        (1, '黑体', 14, False),
        (2, '宋体', 12, False),
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

        # --- Line spacing: fixed 22pt ---
        pf = style.paragraph_format
        pf.line_spacing = Pt(22)


def add_toc(doc, title='目  录', levels='1-2'):
    """Insert a Word Table of Contents field (TOC). On-demand only.

    Formatting:
      - Title: 黑体 三号(16pt) 居中 加粗
      - Entries (aligned with NJUThesis):
        - TOC 1 (一级): 黑体 四号(14pt) 不加粗
        - TOC 2 (二级): 宋体 小四(12pt) 不加粗
      - Line spacing: 固定22磅
      - Levels: default '1-2' (H1 + H2 only)

    After opening in Word, right-click the TOC area and select "更新域" to generate.

    Args:
        title: TOC title text (default '目  录').
        levels: heading levels to include, e.g. '1-2' for H1+H2 only.
    """
    # Pre-configure TOC entry styles
    _setup_toc_styles(doc)

    # Page break before TOC
    doc.add_page_break()

    # TOC title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(title)
    set_run_font(run, '黑体', 16, bold=True, color=RGBColor(0, 0, 0))

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
    """Create a Document with standard page setup.
    Also resets figure/table counters so each document starts from 图1/表1.
    """
    reset_counters()
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # Set default font to 宋体 小四
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), '宋体')

    return doc


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
    ], col_width_cm=[5, 5, 5], font_size=9)
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
