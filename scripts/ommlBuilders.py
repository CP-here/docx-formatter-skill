# -*- coding: utf-8 -*-
"""
ommlBuilders.py
OMML (Office Math Markup Language) math element builders — Python port of mathHelpers.js.

Generates raw OMML XML strings, insertable into python-docx paragraphs via
the add_eq_para() / add_body_with_math() helpers in docxBuilder.py.

Zero external dependencies (stdlib only). Function signatures are identical
to the JS version, so formula definitions port 1:1:

    from ommlBuilders import r, sub, sup, frac, sumOp, func, paren, bracket, math, inlineMath

    eq = math([
        sub("L", "LLM"), r(" = - "),
        sumOp([r("i")], [sub("y", "i"), r(" log("), sub("p", "i"), r(")")]),
    ])

CRITICAL NOTES (same rules as JS version):
1. sumOp() returns a LIST [sum_xml, *body] — all children-receiving functions
   call _flat() so nested lists are handled automatically.
2. Never emit an n-ary operator with empty superscript — it renders an
   invisible superscript box in Word and may corrupt the file. Instead we use
   <m:sSub> with Unicode ∑ (U+2211) as the base character (sumOp does this).
3. All text passed through r() is XML-escaped (&, <, >), so formulas may
   safely contain those characters.
"""

from xml.sax.saxutils import escape

# OOXML math namespace (injected by docxBuilder._insert_omml when parsing)
M_NS_DECL = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def _flat(items):
    """Flatten nested lists/tuples (equivalent of JS flat())."""
    out = []
    for it in items:
        if isinstance(it, (list, tuple)):
            out.extend(_flat(it))
        else:
            out.append(it)
    return out


def r(text):
    """Plain math text run:  r("L_total")"""
    return "<m:r><m:t>%s</m:t></m:r>" % escape(text)


def sub(base, sub_text):
    """Subscript:  sub("L", "LLM") → L_LLM"""
    return (
        "<m:sSub><m:sSubPr/><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>"
        % (r(base), r(sub_text))
    )


def sup(base, sup_text):
    """Superscript:  sup("x", "2") → x²"""
    return (
        "<m:sSup><m:sSupPr/><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>"
        % (r(base), r(sup_text))
    )


def frac(num_children, den_children):
    """Fraction:  frac([r("a")], [r("b")]) → a/b"""
    num = "".join(_flat(num_children))
    den = "".join(_flat(den_children))
    return "<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>" % (num, den)


def sumOp(sub_children, body_children):
    """Summation:  sumOp([r("i")], [sub("y","i")]) → ∑_i y_i

    IMPORTANT: returns a LIST [sum_symbol, *body]. Consumers flatten this
    automatically via _flat(). Do NOT replace with <m:nary> — empty
    superscript renders an invisible placeholder box in Word.
    """
    sum_symbol = (
        "<m:sSub><m:sSubPr/><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>"
        % (r("\u2211"), "".join(_flat(sub_children)))  # ∑ U+2211
    )
    return [sum_symbol] + _flat(body_children)


def func(name, arg_children):
    """Function with parentheses:  func("log", [sub("p","i")]) → log(p_i)"""
    args = "".join(_flat(arg_children))
    return (
        "<m:func><m:funcPr/><m:fName>%s</m:fName>"
        "<m:e><m:d><m:dPr/><m:e>%s</m:e></m:d></m:e></m:func>"
        % (r(name), args)
    )


def paren(children):
    """Round parentheses:  paren([r("a+b")]) → (a+b)"""
    inner = "".join(_flat(children))
    return "<m:d><m:dPr/><m:e>%s</m:e></m:d>" % inner


def bracket(children):
    """Square brackets:  bracket([r("a")]) → [a]"""
    inner = "".join(_flat(children))
    return (
        '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
        "<m:e>%s</m:e></m:d>" % inner
    )


def math(children):
    """Block math formula (centered standalone) — pass to add_eq_para()."""
    return "<m:oMath>%s</m:oMath>" % "".join(_flat(children))


def inlineMath(children):
    """Inline math formula — same XML as math(); insert into a text
    paragraph via add_body_with_math() parts list."""
    return "<m:oMath>%s</m:oMath>" % "".join(_flat(children))
