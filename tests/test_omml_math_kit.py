# -*- coding: utf-8 -*-
"""omml_math_kit.py 与 formula_templates.py 的单元测试。

目标是把源码文档头里三条只写在注释中的 CRITICAL NOTES 固化成可执行断言：
这些规则任何重构都可能悄悄违反，而 docx_validator 的 5 项检查对此完全无感
—— 违反后生成的 XML 依然合法，只有打开 Microsoft Word 才看得出坏。

零第三方依赖（pytest 之外）。
"""
import xml.etree.ElementTree as ET

import pytest

from formula_templates import eq1, eq2, eq3, eq4
from omml_math_kit import (
    M_NS_DECL,
    bracket,
    frac,
    func,
    inlineMath,
    math,
    paren,
    r,
    sub,
    sup,
    sumOp,
)

# 与 M_NS_DECL 对应的命名空间 URI，供解析后按标签定位
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def parse(fragment):
    """把 builder 输出包进带 m 命名空间声明的容器后解析。

    builder 返回的是片段（自身不带 xmlns:m），直接 parse 会报
    unbound prefix；docx_layout_kit._insert_omml 做的也是同一件事。
    解析失败会抛 xml.etree.ElementTree.ParseError，即本测试的断言。
    """
    return ET.fromstring("<wrap %s>%s</wrap>" % (M_NS_DECL, fragment))


def _as_fragment(out):
    """sumOp 契约性地返回 list，其余 builder 返回 str。"""
    return "".join(out) if isinstance(out, list) else out


def test_sum_op_returns_list():
    """CRITICAL NOTE 1：sumOp 返回 [求和符号, *body]，不是 str。

    消费方（math/frac/paren 等）靠 _flat() 摊平嵌套；改成 str 会让
    formula_templates 里所有嵌套 sumOp 的公式结构变化。
    """
    out = sumOp([r("k")], [r("a")])
    assert isinstance(out, list), "sumOp 必须返回 list"
    assert out[1:] == [r("a")], "body 应作为 list 的后续元素紧跟求和符号"


def test_sum_op_no_empty_superscript():
    """CRITICAL NOTE 2a：求和符号不得携带上标。

    空上标会在 Word 里渲染成不可见占位框并可能损坏文件 —— 这类损坏
    docx_validator 查不出来（XML 完全合法），只有单测能挡住。
    """
    sum_symbol = sumOp([r("k")], [r("a")])[0]
    assert "<m:sup" not in sum_symbol, "求和符号出现上标 —— 会在 Word 中产生空占位框"
    assert "<m:sub>" in sum_symbol, "求和符号应由 m:sSub 承载"


def test_sum_op_uses_ssub_not_nary():
    """CRITICAL NOTE 2b：用 <m:sSub> + Unicode ∑，禁止改写为 <m:nary>。

    <m:nary> 才是 OOXML 规范里 n-ary 运算符的正统写法，重构时极易
    「顺手修正」成它，而其空上标正是上一条要防的 bug。
    """
    sum_symbol = sumOp([r("k")], [r("a")])[0]
    assert "<m:nary" not in sum_symbol, "禁止 m:nary（源码注释明确要求）"
    assert "\u2211" in sum_symbol, "应以 U+2211 ∑ 作为基字符"


def test_r_escapes_xml_specials():
    """CRITICAL NOTE 3：r() 对 & < > 转义，公式可安全包含这些字符。

    用解析回读的方式验证语义，而非仅比对字面转义串。
    """
    raw = "a<b>&c and (x <= y)"
    # 结构是 wrap → m:r → m:t，所以要后代搜索而非直接子元素
    node = parse(r(raw)).find(".//{%s}t" % M_NS)
    assert node is not None, "r() 应生成 m:t 文本节点"
    assert node.text == raw


_BUILDERS = [
    ("r", lambda: r("f(t) = x")),
    ("sub", lambda: sub("a", "k")),
    ("sup", lambda: sup("x", "2")),
    ("frac", lambda: frac([r("a")], [r("b")])),
    ("sumOp", lambda: sumOp([r("k")], [sub("a", "k"), r(" + ")])),
    ("func", lambda: func("cos", [r("k"), r("\u03c9t")])),
    ("paren", lambda: paren([r("a+b")])),
    ("bracket", lambda: bracket([r("1"), r("x")])),
    ("math", lambda: math([r("x = 1")])),
    ("inlineMath", lambda: inlineMath([r("x = 1")])),
    # 公式模板：eq3 是最深嵌套（分式内含函数内含分式再套两层 sumOp），
    # 单独覆盖它能在定位失败时立刻知道是深层嵌套出了问题。
    ("eq1", lambda: eq1),
    ("eq2", lambda: eq2),
    ("eq3", lambda: eq3),
    ("eq4", lambda: eq4),
]


@pytest.mark.parametrize("name, build", _BUILDERS, ids=[c[0] for c in _BUILDERS])
def test_all_builders_wellformed_xml(name, build):
    """每个 builder（含 4 个公式模板）的输出必须是良构 XML。"""
    fragment = _as_fragment(build())
    assert fragment, "%s 输出为空" % name
    parse(fragment)
