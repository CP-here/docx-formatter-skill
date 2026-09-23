# -*- coding: utf-8 -*-
"""omml_math_kit.py 与 formula_templates.py 的单元测试。

将源码注释中的三条 CRITICAL NOTE 固化为可执行断言。这些约束
违反后生成的 XML 仍合法，docx_validator 无法检出，仅在 Microsoft
Word 中打开时才暴露渲染异常。

除 pytest 外无第三方依赖。
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
    """将 builder 输出（不含 xmlns:m 的片段）包入声明命名空间的容器后解析。

    直接解析片段会触发 unbound prefix；此处包装与
    docx_layout_kit._insert_omml 的处理一致。解析失败抛出
    xml.etree.ElementTree.ParseError，即本测试的失败信号。
    """
    return ET.fromstring("<wrap %s>%s</wrap>" % (M_NS_DECL, fragment))


def _as_fragment(out):
    """sumOp 契约性地返回 list，其余 builder 返回 str。"""
    return "".join(out) if isinstance(out, list) else out


def test_sum_op_returns_list():
    """CRITICAL NOTE 1：sumOp 返回 [求和符号, *body]，而非 str。

    消费方通过 _flat() 摊平嵌套；若改为 str，formula_templates 中所有
    嵌套 sumOp 的公式结构将随之改变。
    """
    out = sumOp([r("k")], [r("a")])
    assert isinstance(out, list), "sumOp 必须返回 list"
    assert out[1:] == [r("a")], "body 应作为 list 的后续元素紧跟求和符号"


def test_sum_op_no_empty_superscript():
    """CRITICAL NOTE 2a：求和符号不得携带上标。

    空上标在 Word 中渲染为不可见占位框；该缺陷 XML 仍合法，
    docx_validator 无法检出。
    """
    sum_symbol = sumOp([r("k")], [r("a")])[0]
    assert "<m:sup" not in sum_symbol, "求和符号出现上标，会在 Word 中产生空占位框"
    assert "<m:sub>" in sum_symbol, "求和符号应由 m:sSub 承载"


def test_sum_op_uses_ssub_not_nary():
    """CRITICAL NOTE 2b：求和符号使用 <m:sSub> + Unicode ∑，不得改写为 <m:nary>。

    <m:nary> 是 OOXML 对 n-ary 运算符的标准写法，重构时易被误用；
    其空上标即上一条所防缺陷。
    """
    sum_symbol = sumOp([r("k")], [r("a")])[0]
    assert "<m:nary" not in sum_symbol, "禁止使用 m:nary"
    assert "\u2211" in sum_symbol, "应以 U+2211 ∑ 作为基字符"


def test_r_escapes_xml_specials():
    """CRITICAL NOTE 3：r() 对 & < > 转义，公式可安全包含这些字符。

    通过解析回读验证文本节点，而非仅比对字面转义串。
    """
    raw = "a<b>&c and (x <= y)"
    # 结构为 wrap → m:r → m:t，需后代搜索而非直接子元素
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
    # eq3 嵌套最深（分式内含函数、再套两层 sumOp），单独覆盖便于
    # 定位深层嵌套相关的失败。
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
