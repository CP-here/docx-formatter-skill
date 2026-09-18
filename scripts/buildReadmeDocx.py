# -*- coding: utf-8 -*-
"""范例：本技能生成文档的完整标准骨架（同时用于生成项目根的 README.docx）。

【怎么用这个文件】
本技能是函数库，docxBuilder.py 没有 main() 入口，直接运行没有任何输出。
你要做的是**新建一个自己的脚本**，把内容写进去。本文件就是那份脚本的范式：

  1. 读第 17 行起的 import 段 —— 需要哪些函数就从 docxBuilder 导入哪些
  2. 读第 52 行起的正文 —— add_cover_page → add_toc → add_h1/add_h2/add_h3
     → add_body → 图表 → add_bibliography 的**调用顺序**就是标准顺序
  3. 在你的脚本里照这个顺序写自己的内容

【本文件里哪些部分不必照抄】
下面三行是本文件作为「项目内构建脚本」特有的，你在临时工作目录里写自己的脚本
时**不需要**：把你的脚本与 docxBuilder.py 放在同一目录，直接 import 即可。

    SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
    PROJECT_ROOT = os.path.dirname(SCRIPTS_DIR)
    sys.path.insert(0, SCRIPTS_DIR)

同理，末尾的 doc.save(PROJECT_ROOT + "/README.docx") 也只是本文件的输出目标，
换成你自己的输出路径。

【三条最易踩的顺序规则】
  1. set_preset() 必须在 setup_document() 之前调用（页边距与默认字体在此时读取）
  2. add_cover_page() 必须写在所有其它内容之前（它以分节符结束封面节）
  3. 页码分节 —— 有目录时 add_toc() 已自动处理；只有封面无目录时需手动
     start_body(doc)；两者都无则不用管

【用法】在项目根目录执行：python scripts/buildReadmeDocx.py
"""
import sys, os

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPTS_DIR)
sys.path.insert(0, SCRIPTS_DIR)

from docxBuilder import (setup_document, add_cover_page, add_toc, add_h1, add_h2,
    add_h3, add_body, add_item_para, add_eq_para, add_body_with_math,
    add_code_block, add_data_table, add_math_to_cell, add_box, add_multi_line_box,
    add_multi_col_table, add_arrow_down, add_arrow_horizontal, add_arrow_row,
    add_separator_note, add_fig_caption, add_table_caption,
    add_layered_architecture, add_note, add_bibliography, set_run_font)
from ommlBuilders import r, sub, sup, frac, sumOp, func, math, inlineMath
from formulaTemplates import eq1, eq2

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT


def add_link_para(doc, parts):
    """正文格式段落，支持内联外部超链接。

    parts: ("text", str) 纯文本 | ("link", url, text) 超链接。
    格式对齐默认正文预设（宋体 12pt / 1.5 倍行距 / 首行缩进 2 字符 / 中西文分离）。
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(24)
    for part in parts:
        if part[0] == "link":
            _, url, text = part
            run = p.add_run(text)
            set_run_font(run, '宋体', 12, west_font='Times New Roman')
            run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            run.font.underline = True
            r_id = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
            hyperlink = OxmlElement('w:hyperlink')
            hyperlink.set(qn('r:id'), r_id)
            p._p.append(hyperlink)
            hyperlink.append(run._element)
        else:
            run = p.add_run(part[1])
            set_run_font(run, '宋体', 12, west_font='Times New Roman')
    return p


doc = setup_document()

# 封面：独立成节、不显示页码；其后的目录节同样不显示页码，正文从第 1 页起算
add_cover_page(doc, "docx-formatter 中文 Word 专业排版技能",
               subtitle="技能说明与排版效果样例")
add_toc(doc)

# ============ 一、项目简介 ============
add_h1(doc, "一、项目简介")
add_body(doc, "docx-formatter 生成专业排版的中文 Word 文档，覆盖封面、目录、页码分节、章节结构、正文、数学公式、表格、流程图、引用上标与文档验证的完整排版链路，输出符合学术与工程规范的 .docx 文件。排版效果可参考仓库中的 README.docx。")
add_body(doc, "技能采用单引擎架构，仅依赖 Python 3.8 以上版本与 python-docx。数学公式由 ommlBuilders.py 直接生成 OMML XML，即 Word 原生公式格式，公式在 Word 中可二次编辑，无需 Node.js 或任何外部转换工具。")
add_body(doc, "本文档即由该技能自身排版生成：封面由 add_cover_page 生成，目录为 Word 原生目录域，页脚页码为 PAGE 域且只在正文显示，文中所有公式为 Word 原生可编辑公式，流程图与架构图为表格框图，引用标记为真实上标，可作为排版效果的直接样例。")

add_h2(doc, "1.1 如何开始")
add_body(doc, "本技能为函数库，不含可执行入口：docxBuilder.py 无 main() 函数，直接运行不产生任何输出。Agent 生成文档的方式是编写调用脚本，导入库函数后按文档结构依次调用。")
add_body(doc, "Agent 生成文档前应先阅读 scripts/buildReadmeDocx.py。该文件是本技能的标准骨架，从 setup_document() 到 doc.save() 的完整调用顺序均在其正文示范；它同时用于生成项目根目录的 README.docx，因而始终与库保持同步。")
add_table_caption(doc, "参考文件")
add_data_table(doc,
    ["参考文件", "用途"],
    [
        ("scripts/buildReadmeDocx.py", "标准骨架范例，覆盖封面、目录、标题、正文、分点、公式、数据表、流程图、图表标题与参考文献的完整调用顺序"),
        ("SKILL.md", "各函数的完整参数说明、排版数值、规则与陷阱"),
        ("scripts/docxBuilder.py", "函数库本体与预设值，用于确认实现细节"),
    ],
    col_widths=[4.4, 10.0], font_size=9.5)
add_body(doc, "标准流程为：Agent 阅读 buildReadmeDocx.py 的 import 段与调用顺序，复制 docxBuilder.py 到工作目录（文档含公式时再加 ommlBuilders.py 与 formulaTemplates.py），新建调用脚本并按范例顺序写入内容，运行后用 docxValidator.py 验证。")

# ============ 二、功能矩阵 ============
add_h1(doc, "二、功能矩阵")
add_body(doc, "技能共包含 20 个模块，其中 13 个默认启用，2 个按需生成，5 个为可选模块，仅在明确要求时启用。")
add_table_caption(doc, "功能矩阵总览")
add_data_table(doc,
    ["模块", "默认状态", "说明"],
    [
        ("标准排版", "启用", "标题、正文、缩进与行距体系"),
        ("排版预设", "启用", "整套切换排版参数，支持中西文字体分离"),
        ("OMML 数学公式", "启用", "行内与块级公式，Word 原生可编辑"),
        ("引用上标", "启用", "参考文献标记自动渲染为上标"),
        ("参考文献列表", "启用", "GB/T 7714-2015 顺序编码制文献表"),
        ("表格框图流程图", "启用", "表格与 Unicode 箭头构建流程图"),
        ("分层架构图", "启用", "多层架构的表格框图堆叠"),
        ("图表自动编号", "启用", "图表标题自动递增编号"),
        ("代码块", "启用", "等宽字体与浅灰底纹"),
        ("数据表", "启用", "灰色表头与固定列宽"),
        ("单元格公式", "启用", "表格单元格内插入行内公式"),
        ("轻量文档验证", "启用", "纯标准库 5 项检查，生成后默认执行"),
        ("安全解压与重打包", "启用", "编辑现有文档的安全基础"),
        ("目录生成", "按需", "Heading 样式加目录域，末尾分节触发正文页码从 1 起算"),
        ("封面页", "按需", "居中标题块，顶部留白按版心比例自动适配"),
        ("PDF 渲染验证", "可选", "LibreOffice 转换后视觉检查"),
        ("XSD 模式验证", "可选", "OOXML 标准深度验证"),
        ("编辑现有文档", "可选", "解压、编辑与重打包工作流"),
        ("追踪修订", "可选", "接受全部追踪修订"),
        ("批注管理", "可选", "批注插入与交叉链接维护"),
    ],
    col_widths=[3.4, 1.6, 9.4], font_size=9)

# ============ 三、标准排版 ============
add_h1(doc, "三、标准排版")
add_body(doc, "标题使用黑体，正文使用宋体。文档标题一号 26pt 居中加粗，一级标题三号 16pt 左对齐，二级与三级标题小四 12pt。正文小四 12pt 两端对齐，首行缩进 2 字符，全文 1.5 倍行距。纸张为 A4，页边距上下 2.54cm、左右 3.18cm。分级页码：封面与目录不显示页码，正文从第 1 页起算，页脚居中放置五号 Times New Roman 的 PAGE 域。封面标题块由顶部留白定位，留白高度按版心高比例换算，换纸张或页边距时自动等比适配。各级标题的字体、字号、加粗与黑色统一写在 Word 内置 Heading 样式上，标题文字本身不带直接字符格式；这样既覆盖了样式默认的蓝色与模板主题字体引用，也让 Word 更新目录域时无法把标题字体搬进条目，目录条目因而严格按样式呈现。中西文字体分离：正文西文用 Times New Roman 衬线体，标题西文用 Arial 无衬线体，对齐 NJUThesis 的字体族设计。本文档各级标题与正文即为标准排版的实际效果。")
add_table_caption(doc, "标准排版构件函数")
add_data_table(doc,
    ["函数", "用途"],
    [
        ("setup_document", "创建文档并应用页面设置，自动重置图表计数器"),
        ("add_cover_page", "封面页，居中标题块并用顶部留白定位，独立成节且不显示页码"),
        ("add_title", "不放封面时的文档主标题，黑体三号居中加粗"),
        ("add_h1、add_h2、add_h3", "一级、二级、三级标题，使用 Word 内置 Heading 样式，可进入目录"),
        ("add_body", "正文段落，支持加粗标记与引用上标"),
        ("add_item_para", "加粗标签加分点内容的段落，标签整体加粗"),
    ],
    col_widths=[4.2, 10.2], font_size=9.5)

# ============ 四、排版预设 ============
add_h1(doc, "四、排版预设")
add_body(doc, "排版参数集中在 docxBuilder.py 顶部的 PRESETS 字典中，按标题、正文、分点、图注、页面、目录、封面等分组管理，各排版函数统一读取当前激活的预设。")
add_item_para(doc, "set_preset 整套切换：", "须在 setup_document 之前调用，默认预设与标准排版一致，新增预设只需向 PRESETS 追加一份同结构的嵌套字典。")
add_item_para(doc, "中西文字体分离：", "开启后西文使用 Times New Roman，中文保持原有字体不变。")
add_item_para(doc, "封面留白比例：", "位于预设的封面分组内，留白高度由它乘以版心高得出，换纸张与页边距时封面位置随之等比适配。")

# ============ 五、OMML 数学公式 ============
add_h1(doc, "五、OMML 数学公式")
add_body(doc, "公式构建器提供一组组合函数，直接生成 OMML XML 并插入文档，产物为 Word 原生可编辑公式，与手动插入的公式完全一致。下表为全部构建函数：")
add_table_caption(doc, "OMML 公式构建函数")
add_data_table(doc,
    ["构建函数", "用途"],
    [
        ("r", "纯数学文本，自动转义 XML 特殊字符"),
        ("sub", "下标，如 a 的下标 k"),
        ("sup", "上标，如 x 的平方"),
        ("frac", "分式"),
        ("sumOp", "求和符号，下标结构配合 Unicode ∑，规避空上标渲染异常"),
        ("func", "函数名加参数，如 cos 与 exp"),
        ("paren", "圆括号结构"),
        ("bracket", "方括号结构"),
        ("math", "块级公式，居中独立成行"),
        ("inlineMath", "行内公式，与正文混排"),
    ],
    col_widths=[3.2, 11.2], font_size=9.5)

add_h2(doc, "5.1 块级公式")
add_body(doc, "块级公式居中独立成行。以傅里叶级数为例，下方公式即为 Word 原生可编辑公式：")
add_eq_para(doc, math([
    r("f(t) = "),
    frac([sub("a", "0")], [r("2")]),
    r(" + "),
    sumOp([r("k")], [sub("a", "k"), func("cos", [r("k"), r("ωt")])]),
]))
add_fig_caption(doc, "傅里叶级数展开的块级公式")

add_h2(doc, "5.2 行内公式")
add_body(doc, "行内公式与正文混排，如下段中的傅里叶系数：")
add_body_with_math(doc, [
    ("text", "其中，"),
    ("math", inlineMath([sub("a", "k")])),
    ("text", "为傅里叶系数，其数值由欧拉公式确定，收敛速度与信号的光滑性相关。"),
])

add_h2(doc, "5.3 复杂嵌套结构")
add_body(doc, "对于分式套分式、双重求和、函数嵌套的复杂结构，构建函数可任意组合。下方为似然比检测损失的完整形态，包含分式嵌套、exp 与 log 嵌套、集合记号与权重系数：")
add_eq_para(doc, math([
    sub("E", "det"),
    r(" = −"),
    sumOp([r("n∈N")], [
        r("["),
        frac([r("1")], [r("|K(n)|")]),
        r(" "),
        sumOp([r("k∈K(n)")], [
            func("log", [
                frac(
                    [func("exp", [frac([r("s("), sub("x", "n"), r(","), sub("h", "k"), r(")")], [r("τ")])])],
                    [func("exp", [frac([r("s("), sub("x", "n"), r(","), sub("h", "k"), r(")")], [r("τ")])]),
                     r(" + "),
                     sumOp([r("m∉K(n)")], [sub("w", "m"), r(" · "),
                        func("exp", [frac([r("s("), sub("x", "n"), r(","), sub("h", "m"), r(")")], [r("τ")])])])],
                ),
            ]),
        ]),
        r("]"),
    ]),
]))
add_fig_caption(doc, "似然比检测损失的嵌套公式")

add_h2(doc, "5.4 希腊字母与数学符号")
add_body(doc, "formulaTemplates.py 内置若干常用公式模板，可直接插入文档，也可作为自定义公式的参考模式。希腊字母与数学符号以 Unicode 直接书写，如 λ、τ、α、β、∑、×、·、∈。")
add_body(doc, "求和符号必须使用 sumOp 构建。求和若采用 n 元运算符加空上标的写法，Word 会渲染一个不可见的上标占位框，导致文件看起来损坏；sumOp 改用下标结构配合 Unicode ∑ 实现，彻底规避该问题。所有接收子元素的函数内部自动展平嵌套列表，防止 XML 拼接错误。")

add_h2(doc, "5.5 预定义公式模板")
add_body(doc, "公式模板是现成的 OMML 公式对象，导入即可插入文档，无需重新拼装构建函数。下方两个模板分别覆盖加权组合与交叉熵两类常见形态，均为 Word 原生可编辑公式：")
add_eq_para(doc, eq1)
add_fig_caption(doc, "时域与频域加权组合公式模板")
add_eq_para(doc, eq2)
add_fig_caption(doc, "检测误差的交叉熵公式模板")

# ============ 六、引用上标 ============
add_h1(doc, "六、引用上标")
add_body(doc, "正文中的参考文献标记自动渲染为上标，适用于正文段落与分点段落，支持三种格式：")
add_table_caption(doc, "引用上标支持的格式")
add_data_table(doc,
    ["格式", "示例"],
    [
        ("单个引用", "[1]"),
        ("多个引用", "[1,2] 或 [1, 2]"),
        ("范围引用", "[1-3]"),
    ],
    col_widths=[3.5, 6.5], font_size=9.5)
add_body(doc, "本段即为实际渲染效果：傅里叶分析方法在信号处理任务中表现出色[1]，后续研究[2,3]进一步验证了这一结论，系统综述见[1-3]。")

# ============ 七、参考文献列表 ============
add_h1(doc, "七、参考文献列表")
add_body(doc, "参考文献表生成函数产出 GB/T 7714-2015 顺序编码制文献表，格式对齐 NJUThesis LaTeX 模板，与正文引用上标配套使用：正文中的 [1] 指向条目列表第一项，以此类推。")
add_body(doc, "章节标题为黑体三号加粗居中，使用 Heading 1 样式，进入目录；条目为宋体五号，两端对齐，1.5 倍行距，悬挂缩进 2 字符，折行后与序号后文字对齐；序号 [1]、[2] 按传入顺序自动生成，无需手写。下方即本文档的参考文献表实际渲染效果：")
add_bibliography(doc, [
    "张三, 李四. 示例期刊论文题名[J]. 示例期刊, 2020, 12(3): 45-56",
    "王五. 示例图书题名[M]. 示例市: 示例出版社, 2019",
    "赵六. 示例学位论文题名[D]. 示例市: 示例大学, 2021",
    "钱七. 示例会议论文题名[C]//示例会议论文集. 示例市: 示例出版社, 2022: 78-85",
    "示例标准化技术委员会. 示例标准题名: 示例标准编号[S]. 示例市: 示例出版社, 2018",
    "孙八. 示例网页题名[EB/OL]. (2023-01-01)[2023-06-05]. https://example.com",
])
add_body(doc, '条目按文献类型组织，传入时不带序号：期刊以 [J] 标注，图书 [M]，会议论文 [C]，学位论文 [D]，标准 [S]，电子资源 [EB/OL] 并附更新日期与访问路径。多作者用逗号分隔，三位以上作者取前三名加“等”或"et al"。章节标题默认为“参考文献”，可通过 title 参数自定义。')

# ============ 八、表格框图流程图 ============
add_h1(doc, "八、表格框图流程图")
add_body(doc, "流程图由 Word 表格与 Unicode 箭头构成，不使用图片，不使用 emoji，产物在任何 Word 版本中均可编辑。本章各构件均为实际渲染结果。")

add_h2(doc, "8.1 纵向流程图")
add_body(doc, "纵向流程图由横向箭头行与向下箭头交替堆叠而成，以信号处理流程为例：")
add_arrow_row(doc, "① 时域采样", "离散信号 x(n)")
add_arrow_down(doc)
add_arrow_row(doc, "② 加窗处理", "加窗信号")
add_arrow_down(doc)
add_arrow_row(doc, "③ 离散傅里叶变换", "频谱 X(k)", left_shade="D9D9D9")
add_arrow_down(doc)
add_arrow_row(doc, "④ 频谱分析", "特征向量", left_shade="D9D9D9")
add_arrow_down(doc)
add_separator_note(doc, "----- 信号预处理止于此处，后续由分类器承接 -----")
add_arrow_down(doc)
add_box(doc, "分类决策：特征归一化 / 模式匹配 / 结果输出")
add_fig_caption(doc, "信号处理与分类流程图")
add_note(doc, "注：左侧灰色底纹表示该步骤属于统一的预处理层级，右侧白色表示其产出为中间结果；分隔注释用于标记阶段边界。")

add_h2(doc, "8.2 横向箭头行")
add_body(doc, "横向箭头行是本技能的特色构件，用于表达过程到产出的转换。它采用 3 列表格实现，左右等宽并锁定固定布局，箭头恒位于整表几何中心，与两侧文字长短无关；中间竖线融合后左右单元格视觉连通。")

add_h3(doc, "底纹规则")
add_item_para(doc, "地位相等：", "左右不加底纹，两侧视觉权重一致。")
add_item_para(doc, "层级不同：", "传 left_shade 参数生成左灰右白效果，用于表达预处理阶段向决策阶段一类的层级转换。")

add_h3(doc, "几何居中与固定布局")
add_body(doc, "箭头位置由表格几何中心决定，因此左右文字长短悬殊时箭头仍然居中。固定布局同时保证换页、缩放或在不同 Word 版本中打开时列宽不发生重排。")

add_h2(doc, "8.3 并排框与多行框")
add_body(doc, "并排框用于表达地位相等的并列分支，多行框用于容纳多行文字。下图为两路并行处理在融合节点汇合的结构：")
add_multi_col_table(doc, [
    ("时域分支", "采样 / 加窗 / 自相关"),
    ("频域分支", "FFT / 谱估计 / 谱质心"),
])
add_arrow_down(doc)
add_multi_line_box(doc, ["特征融合节点", "拼接 / 加权 / 归一化"])
add_fig_caption(doc, "并行分支与特征融合")
add_body(doc, "横向箭头是独立居中一行的箭头，不与任何表格绑定，用于表达同一行内的前后关系：")
add_arrow_horizontal(doc)

add_h2(doc, "8.4 构件函数")
add_table_caption(doc, "流程图构件函数")
add_data_table(doc,
    ["函数", "用途"],
    [
        ("add_box", "单个居中框"),
        ("add_multi_line_box", "多行文字框"),
        ("add_multi_col_table", "并排框行"),
        ("add_arrow_down", "居中向下箭头，连接上下框"),
        ("add_arrow_horizontal", "独立居中横向箭头"),
        ("add_arrow_row", "横向箭头行，表达过程到产出的转换"),
        ("add_separator_note", "居中虚线分隔注释"),
    ],
    col_widths=[4.2, 10.2], font_size=9.5)

# ============ 九、分层架构图 ============
add_h1(doc, "九、分层架构图")
add_body(doc, "分层架构图以表格框图堆叠表达多层架构，全宽层与并列层可混合使用，层间自动插入间隔段落。全宽层与并列层的每个框首行均为加粗名称。整图总宽只有一个来源：传 col_widths 时其和即总宽，全宽层与并列层共用同一个总宽（上下框体必然等宽）；只传 width_cm（默认 14cm）时它即总宽，并列层按总宽均分。以信号处理系统为例：")
add_layered_architecture(doc, [
    ["采集层：传感器阵列 / 信号调理 / 抗混叠滤波"],
    ["预处理层：去直流 / 滤波 / 加窗"],
    [["变换模块", "快速傅里叶变换"],
     ["分析模块", "频带能量 / 谱质心"],
     ["决策模块", "模式匹配 / 阈值判断"]],
    ["应用层：频谱显示 / 报告输出"],
], col_widths=[4.0, 5.0, 5.0])
add_fig_caption(doc, "信号处理系统分层架构")

# ============ 十、图表自动编号 ============
add_h1(doc, "十、图表自动编号")
add_body(doc, "图标题与表标题由内部计数器自动编号，无需手动填写编号。图标题位于图下方，表标题位于表上方，两类编号相互独立。创建文档时自动重置计数器，每次生成文档均从图1与表1开始。图标题传入空字符串时生成纯间距段落且不递增编号，图注则用于在图下方补充说明并居中以斜体呈现。本文档中的图、表编号与图注均为自动生成。")

# ============ 十一、代码块与数据表 ============
add_h1(doc, "十一、代码块与数据表")
add_h2(doc, "11.1 代码块")
add_body(doc, "代码块以 Consolas 等宽字体渲染代码内容，默认 9pt，浅灰底纹，无首行缩进，逐行成段：")
add_code_block(doc, "python docxValidator.py output.docx --verbose")
add_h2(doc, "11.2 数据表")
add_body(doc, "数据表函数生成数据型表格，灰色表头黑体加粗居中，固定列宽，数据行末列左对齐、其余列居中。本文档中各表均为实际渲染效果。")
add_h2(doc, "11.3 单元格公式")
add_body(doc, "单元格公式函数向表格单元格插入行内 OMML 公式，适用于函数调用与渲染效果对照一类的场景，下表第三列为单元格内原生公式：")
add_table_caption(doc, "公式构建函数与渲染效果对照")
t = add_data_table(doc,
    ["构建函数", "调用形式", "渲染效果"],
    [
        ("下标", 'sub("a", "k")', ""),
        ("上标", 'sup("x", "2")', ""),
        ("分式", 'frac([r("η")], [r("N")])', ""),
        ("求和", 'sumOp([r("k")], [sub("a", "k")])', ""),
    ],
    col_widths=[2.4, 5.8, 6.2], font_size=9.5)
add_math_to_cell(t.cell(1, 2), inlineMath([sub("a", "k")]))
add_math_to_cell(t.cell(2, 2), inlineMath([sup("x", "2")]))
add_math_to_cell(t.cell(3, 2), inlineMath([frac([r("η")], [r("N")])]))
add_math_to_cell(t.cell(4, 2), inlineMath([sumOp([r("k")], [sub("a", "k")])]))
add_body(doc, "辅助函数 set_table_border 与 set_cell_shading 分别控制表格边框与单元格底纹。")

# ============ 十二、封面 ============
add_h1(doc, "十二、封面")
add_body(doc, "封面按需生成，默认不插入，由 add_cover_page 生成，须在写任何其它内容之前调用。封面为居中标题块：大标题黑体 26pt、副标题黑体 20pt、单位与日期宋体 15pt，其中副标题、单位、日期都可以省略，省略则该行不生成。本文档的封面即由该函数生成。")
add_table_caption(doc, "封面版式")
add_data_table(doc,
    ["元素", "字体", "字号", "段后"],
    [
        ("大标题", "黑体", "26pt", "8pt"),
        ("副标题", "黑体", "20pt", "60pt"),
        ("单位", "宋体", "15pt", "0"),
        ("日期", "宋体", "15pt", "段前 8pt"),
    ],
    col_widths=[2.6, 2.6, 2.8, 6.4], font_size=9.5)
add_body(doc, "标题块用顶部留白定位，而不是垂直居中：标题块距版心顶端一段按比例算出的距离，下方留白自然延伸。默认留白比例为 0.264，A4 纸张下留白 6.50cm，大标题段起点距页顶约 9.04cm（页面高度 30.4%）；行数少的封面大标题上缘位置不变，下方留白自然更大。留白高度不写死行数，而是先由纸张高与上下页边距算出版心高，再乘以封面留白比例，换算为若干固定行距的空段落，不足一段的余数挂到大标题段的段前。因此换纸张、改页边距或整套切换预设时，封面位置自动等比适配，不需要手工调整。留白空段使用固定行距，每段高度严格等于设定值，总留白与目标值之间的误差远小于 1 磅。留白比例上限为 0.5，超过半个版心会把标题块挤到第二页。")
add_body(doc, "封面独立成一节，不显示页码；其后的目录节同样不显示页码，正文从第 1 页起算。用了封面就不必再调用 add_title。留白由空段落构成，Word 与 WPS 正常渲染；若交付链路中经过会剥离空段落的 docx 转网页转换器，留白可能丢失。")

# ============ 十三、目录生成 ============
add_h1(doc, "十三、目录生成")
add_body(doc, '目录按需生成，默认不插入。标题函数使用 Word 内置 Heading 样式，add_toc 在封面之后、正文之前插入目录域，在 Word 中按住 Ctrl 并点击条目可跳转至对应标题。Word 打开文档时提示更新域，选择"是"即生成目录条目。本文档开头的目录即为实际效果。')
add_table_caption(doc, "目录条目格式")
add_data_table(doc,
    ["元素", "对应标题", "字体", "字号", "行距", "段前", "段后", "缩进"],
    [
        ("章标题行", "一级标题", "黑体", "小四 12pt", "固定 20pt", "6pt", "0", "居左不缩进"),
        ("一级节标题行", "二级标题", "宋体", "小四 12pt", "固定 20pt", "0", "0", "缩进 1 个汉字符"),
        ("二级节标题行", "三级标题", "宋体", "小四 12pt", "固定 20pt", "0", "0", "缩进 2 个汉字符"),
    ],
    col_widths=[2.4, 1.8, 1.5, 2.2, 1.9, 1.2, 1.2, 2.2], font_size=8.5)
add_body(doc, "目录标题为黑体三号居中，文档主标题不进入目录。默认层级为 1 到 3 级，传入 levels 参数可只收一至二级；打开文档时的域自动更新由 auto_update 参数控制，默认开启，设为 False 后改为在目录上右键手动更新。目录条目由域在 Word 更新时计算生成，更新前显示占位文字。目录末尾为分节符，与正文分隔并触发正文页码从 1 起算。目录条目的字体严格按上表呈现，不会被标题的黑体带跑：Word 更新域时会把标题文字的直接字符格式搬进条目，而本技能的标题一律不在文字上设字体、只把字符格式写在 Heading 样式里，Word 因此无从覆盖。")

# ============ 十四、页码 ============
add_h1(doc, "十四、页码")
add_body(doc, "封面页与目录页不显示页码，页码从正文第一页起算为 1。该行为由 start_body 实现：前置内容所在节的页脚被清空，正文另起一节并重置页码计数。本文档采用封面加目录的结构，三级分节由 add_cover_page 与 add_toc 自动完成，正文第一个标题显示的页码为 1。")
add_table_caption(doc, "分节的调用时机")
add_data_table(doc,
    ["文档结构", "是否需要手动调用 start_body", "说明"],
    [
        ("无封面、无目录", "不需要", "setup_document 已让页码从第 1 页起连续编号"),
        ("有目录（无论有无封面）", "不需要", "add_toc 末尾已自动调用 start_body"),
        ("有封面但无目录", "需要", "在正文第一个标题前手动调用 start_body"),
    ],
    col_widths=[3.6, 4.4, 6.4], font_size=9.5)
add_body(doc, "分节符类型为下一页，因此兼作封面或目录与正文之间的分页，不要再另加分页符。正文节通过 w:pgNumType 将页码起始值设为预设 page.body_page_start，默认 1，页脚沿用居中五号 Times New Roman 的 PAGE 域。重复调用 start_body 不会重复分节，第二次起为空操作。封面以分节符结束时末节仍是空的，此时 start_body 直接复用该节作为正文节而不再新建，add_toc 也会跳过自己的前导分页符，两者共同避免多出空白页。")

# ============ 十五、轻量文档验证 ============
add_h1(doc, "十五、轻量文档验证")
add_body(doc, "轻量验证脚本基于纯 Python 标准库实现，无外部依赖，在文档生成后执行 5 项结构检查，全部检查通过时退出码为 0，否则为 1：")
add_table_caption(doc, "轻量验证检查项")
add_data_table(doc,
    ["检查项", "说明"],
    [
        ("ZIP 完整性", "所有 ZIP 条目均可解压"),
        ("XML 格式良好性", "所有 XML 与关系文件可被解析"),
        ("文件引用完整性", "关系文件中每个目标在包内存在"),
        ("内容类型声明", "包内 XML 文件均有内容类型声明"),
        ("空白保留", "含首尾空白的文本节点均带保留属性"),
    ],
    col_widths=[3.6, 10.8], font_size=9.5)
add_code_block(doc, "python docxValidator.py output.docx --verbose")

# ============ 十六、安全解压与重打包 ============
add_h1(doc, "十六、安全解压与重打包")
add_body(doc, "安全解压函数在解压 .docx 时拒绝符号链接并校验路径边界，防止路径遍历攻击；重打包函数将目录重新打包为 .docx，确保 [Content_Types].xml 首位存储并原子写入目标文件。二者构成编辑现有文档的安全基础。")

# ============ 可选模块 ============
add_h1(doc, "十七、可选模块")
add_body(doc, "以下模块默认不启用、不检查环境、不安装依赖，仅当明确要求对应功能时使用。")
add_h2(doc, "17.1 PDF 渲染验证")
add_body(doc, "将生成的 .docx 转换为 PDF 进行视觉检查，依赖 LibreOffice：")
add_code_block(doc, "python scripts/optional/office/soffice.py convert output.docx --outdir ./pdf_out")
add_h2(doc, "17.2 XSD 模式验证")
add_body(doc, "对 .docx 内的 XML 部件执行 OOXML 模式深度验证，比轻量验证更严格，依赖 lxml：")
add_code_block(doc, "python scripts/optional/office/validate.py output.docx")
add_h2(doc, "17.3 编辑现有文档")
add_body(doc, "修改已有 .docx 文件内容，工作流为安全解压、合并碎片 run、编辑 XML、重新打包、验证。")
add_h2(doc, "17.4 追踪修订与批注")
add_body(doc, "accept_changes.py 通过 LibreOffice 宏接受文档中的所有追踪修订；comment.py 管理批注，维护 6 个批注相关文件的交叉链接系统，批注 XML 模板位于 scripts/optional/templates。")

# ============ 排版标准 ============
add_h1(doc, "十八、排版标准")
add_body(doc, "页面为 A4 纸张，页边距上下 2.54cm、左右 3.18cm，正文行距 1.5 倍。页脚居中放置页码，五号 Times New Roman，由 PAGE 域自动编号；含封面或目录时前置部分不显示页码，正文从第 1 页起算。封面标题块用顶部留白定位，留白高度按版心高比例换算，换纸张或页边距时自动等比适配。中西文字体分离：正文西文用 Times New Roman 衬线体，标题西文用 Arial 无衬线体。间距体系对齐 NJUThesis LaTeX 模板，该模板全局行距倍数为 1.625，与 Word 1.5 倍行距的视觉效果一致。")
add_table_caption(doc, "排版标准")
add_data_table(doc,
    ["元素", "字体", "字号", "样式", "段前", "段后"],
    [
        ("文档标题", "黑体", "一号 26pt", "加粗居中", "24pt", "18pt"),
        ("一级标题", "黑体", "三号 16pt", "加粗左对齐", "24pt", "6pt"),
        ("二级标题", "黑体", "小四 12pt", "加粗左对齐", "12pt", "6pt"),
        ("三级标题", "黑体", "小四 12pt", "加粗左对齐", "12pt", "6pt"),
        ("正文", "宋体", "小四 12pt", "两端对齐，首行缩进 2 字符", "0", "7pt"),
        ("分点段落", "宋体", "小四 12pt", "加粗标签加正文", "0", "7pt"),
        ("块级公式", "—", "—", "居中", "6pt", "6pt"),
        ("图标题", "宋体", "五号 10.5pt", "标签加粗，全角空格分隔，居中，位于图下方", "6pt", "12pt"),
        ("表标题", "宋体", "五号 10.5pt", "标签加粗，全角空格分隔，居中，位于表上方", "12pt", "6pt"),
        ("封面·大标题", "黑体", "26pt", "居中，位于顶部留白之后", "—", "8pt"),
        ("封面·副标题", "黑体", "20pt", "居中", "0", "60pt"),
        ("封面·单位与日期", "宋体", "15pt", "居中", "日期 8pt", "0"),
        ("目录·章标题行", "黑体", "小四 12pt", "居左，固定行距 20pt", "6pt", "0"),
        ("目录·一级节标题行", "宋体", "小四 12pt", "缩进 1 个汉字符，固定行距 20pt", "0", "0"),
        ("目录·二级节标题行", "宋体", "小四 12pt", "缩进 2 个汉字符，固定行距 20pt", "0", "0"),
    ],
    col_widths=[2.2, 1.5, 2.0, 5.5, 1.4, 1.4], font_size=9)

# ============ 环境要求 ============
add_h1(doc, "十九、环境要求")
add_body(doc, "必需依赖仅两项，Python 3.8 以上版本与 python-docx。可选依赖按需安装：LibreOffice 用于 PDF 渲染验证与接受追踪修订，lxml 用于 XSD 模式验证，Poppler 的 pdftoppm 用于 PDF 转图片检查，pandoc 用于读取现有文档内容。")
add_code_block(doc, "pip install python-docx")

# ============ 文件结构 ============
add_h1(doc, "二十、文件结构")
add_code_block(doc, """docx-formatter/
├── SKILL.md                    # 技能指令文件
├── README.md                   # 项目说明
├── README.docx                 # 排版效果样例（由 scripts/buildReadmeDocx.py 生成）
├── LICENSE
└── scripts/
    ├── docxBuilder.py          # 文档构建库，排版与公式插入
    ├── ommlBuilders.py         # OMML 数学元素构建器
    ├── formulaTemplates.py     # 预定义公式模板
    ├── docxValidator.py        # 轻量验证脚本
    ├── buildReadmeDocx.py      # 标准骨架范例，写自己的脚本前先读它
    └── optional/               # 可选功能脚本
        ├── merge_runs.py       # 合并碎片 run
        ├── accept_changes.py   # 接受所有追踪修订
        ├── comment.py          # 批注管理
        ├── templates/          # 批注 XML 模板
        └── office/             # 验证与转换工具
            ├── soffice.py      # LibreOffice 跨平台调用
            ├── validate.py     # XSD 模式验证入口
            ├── helpers/        # 通用辅助函数包
            ├── schemas/        # OOXML 模式文件
            └── validators/     # 验证器""")

# ============ 安装与使用 ============
add_h1(doc, "二十一、安装与使用")
add_h2(doc, "21.1 安装")
add_body(doc, "方式一：npx skills（推荐）。在命令行执行：")
add_code_block(doc, "npx skills add https://github.com/CP-here/docx-formatter-skill --skill docx-formatter")
add_body(doc, "方式二：复制到项目目录，随仓库携带。若希望某个项目内的 AI 助手都遵循本技能的排版规则，把 docx-formatter 目录复制到该项目的技能目录，目录名与 SKILL.md 中 name 字段保持一致：")
add_table_caption(doc, "项目级技能目录")
add_data_table(doc,
    ["客户端", "项目级路径"],
    [
        ("Claude Code", "<项目>/.claude/skills/docx-formatter/"),
        ("Codex", "<项目>/.codex/skills/docx-formatter/"),
        ("Cursor", "<项目>/.cursor/skills/docx-formatter/"),
        ("WorkBuddy", "<项目>/.workbuddy/skills/docx-formatter/"),
        ("其他框架", "<项目> 下助手能发现 SKILL.md 的任意位置（须放在项目文件夹内）"),
    ],
    col_widths=[3.0, 11.4], font_size=9.5)
add_body(doc, "放在项目级的好处是技能随仓库走，提交后每个打开该项目的人都能自动获得这套排版能力。")
add_body(doc, "方式三：复制到用户级技能目录。复制到当前用户的技能目录，所有项目通用：")
add_table_caption(doc, "用户级技能目录")
add_data_table(doc,
    ["客户端", "用户级路径"],
    [
        ("Claude Code", "~/.claude/skills/docx-formatter/"),
        ("Codex", "~/.codex/skills/docx-formatter/"),
        ("Cursor", "~/.cursor/skills/docx-formatter/"),
        ("WorkBuddy", "~/.workbuddy/skills/docx-formatter/"),
        ("其他框架", "当前用户主目录下助手能发现 SKILL.md 的任意位置"),
    ],
    col_widths=[3.0, 11.4], font_size=9.5)
add_body(doc, "放在用户级的好处是当前用户的所有项目都能用，无需逐个项目复制；安装后在对话中以斜杠引用该技能，例如 /docx-formatter。安装完成后重启助手会话，使技能索引刷新。")

add_h2(doc, "21.2 使用方式")
add_body(doc, "安装后在对话中直接描述需求即可触发技能，例如生成一份带封面与目录的 Word 文档、排版一份技术报告为 Word、用数学方程格式化这段文档、在 Word 中创建流程图。可选功能需明确要求，例如将文档转为 PDF 查看效果、对文档添加批注、以追踪修订方式修改现有文档。")

# ============ 许可证与致谢 ============
add_h1(doc, "二十二、许可证")
add_body(doc, "本项目采用 MIT 许可证，可自由使用、修改和分发。")
add_h1(doc, "二十三、致谢")
add_link_para(doc, [
    ("text", "排版部分参考 "),
    ("link", "https://github.com/nju-lug/NJUThesis", "NJUThesis"),
    ("text", " LaTeX 模板，部分参考"),
    ("link", "https://grs.pku.edu.cn/docs/2019-03/20190304101034750506.pdf", "北京大学研究生学位论文写作指南"),
])
add_link_para(doc, [
    ("text", "可选功能参考 "),
    ("link", "https://github.com/anthropics/skills/tree/main/skills/docx", "Anthropic 官方 docx 技能"),
])

out = os.path.join(PROJECT_ROOT, "README.docx")
doc.save(out)
print("saved:", out)
