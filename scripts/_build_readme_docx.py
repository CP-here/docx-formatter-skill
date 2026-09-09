# -*- coding: utf-8 -*-
"""将 README 渲染为 docx，特有结构以原生形式呈现"""
import sys, os
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from build_docx import (setup_document, add_title, add_toc, add_h1, add_h2,
    add_body, add_item_para, add_eq_para, add_body_with_math, add_code_block,
    add_data_table, add_math_to_cell, add_box, add_multi_line_box,
    add_arrow_down, add_arrow_row, add_separator_note, add_fig_caption,
    add_table_caption, add_layered_architecture, add_note, add_bibliography)
from mathHelpers import r, sub, sup, frac, sumOp, func, math, inlineMath

doc = setup_document()
add_title(doc, "docx-formatter 中文 Word 专业排版技能")
add_toc(doc)

# ============ 一、项目简介 ============
add_h1(doc, "一、项目简介")
add_body(doc, "docx-formatter 生成专业排版的中文 Word 文档，覆盖章节结构、正文、数学公式、表格、流程图、目录、引用上标与文档验证的完整排版链路，输出符合学术与工程规范的 .docx 文件。排版效果可参考仓库中的 example.docx。")
add_body(doc, "技能采用单引擎架构，仅依赖 Python 3.8 以上版本与 python-docx。数学公式由 mathHelpers.py 直接生成 OMML XML，即 Word 原生公式格式，公式在 Word 中可二次编辑，无需 Node.js 或任何外部转换工具。")
add_body(doc, "本文档即由该技能自身排版生成：文中所有公式为 Word 原生可编辑公式，流程图与架构图为表格框图，引用标记为真实上标，可作为排版效果的直接样例。")

# ============ 二、功能矩阵 ============
add_h1(doc, "二、功能矩阵")
add_body(doc, "技能共包含 18 个模块，其中 14 个默认启用，4 个为可选模块，仅在明确要求时启用。")
add_table_caption(doc, "功能矩阵总览")
add_data_table(doc,
    ["模块", "默认状态", "说明"],
    [
        ("标准排版", "启用", "标题、正文、缩进与行距体系"),
        ("排版预设", "启用", "整套切换排版参数，支持中西文字体分离"),
        ("OMML 数学公式", "启用", "行内与块级公式，Word 原生可编辑"),
        ("引用上标", "启用", "参考文献标记自动渲染为上标"),
        ("表格框图流程图", "启用", "表格与 Unicode 箭头构建流程图"),
        ("分层架构图", "启用", "多层架构的表格框图堆叠"),
        ("图表自动编号", "启用", "图表标题自动递增编号"),
        ("代码块", "启用", "等宽字体与浅灰底纹"),
        ("数据表", "启用", "灰色表头与固定列宽"),
        ("单元格公式", "启用", "表格单元格内插入行内公式"),
        ("目录生成", "按需", "Heading 样式加目录域"),
        ("轻量文档验证", "启用", "纯标准库 5 项检查，生成后默认执行"),
        ("安全解压与重打包", "启用", "编辑现有文档的安全基础"),
        ("PDF 渲染验证", "可选", "LibreOffice 转换后视觉检查"),
        ("XSD 模式验证", "可选", "OOXML 标准深度验证"),
        ("编辑现有文档", "可选", "解压、编辑与重打包工作流"),
        ("追踪修订", "可选", "接受全部追踪修订"),
        ("批注管理", "可选", "批注插入与交叉链接维护"),
    ],
    col_widths=[3.4, 1.6, 9.4], font_size=9)

# ============ 三、标准排版 ============
add_h1(doc, "三、标准排版")
add_body(doc, "标题使用黑体，正文使用宋体。文档标题一号 26pt 居中加粗，一级标题三号 16pt 左对齐，二级与三级标题小四 12pt。正文小四 12pt 两端对齐，首行缩进 2 字符，全文 1.5 倍行距。纸张为 A4，页边距上下 2.54cm、左右 3.18cm，页脚居中页码为五号 Times New Roman。各级标题颜色显式覆盖为黑色，规避 Word 内置 Heading 样式的默认蓝色。中西文字体分离：正文西文用 Times New Roman 衬线体，标题西文用 Arial 无衬线体。本文档各级标题与正文即为标准排版的实际效果。")
add_table_caption(doc, "标准排版构件函数")
add_data_table(doc,
    ["函数", "用途"],
    [
        ("setup_document", "创建文档并应用页面设置，自动重置图表计数器"),
        ("add_title", "文档主标题，黑体三号居中加粗"),
        ("add_h1、add_h2、add_h3", "一级、二级、三级标题，使用 Word 内置 Heading 样式，可进入目录"),
        ("add_body", "正文段落，支持加粗标记与引用上标"),
        ("add_item_para", "加粗标签加分点内容的段落，标签整体加粗"),
    ],
    col_widths=[4.2, 10.2], font_size=9.5)

# ============ 四、排版预设 ============
add_h1(doc, "四、排版预设")
add_body(doc, "排版参数集中在 build_docx.py 顶部的 PRESETS 字典中，按标题、正文、分点、图注、页面、目录等分组管理，各排版函数统一读取当前激活的预设。")
add_item_para(doc, "set_preset 整套切换：", "须在 setup_document 之前调用，默认预设与标准排版一致，新增预设只需向 PRESETS 追加一份同结构的嵌套字典。")
add_item_para(doc, "中西文字体分离：", "开启后西文使用 Times New Roman，中文保持原有字体不变。")

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
        "[",
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
        "]",
    ]),
]))
add_fig_caption(doc, "似然比检测损失的嵌套公式")

add_h2(doc, "5.4 希腊字母与数学符号")
add_body(doc, "希腊字母与数学符号以 Unicode 直接书写，formulas.py 内置若干常用公式模板，可直接插入文档，也可作为自定义公式的参考模式。求和符号必须使用 sumOp 构建：求和若采用 n 元运算符加空上标的写法，Word 会渲染一个不可见的上标占位框，导致文件看起来损坏；sumOp 改用下标结构配合 Unicode ∑ 实现，彻底规避该问题。所有接收子元素的函数内部自动展平嵌套列表，防止 XML 拼接错误。")

# ============ 六、引用上标 ============
add_h1(doc, "六、引用上标")
add_body(doc, "正文中的参考文献标记自动渲染为上标，适用于 add_body 与 add_item_para，支持单个引用、多个引用与范围引用三种格式。本段即为实际渲染效果：傅里叶分析方法在信号处理任务中表现出色[1]，后续研究[2,3]进一步验证了这一结论，系统综述见[1-3]。")

# ============ 七、参考文献列表 ============
add_h1(doc, "七、参考文献列表")
add_body(doc, "add_bibliography 生成 GB/T 7714-2015 顺序编码制参考文献表，格式对齐 NJUThesis LaTeX 模板，与正文引用上标配套使用：正文中的 [1] 指向条目列表第一项，以此类推。")
add_body(doc, "章节标题为黑体三号加粗居中，使用 Heading 1 样式，进入目录；条目为宋体五号，两端对齐，1.5 倍行距，悬挂缩进 2 字符，折行后与序号后文字对齐；序号 [1]、[2] 按传入顺序自动生成，无需手写。下方即本文档的参考文献表实际渲染效果：")
add_bibliography(doc, [
    "程佩青. 数字信号处理教程[M]. 4版. 北京: 清华大学出版社, 2015",
    "陈大明, 刘思远. 基于加窗傅里叶变换的谐波检测方法[J]. 信号处理学报, 2023, 41(2): 115-128",
    "Oppenheim A V, Willsky A S, Nawab S H. Signals and Systems[M]. 2nd ed. Upper Saddle River: Prentice Hall, 1997",
    "周文静. 短时傅里叶变换在振动信号分析中的应用[D]. 南京: 东南大学, 2021",
    "Zhang W, Liu H. Window function selection for spectral leakage suppression[C]//Proceedings of the International Conference on Signal Processing. Beijing: IEEE, 2021: 45-52",
    "全国信息与文献标准化技术委员会. 信息与文献 参考文献著录规则: GB/T 7714-2015[S]. 北京: 中国标准出版社, 2015",
])

# ============ 八、表格框图流程图 ============
add_h1(doc, "八、表格框图流程图")
add_body(doc, "流程图由 Word 表格与 Unicode 箭头构成，不使用图片，不使用 emoji，产物在任何 Word 版本中均可编辑。以信号处理流程为例，下方流程图即为表格框图的真实渲染结果：")
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
add_body(doc, "横向箭头行是本技能的特色构件。采用 3 列表格实现，左右等宽并锁定固定布局，箭头恒位于整表几何中心，与两侧文字长短无关；中间竖线融合后左右单元格视觉连通。左右地位相等时不加底纹，层级不同时左灰右白，上图中前两步地位相等为全白，后两步表达从预处理阶段到变换阶段的层级转换。")
add_table_caption(doc, "流程图构件函数")
add_data_table(doc,
    ["函数", "用途"],
    [
        ("add_box", "单个居中框"),
        ("add_multi_line_box", "多行文字框"),
        ("add_multi_col_table", "并排框行"),
        ("add_arrow_down", "居中向下箭头，连接上下框"),
        ("add_arrow_row", "横向箭头行，表达过程到产出的转换"),
        ("add_separator_note", "居中虚线分隔注释"),
    ],
    col_widths=[4.2, 10.2], font_size=9.5)

# ============ 八、分层架构图 ============
add_h1(doc, "九、分层架构图")
add_body(doc, "add_layered_architecture 以表格框图堆叠表达多层架构，全宽层与并列层可混合使用，层间自动插入间隔段落。全宽层与并列层的每个框首行均为加粗名称，col_widths 控制并列框列宽，不传则按默认总宽 14cm 均分。以信号处理系统为例：")
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
add_body(doc, "图标题与表标题由内部计数器自动编号，无需手动填写编号。图标题位于图下方，表标题位于表上方，两类编号相互独立。setup_document 自动调用 reset_counters，每次生成文档均从图1与表1开始。add_fig_caption 传入空字符串时生成纯间距段落且不递增编号。add_note 在图下方添加居中斜体图注。本文档中的图与表编号均为自动生成。")

# ============ 十、代码块与数据表 ============
add_h1(doc, "十一、代码块与数据表")
add_h2(doc, "11.1 代码块")
add_body(doc, "add_code_block 以 Consolas 等宽字体渲染代码内容，默认 9pt，浅灰底纹，无首行缩进，逐行成段：")
add_code_block(doc, "python validate_docx.py output.docx --verbose")
add_h2(doc, "11.2 数据表")
add_body(doc, "add_data_table 生成数据型表格，灰色表头黑体加粗居中，固定列宽，数据行末列左对齐、其余列居中。本文档中各表均为实际渲染效果。")
add_h2(doc, "11.3 单元格公式")
add_body(doc, "add_math_to_cell 向表格单元格插入行内 OMML 公式，适用于构建函数与渲染效果对照一类的场景，下表第三列为单元格内原生公式：")
add_table_caption(doc, "公式构建函数与渲染效果对照")
t = add_data_table(doc,
    ["构建函数", "调用形式", "渲染效果"],
    [
        ("下标", "sub a k", ""),
        ("上标", "sup x 2", ""),
        ("分式", "frac η N", ""),
        ("求和", "sumOp Σ", ""),
    ],
    col_widths=[2.6, 5.0, 6.8], font_size=9.5)
add_math_to_cell(t.cell(1, 2), inlineMath([sub("a", "k")]))
add_math_to_cell(t.cell(2, 2), inlineMath([sup("x", "2")]))
add_math_to_cell(t.cell(3, 2), inlineMath([frac([r("η")], [r("N")])]))
add_math_to_cell(t.cell(4, 2), inlineMath([sumOp([r("k")], [sub("a", "k")])]))
add_body(doc, "辅助函数 set_table_border 与 set_cell_shading 分别控制表格边框与单元格底纹。")

# ============ 十二、目录生成 ============
add_h1(doc, "十二、目录生成")
add_body(doc, "目录按需生成，默认不插入。标题函数使用 Word 内置 Heading 样式，add_toc 在文档标题后、正文前插入目录域，在 Word 中按住 Ctrl 并点击条目可跳转至对应标题。目录标题黑体三号居中，一级条目黑体四号，二级条目宋体小四，行距固定 22pt，样式对齐 NJUThesis 模板。默认层级为 1 到 2 级，文档主标题不进入目录。目录域需在 Word 中右键选择更新域后生成条目，目录后自动分页。本文档开头的目录即为实际效果。")

# ============ 十三、轻量文档验证 ============
add_h1(doc, "十三、轻量文档验证")
add_body(doc, "validate_docx.py 基于纯 Python 标准库实现，无外部依赖，在文档生成后执行 5 项结构检查，全部通过时退出码为 0，否则为 1：")
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
add_code_block(doc, "python validate_docx.py output.docx --verbose")

# ============ 十四、安全解压与重打包 ============
add_h1(doc, "十四、安全解压与重打包")
add_body(doc, "safe_extract 在解压 .docx 时拒绝符号链接并校验路径边界，防止路径遍历攻击；rezip 将目录重新打包为 .docx，确保 [Content_Types].xml 首位存储并原子写入目标文件。二者构成编辑现有文档的安全基础。")

# ============ 可选模块 ============
add_h1(doc, "十五、可选模块")
add_body(doc, "以下模块默认不启用、不检查环境、不安装依赖，仅当明确要求对应功能时使用。")
add_h2(doc, "15.1 PDF 渲染验证")
add_body(doc, "将生成的 .docx 转换为 PDF 进行视觉检查，依赖 LibreOffice：")
add_code_block(doc, "python scripts/optional/office/soffice.py convert output.docx --outdir ./pdf_out")
add_h2(doc, "15.2 XSD 模式验证")
add_body(doc, "对 .docx 内的 XML 部件执行 OOXML 模式深度验证，比轻量验证更严格，依赖 lxml：")
add_code_block(doc, "python scripts/optional/office/validate.py output.docx")
add_h2(doc, "15.3 编辑现有文档")
add_body(doc, "修改已有 .docx 文件内容，工作流为安全解压、合并碎片 run、编辑 XML、重新打包、验证。")
add_h2(doc, "15.4 追踪修订与批注")
add_body(doc, "accept_changes.py 通过 LibreOffice 宏接受文档中的所有追踪修订；comment.py 管理批注，维护 6 个批注相关文件的交叉链接系统，批注 XML 模板位于 scripts/optional/templates。")

# ============ 排版标准 ============
add_h1(doc, "十六、排版标准")
add_body(doc, "页面为 A4 纸张，页边距上下 2.54cm、左右 3.18cm，正文行距 1.5 倍。页脚居中放置页码，五号 Times New Roman，由 PAGE 域自动编号。中西文字体分离：正文西文用 Times New Roman 衬线体，标题西文用 Arial 无衬线体。间距体系对齐 NJUThesis LaTeX 模板与北京大学研究生学位论文写作指南，NJUThesis 全局行距倍数为 1.625，与 Word 1.5 倍行距的视觉效果一致。")
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
        ("目录条目", "黑体或宋体", "四号或小四", "固定行距 22pt", "0", "0"),
    ],
    col_widths=[2.0, 2.0, 2.2, 5.0, 1.6, 1.6], font_size=9)

# ============ 环境要求 ============
add_h1(doc, "十七、环境要求")
add_body(doc, "必需依赖仅两项，Python 3.8 以上版本与 python-docx。可选依赖按需安装：LibreOffice 用于 PDF 渲染验证与接受追踪修订，lxml 用于 XSD 模式验证，Poppler 的 pdftoppm 用于 PDF 转图片检查，pandoc 用于读取现有文档内容。")
add_code_block(doc, "pip install python-docx")

# ============ 文件结构 ============
add_h1(doc, "十八、文件结构")
add_code_block(doc, """docx-formatter/
├── SKILL.md                    # 技能指令文件
├── README.md                   # 项目说明
├── README.docx                 # 排版效果样例
├── LICENSE
└── scripts/
    ├── build_docx.py           # 文档构建模板，排版与公式插入
    ├── mathHelpers.py          # OMML 数学元素构建器
    ├── formulas.py             # 预定义公式模板
    ├── validate_docx.py        # 轻量验证脚本
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
add_h1(doc, "十九、安装与使用")
add_body(doc, "通过客户端安装：将技能压缩包上传至 AI 客户端对应的技能安装位置，客户端自动读取 SKILL.md 中的指令。手动安装：将仓库内容复制到项目的 skills/docx-formatter 目录，目录名与 SKILL.md 中 name 字段保持一致。")
add_body(doc, "安装后在对话中直接描述需求即可触发技能，例如生成一份带公式的 Word 文档、排版一份技术报告为 Word、用数学方程格式化这段文档、在 Word 中创建流程图。可选功能需明确要求，例如将文档转为 PDF 查看效果、对文档添加批注、以追踪修订方式修改现有文档。")

# ============ 许可证与致谢 ============
add_h1(doc, "二十、许可证")
add_body(doc, "本项目可自由使用和修改。")
add_h1(doc, "二十一、致谢")
add_body(doc, "排版部分标准对齐 NJUThesis LaTeX 模板，部分对齐北京大学研究生学位论文写作指南；可选功能参考 Anthropic 官方 docx 技能。")

out = os.path.join(os.path.dirname(BASE), "README.docx")
doc.save(out)
print("saved:", out)
