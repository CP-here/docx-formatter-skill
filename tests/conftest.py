# -*- coding: utf-8 -*-
"""pytest 共享配置：导入路径注入 + 最小 .docx 构造 fixture。

conftest.py 由 pytest 自动加载，且先于同目录下的测试模块，
因此这里的 sys.path 注入能保证测试文件顶层的
`from omml_math_kit import ...` 成功执行。

make_docx 以 zipfile 手写最小 .docx 包，不引入 python-docx，
使 test_omml_math_kit.py 与 test_docx_validator.py 除 pytest 外
无第三方依赖。整个测试套件仍依赖 python-docx，因为
test_build_readme.py 需端到端运行依赖 python-docx 的构建脚本。
"""
import sys
import zipfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = REPO_ROOT / "scripts"
ASSETS = REPO_ROOT / "assets"

sys.path.insert(0, str(SCRIPTS))

# ---- 最小 .docx 的四个部件 -------------------------------------------------
# 最小包结构自洽，可通过 docx_validator 全部 5 项检查；
# 每个 break_* 开关仅破坏其中一项，便于定位。

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

# 正文唯一一个 w:t 带边空格且声明 xml:space="preserve"（check 5 默认通过）
_DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve"> 正文 </w:t></w:r></w:p>
  </w:body>
</w:document>
"""

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
"""


@pytest.fixture
def make_docx(tmp_path):
    """生成一个内部自洽的最小 .docx，可用开关制造特定损坏。

    每个开关对应 docx_validator 的某一项 check：
        break_xml        → check 2  XML 合法性
        break_rel        → check 3  .rels 指向不存在的部件
        break_whitespace → check 5  边空格缺 xml:space="preserve"
    """
    counter = {"n": 0}

    def _make(break_xml=False, break_rel=False, break_whitespace=False):
        counter["n"] += 1
        path = tmp_path / ("case%d.docx" % counter["n"])

        document = _DOCUMENT
        if break_whitespace:
            document = document.replace('xml:space="preserve"', "")
        if break_xml:
            document = document.replace("</w:document>", "")

        rels = _DOC_RELS
        if break_rel:
            rels = rels.replace(
                "</Relationships>",
                '  <Relationship Id="rId99" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/styles" Target="missing.xml"/>\n</Relationships>',
            )

        with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
            zf.writestr("_rels/.rels", _ROOT_RELS)
            zf.writestr("word/document.xml", document)
            zf.writestr("word/_rels/document.xml.rels", rels)
            zf.writestr("word/styles.xml", _STYLES)
        return str(path)

    return _make
