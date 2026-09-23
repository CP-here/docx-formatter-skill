# -*- coding: utf-8 -*-
"""端到端：跑真实构建链路并检查产物结构。

覆盖「文档结构」「OMML 输出」「代表性排版工作流」三项。

构建方式刻意复刻 SKILL.md 文档化的标准流程 —— 把库文件复制到工作目录
再运行。这样既不修改任何源码（构建脚本的输出路径是写死的），也不会
污染仓库工作区（产物落在 pytest 临时目录里）。
"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

import pytest

from conftest import ASSETS, REPO_ROOT, SCRIPTS
from docx_validator import validate_docx

# SKILL.md「标准流程」要求复制的文件（不含 scripts/optional/）
_COPIED = [
    "docx_layout_kit.py",
    "omml_math_kit.py",
    "formula_templates.py",
    "_build_readme_docx.py",
]

# 封面日期取自 datetime.date.today()（_build_readme_docx.py 第 110 行），
# 每次构建都不同；比对前归一化，否则每月 1 号同步测试必然失败。
_DATE_RE = re.compile(r"\d{4}\u5e74\d{1,2}\u6708")


@pytest.fixture(scope="module")
def built_readme(tmp_path_factory):
    """在临时目录中按标准流程构建 README.docx。"""
    tmp = tmp_path_factory.mktemp("readme-build")

    sdir = tmp / "scripts"
    sdir.mkdir()
    for name in _COPIED:
        shutil.copy(str(SCRIPTS / name), str(sdir / name))

    adir = tmp / "assets"
    adir.mkdir()
    shutil.copy(str(ASSETS / "effect-overview.png"), str(adir / "effect-overview.png"))

    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run(
        [sys.executable, str(sdir / "_build_readme_docx.py")],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )
    assert proc.returncode == 0, "构建失败：\n%s\n%s" % (proc.stdout, proc.stderr)

    out = tmp / "README.docx"
    assert out.exists(), "构建返回成功但未产出 README.docx"
    return out


def _document_xml(path):
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read("word/document.xml").decode("utf-8")


def _body_paragraphs(path):
    """逐段提取可见文本，并把动态封面日期归一化。"""
    xml = _document_xml(path)
    paras = re.findall(r"<w:p[ >].*?</w:p>|<w:p/>", xml, re.S)
    out = []
    for para in paras:
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para, re.S))
        out.append(_DATE_RE.sub("<DATE>", text))
    return out


def test_build_succeeds(built_readme):
    """代表性工作流端到端跑通并产出非平凡大小的文件。"""
    size = built_readme.stat().st_size
    assert size > 3000, "产物仅 %d 字节，空 docx 约 3KB" % size


def test_output_passes_own_validator(built_readme):
    """产物必须通过本仓库自带的轻量完整性验证（5 项结构检查）。"""
    assert validate_docx(str(built_readme)) is True


def test_output_contains_native_omath(built_readme):
    """公式必须是 Word 原生 OMML 对象，而不是图片或纯文本。"""
    count = _document_xml(built_readme).count("<m:oMath")
    assert count >= 4, "<m:oMath 出现 %d 次，期望 ≥4" % count


def test_output_has_toc_field(built_readme):
    """必须插入 TOC 目录域（而非手写目录文本）。"""
    doc = _document_xml(built_readme)
    toc = re.findall(r"<w:instrText[^>]*>[^<]*TOC", doc)
    assert len(toc) >= 1, "缺少 TOC 目录域"


def test_output_has_heading_styles(built_readme):
    """标题须走 Word 内置 Heading 样式（否则进不了目录）。

    不断言 Heading3 —— 实测当前 README.docx 未使用三级标题。
    """
    doc = _document_xml(built_readme)
    with zipfile.ZipFile(str(built_readme)) as zf:
        styles = zf.read("word/styles.xml").decode("utf-8")

    for level, minimum in (("Heading1", 3), ("Heading2", 3)):
        assert 'w:val="%s"' % level in styles, "styles.xml 未定义 %s" % level
        used = len(re.findall(r'<w:pStyle[^>]*w:val="%s"' % level, doc))
        assert used >= minimum, "%s 实际使用 %d 次，期望 ≥%d" % (level, used, minimum)


def test_output_has_superscript_citations(built_readme):
    """正文引用标记须渲染为上标（README 列出的卖点之一）。"""
    count = _document_xml(built_readme).count('w:val="superscript"')
    assert count >= 3, "上标出现 %d 处，期望 ≥3" % count


def test_output_has_tables(built_readme):
    """表格框图与数据表须真实生成 <w:tbl> 元素。"""
    count = _document_xml(built_readme).count("<w:tbl>")
    assert count >= 5, "表格 %d 个，期望 ≥5" % count


def test_committed_readme_in_sync(built_readme):
    """抓「改了构建脚本/库却忘了重新生成 README.docx」。

    只比正文文本与段落数，不比字节 —— docx 是 zip，
    docProps 时间戳每次构建都不同，字节比对必然假阳性。
    """
    committed = REPO_ROOT / "README.docx"
    assert committed.exists(), "README.docx 应已提交到仓库"

    fresh = _body_paragraphs(built_readme)
    done = _body_paragraphs(committed)

    assert len(fresh) == len(done), "段落数不一致：新构建 %d，已提交 %d" % (
        len(fresh),
        len(done),
    )

    diff = [(i, a, b) for i, (a, b) in enumerate(zip(fresh, done)) if a != b]
    sample = "\n".join(
        "  段落 %d:\n    新构建: %r\n    已提交: %r" % (i, a, b) for i, a, b in diff[:5]
    )
    assert not diff, (
        "README.docx 与构建脚本不同步 —— "
        "运行 python scripts/_build_readme_docx.py 后重新提交\n%s" % sample
    )
