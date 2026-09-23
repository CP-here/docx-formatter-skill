# -*- coding: utf-8 -*-
"""docx_validator.py 的单元测试。

正面用例验证完好包通过；负面用例验证各检查项能被触发，
防止校验器退化为始终通过的空壳。

除 pytest 外无第三方依赖。
"""
import os
import subprocess
import sys

from conftest import SCRIPTS
from docx_validator import validate_docx


def test_valid_docx_passes(make_docx):
    """完好的最小包应全项通过。"""
    assert validate_docx(make_docx()) is True


def test_malformed_xml_fails(make_docx):
    """check 2：未闭合根元素应被检出。"""
    assert validate_docx(make_docx(break_xml=True)) is False


def test_missing_rel_target_fails(make_docx):
    """check 3：.rels 指向包内不存在的部件应被检出。"""
    assert validate_docx(make_docx(break_rel=True)) is False


def test_edge_whitespace_without_preserve_fails(make_docx):
    """check 5：带边空格的 w:t 缺 xml:space='preserve' 应被检出。"""
    assert validate_docx(make_docx(break_whitespace=True)) is False


def test_cli_exit_codes(make_docx):
    """验证 CLI 退出码契约（0=通过，1=失败）。

    固定 PYTHONIOENCODING=utf-8，避免 Windows 子进程按本地代码页
    输出导致解码失败。
    """
    env = dict(os.environ, PYTHONIOENCODING="utf-8")

    ok = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_validator.py"), make_docx()],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )
    assert ok.returncode == 0, "完好文件应返回 0：\n%s\n%s" % (ok.stdout, ok.stderr)

    bad = subprocess.run(
        [
            sys.executable,
            str(SCRIPTS / "docx_validator.py"),
            make_docx(break_xml=True),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )
    assert bad.returncode == 1, "损坏文件应返回 1：\n%s\n%s" % (bad.stdout, bad.stderr)
