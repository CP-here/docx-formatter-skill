# -*- coding: utf-8 -*-
"""docx_validator.py 的单元测试。

重点不是证明校验器「能通过」，而是证明它「能发现问题」——
一个停止检测目标的校验器会永远报告通过，比没有校验器更危险。

零第三方依赖（pytest 之外）。
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
    """check 2：未闭合的根元素必须被抓住。"""
    assert validate_docx(make_docx(break_xml=True)) is False


def test_missing_rel_target_fails(make_docx):
    """check 3：.rels 指向包内不存在的部件必须被抓住。"""
    assert validate_docx(make_docx(break_rel=True)) is False


def test_edge_whitespace_without_preserve_fails(make_docx):
    """check 5：带边空格的 w:t 缺 xml:space='preserve' 必须被抓住。"""
    assert validate_docx(make_docx(break_whitespace=True)) is False


def test_cli_exit_codes(make_docx):
    """退出码是 CI 与外部调用方唯一能依赖的契约（0=通过，1=失败）。

    用 PYTHONIOENCODING 固定子进程输出编码，避免 Windows 上按本地
    代码页写出导致解码失败。
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
