"""
Render a .docx to PDF with the Microsoft Word installed on this machine (COM automation).

Windows only: needs Microsoft Word plus pywin32 (``pip install pywin32``).
Unlike the LibreOffice path (soffice.py), the file goes through the real Word
layout engine, so the PDF matches what a user sees on opening the .docx --
中文字体、OMML 公式、目录域、页脚页码分节都按 Word 自身的渲染结果输出。

Word 内的域（TOC / PAGE）在导出前会先 Update，因此目录不会是占位文字。

Usage:
    python word2pdf.py output.docx                  # -> output.pdf（与源文件同目录）
    python word2pdf.py output.docx -o dist/a.pdf    # 指定输出文件
    python word2pdf.py output.docx --outdir pdf_out # 指定输出目录
    python word2pdf.py output.docx --pdfa           # 导出 PDF/A-1b

编程式调用::

    from word2pdf import convert_docx_to_pdf
    convert_docx_to_pdf("output.docx", "output.pdf")

实现细节：每次使用独立的 Word 实例（DispatchEx），文档只读打开，导出后
关闭文档并退出该实例，因此不会干扰用户已经打开着的 Word 窗口；导出结束后
显式释放 COM 引用，避免留下后台 WINWORD.EXE 进程。
"""

import argparse
import gc
import sys
from pathlib import Path


# Word 枚举值（避免依赖 pywin32 生成的 makepy 常量模块）
WD_EXPORT_FORMAT_PDF = 17        # wdExportFormatPDF
WD_EXPORT_OPTIMIZE_PRINT = 0     # wdExportOptimizeForPrint
WD_EXPORT_ALL_DOCUMENT = 0       # wdExportAllDocument
WD_EXPORT_DOCUMENT_CONTENT = 0   # wdExportDocumentContent（正文，不含批注/修订页）
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1  # wdExportCreateHeadingBookmarks
WD_DO_NOT_SAVE_CHANGES = 0       # wdDoNotSaveChanges
WD_ALERTS_NONE = 0               # wdAlertsNone

IS_WINDOWS = sys.platform == "win32"


def _fail(message: str) -> None:
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(2)


def _load_word_module():
    """Import pywin32 lazily so --help works without the dependency installed."""
    if not IS_WINDOWS:
        _fail(
            "Word 渲染仅支持 Windows；其它平台请改用 "
            "scripts/optional/office/soffice.py（LibreOffice 转换）"
        )
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        _fail("未安装 pywin32，请先执行：pip install pywin32")
    return pythoncom, win32


def convert_docx_to_pdf(
    src,
    dst=None,
    *,
    pdfa: bool = False,
    update_fields: bool = True,
    verbose: bool = True,
):
    """用本机 Word 把 *src* 渲染为 PDF，返回输出文件的 :class:`~pathlib.Path`。

    Args:
        src: 源 .docx 路径。
        dst: 输出 .pdf 路径；省略时与源文件同目录、同名。
        pdfa: 为 True 时导出 PDF/A-1b（ISO 19005-1）。
        update_fields: 导出前执行 ``Fields.Update()``，让目录域先生成条目。
        verbose: 是否打印过程信息。

    Raises:
        FileNotFoundError: 源文件不存在。
        RuntimeError: 未检测到 Word，或导出未产生文件。
    """
    pythoncom, win32 = _load_word_module()

    src = Path(src).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"找不到源文件: {src}")

    if dst is None:
        dst = src.with_suffix(".pdf")
    dst = Path(dst).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst == src:
        raise ValueError("输出路径不能与源文件相同")

    if verbose:
        print(f"[word2pdf] 源文件: {src}")
        print(f"[word2pdf] 输出:   {dst}")

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        try:
            # DispatchEx: 新建独立 Word 实例，不复用用户已打开的 Word
            word = win32.DispatchEx("Word.Application")
        except Exception as exc:  # pywin32.com_error 未在此导入
            raise RuntimeError(
                f"无法启动 Word（是否已安装 Microsoft Word？）: {exc}"
            ) from exc

        word.Visible = False
        word.DisplayAlerts = WD_ALERTS_NONE

        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        # 只读打开，避免占用文件、也不改动源文档
        doc = word.Documents.Open(str(src), False, True, False)

        if update_fields:
            # 目录（TOC）与页码（PAGE）都是域：不更新则目录只剩占位文字
            try:
                doc.Fields.Update()
            except Exception as exc:
                print(f"[word2pdf] 提示: 更新域失败（继续导出）: {exc}", file=sys.stderr)

        doc.ExportAsFixedFormat(
            OutputFileName=str(dst),
            ExportFormat=WD_EXPORT_FORMAT_PDF,
            OpenAfterExport=False,
            OptimizeFor=WD_EXPORT_OPTIMIZE_PRINT,
            Range=WD_EXPORT_ALL_DOCUMENT,
            Item=WD_EXPORT_DOCUMENT_CONTENT,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=pdfa,
        )
    finally:
        if doc is not None:
            try:
                doc.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        del doc, word
        gc.collect()          # 释放 COM 引用，避免残留 WINWORD.EXE
        pythoncom.CoUninitialize()

    if not dst.is_file():
        raise RuntimeError(f"导出结束但未生成 PDF: {dst}")

    if verbose:
        size_kb = dst.stat().st_size / 1024
        print(f"[word2pdf] 完成: {dst} ({size_kb:.1f} KB)")
    return dst


def main() -> None:
    parser = argparse.ArgumentParser(
        description="用本机 Microsoft Word 将 .docx 渲染为 PDF（Windows / COM）",
    )
    parser.add_argument("path", help="源 .docx 路径")
    parser.add_argument("-o", "--output", default=None, help="输出 .pdf 路径（默认与源文件同目录同名）")
    parser.add_argument("--outdir", default=None, help="输出目录（未给 -o 时生效）")
    parser.add_argument("--pdfa", action="store_true", help="导出 PDF/A-1b（ISO 19005-1）")
    parser.add_argument("--no-update-fields", action="store_true", help="导出前不更新域（保留目录占位文字）")
    args = parser.parse_args()

    if args.output and args.outdir:
        _fail("--output 与 --outdir 不能同时使用")

    src = Path(args.path)
    if not src.is_file():
        _fail(f"{src} 不存在或不是文件")

    if args.output:
        dst = Path(args.output)
    elif args.outdir:
        dst = Path(args.outdir) / src.with_suffix(".pdf").name
    else:
        dst = None

    try:
        convert_docx_to_pdf(
            src,
            dst,
            pdfa=args.pdfa,
            update_fields=not args.no_update_fields,
        )
    except (RuntimeError, FileNotFoundError, ValueError) as exc:
        _fail(str(exc))


if __name__ == "__main__":
    main()
