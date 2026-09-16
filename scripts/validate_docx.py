# -*- coding: utf-8 -*-
"""
validate_docx.py
Lightweight .docx validator using only Python standard library.

No external dependencies (no lxml, no defusedxml).
Performs 5 structural checks that catch the most common corruption issues:

1. ZIP integrity — all entries can be decompressed
2. XML well-formedness — all .xml/.rels files parse correctly
3. File reference integrity — every Target in .rels exists in the package
4. Content-Type declaration — every XML part has an Override in [Content_Types].xml
5. Whitespace preservation — w:t elements with edge whitespace have xml:space="preserve"

Usage:
    python validate_docx.py output.docx
    python validate_docx.py output.docx --verbose

Exit codes:
    0 — all checks passed
    1 — one or more checks failed
"""

import argparse
import os
import stat
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# OOXML namespaces
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

WHITESPACE_CHARS = " \t\n\r"


def validate_zip(zf: zipfile.ZipFile, verbose: bool = False) -> bool:
    """Check 1: ZIP integrity."""
    bad = zf.testzip()
    if bad is not None:
        print(f"FAILED - ZIP corrupt at entry: {bad}")
        return False
    if verbose:
        print(f"PASSED - ZIP integrity ({len(zf.namelist())} entries)")
    return True


def validate_xml(zf: zipfile.ZipFile, verbose: bool = False) -> bool:
    """Check 2: All XML files are well-formed."""
    errors = []
    for name in zf.namelist():
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            data = zf.read(name)
            ET.fromstring(data)
        except ET.ParseError as e:
            errors.append(f"  {name}: {e}")

    if errors:
        print(f"FAILED - {len(errors)} XML parse error(s):")
        for err in errors:
            print(err)
        return False
    if verbose:
        xml_count = sum(1 for n in zf.namelist() if n.endswith((".xml", ".rels")))
        print(f"PASSED - All {xml_count} XML files well-formed")
    return True


def validate_references(zf: zipfile.ZipFile, verbose: bool = False) -> bool:
    """Check 3: Every relationship Target file exists in the package."""
    errors = []
    all_names = set(zf.namelist())
    rels_files = [n for n in all_names if n.endswith(".rels")]

    for rels_path in rels_files:
        try:
            root = ET.fromstring(zf.read(rels_path))
        except ET.ParseError:
            continue  # Already caught in validate_xml

        # Determine base directory for relative targets
        rels_dir = os.path.dirname(rels_path)
        # rels files are in _rels/ subdirs; parent is the actual part dir
        if rels_dir.endswith("_rels"):
            base_dir = os.path.dirname(rels_dir)
        else:
            base_dir = rels_dir

        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
            target = rel.get("Target", "")
            target_mode = rel.get("TargetMode", "")

            # Skip external targets
            if target_mode == "External":
                continue
            # Skip URL targets
            if target.startswith(("http://", "https://", "mailto:")):
                continue

            # Resolve target path
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = os.path.normpath(os.path.join(base_dir, target)).replace("\\", "/")

            if resolved not in all_names:
                errors.append(f"  {rels_path}: Target '{target}' -> '{resolved}' not found in package")

    if errors:
        print(f"FAILED - {len(errors)} broken reference(s):")
        for err in errors:
            print(err)
        return False
    if verbose:
        print(f"PASSED - All relationship targets exist ({len(rels_files)} .rels files checked)")
    return True


def validate_content_types(zf: zipfile.ZipFile, verbose: bool = False) -> bool:
    """Check 4: Every XML part in word/ has a content type declaration."""
    errors = []

    ct_path = "[Content_Types].xml"
    if ct_path not in zf.namelist():
        print("FAILED - [Content_Types].xml not found")
        return False

    try:
        ct_root = ET.fromstring(zf.read(ct_path))
    except ET.ParseError as e:
        print(f"FAILED - Cannot parse [Content_Types].xml: {e}")
        return False

    # Collect declared parts
    declared_parts = set()
    for override in ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        part_name = override.get("PartName", "")
        if part_name:
            declared_parts.add(part_name.lstrip("/"))

    # Collect declared extensions
    declared_exts = set()
    for default in ct_root.findall(f"{{{CONTENT_TYPES_NS}}}Default"):
        ext = default.get("Extension", "")
        if ext:
            declared_exts.add(ext.lower())

    # Check XML files in word/ that should have Override
    for name in zf.namelist():
        if name == ct_path or name.endswith(".rels") or "_rels/" in name:
            continue
        if not name.endswith(".xml"):
            continue
        if name.startswith("docProps/"):
            continue

        if name not in declared_parts:
            # Check if it's covered by a Default extension
            ext = os.path.splitext(name)[1].lstrip(".").lower()
            if ext not in declared_exts:
                errors.append(f"  {name}: not declared in [Content_Types].xml")

    if errors:
        print(f"FAILED - {len(errors)} content type declaration error(s):")
        for err in errors:
            print(err)
        return False
    if verbose:
        print(f"PASSED - All parts declared in [Content_Types].xml")
    return True


def validate_whitespace(zf: zipfile.ZipFile, verbose: bool = False) -> bool:
    """Check 5: w:t elements with edge whitespace must have xml:space='preserve'."""
    errors = []
    doc_path = "word/document.xml"

    if doc_path not in zf.namelist():
        if verbose:
            print("PASSED - No word/document.xml found (skipping whitespace check)")
        return True

    try:
        root = ET.fromstring(zf.read(doc_path))
    except ET.ParseError:
        return True  # Already caught in validate_xml

    count = 0
    for elem in root.iter(f"{{{WORD_NS}}}t"):
        text = elem.text
        if not text:
            continue
        if text[0] in WHITESPACE_CHARS or text[-1] in WHITESPACE_CHARS:
            space_attr = elem.get(XML_SPACE_ATTR)
            if space_attr != "preserve":
                preview = repr(text[:50])
                if len(text) > 50:
                    preview += "..."
                errors.append(f"  w:t with edge whitespace missing xml:space='preserve': {preview}")
                count += 1

    if errors:
        print(f"FAILED - {len(errors)} whitespace preservation violation(s):")
        for err in errors[:10]:  # Limit output
            print(err)
        if count > 10:
            print(f"  ... and {count - 10} more")
        return False
    if verbose:
        print(f"PASSED - Whitespace preservation OK")
    return True


def count_paragraphs(zf: zipfile.ZipFile) -> int:
    """Count w:p elements in document.xml for sanity reporting."""
    doc_path = "word/document.xml"
    if doc_path not in zf.namelist():
        return 0
    try:
        root = ET.fromstring(zf.read(doc_path))
        return len(list(root.iter(f"{{{WORD_NS}}}p")))
    except ET.ParseError:
        return 0


def validate_docx(path: str, verbose: bool = False) -> bool:
    """Run all validation checks on a .docx file.

    Args:
        path: Path to the .docx file.
        verbose: Print PASSED messages for each check.

    Returns:
        True if all checks passed, False otherwise.
    """
    file_path = Path(path)
    if not file_path.exists():
        print(f"Error: File not found: {path}")
        return False

    if file_path.suffix.lower() != ".docx":
        print(f"Error: Not a .docx file: {path}")
        return False

    file_size = file_path.stat().st_size
    if file_size < 3000:
        print(f"Warning: File is only {file_size} bytes (empty docx is ~3KB)")

    print(f"Validating: {file_path.name} ({file_size:,} bytes)")
    print()

    all_passed = True

    with zipfile.ZipFile(file_path, "r") as zf:
        all_passed &= validate_zip(zf, verbose)
        all_passed &= validate_xml(zf, verbose)
        all_passed &= validate_references(zf, verbose)
        all_passed &= validate_content_types(zf, verbose)
        all_passed &= validate_whitespace(zf, verbose)

        para_count = count_paragraphs(zf)
        print(f"\nParagraphs: {para_count}")

    print()
    if all_passed:
        print("All validations PASSED!")
    else:
        print("Some validations FAILED — see above for details.")

    return all_passed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Lightweight .docx validator (Python stdlib only, no external dependencies)"
    )
    parser.add_argument("path", help="Path to the .docx file to validate")
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Print PASSED messages for each check",
    )
    args = parser.parse_args()

    success = validate_docx(args.path, verbose=args.verbose)
    sys.exit(0 if success else 1)
