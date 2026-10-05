#!/usr/bin/env python3
"""
General PDF -> Markdown support for TestSavvy Pretty Printer.

This module deliberately keeps general-document PDF conversion separate from
the specialized TestSavvy execution-result parser in testsavvy_pdf.py.

The implementation uses OpenDataLoader PDF for layout-aware Markdown,
complex-table handling, page provenance, and extracted images.
"""

from __future__ import annotations

import csv
import json
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


DEFAULT_GENERAL_PDF_SETTINGS: Dict[str, Any] = {
    "markdown_with_html": True,
    "image_output": "external",
    "image_format": "png",
    "table_method": "cluster",
    "reading_order": "xycut",
    "markdown_page_separator": "\n\n---\n\n**Source PDF page %page-number%**\n\n",
    "include_header_footer": False,
    "keep_line_breaks": False,
    "threads": "1",
    "quiet": False,
    "hybrid_enabled": False,
    "hybrid_backend": "docling-fast",
    "hybrid_mode": "auto",
}


def _settings_from_config(cfg: Any) -> Dict[str, Any]:
    pdf_settings = dict(getattr(cfg, "pdf_settings", {}) or {})
    user = dict(pdf_settings.get("general_markdown", {}) or {})
    merged = dict(DEFAULT_GENERAL_PDF_SETTINGS)
    merged.update(user)
    return merged


def _load_opendataloader():
    try:
        import opendataloader_pdf
    except ImportError as exc:
        raise RuntimeError(
            "General PDF Markdown conversion requires the 'opendataloader-pdf' "
            "Python package. Run setup_windows.cmd again or install it with "
            "'python -m pip install -U opendataloader-pdf'."
        ) from exc
    return opendataloader_pdf


def _convert_kwargs(settings: Dict[str, Any]) -> Dict[str, Any]:
    page_separator = str(
        settings.get(
            "markdown_page_separator",
            "\n\n---\n\n**Source PDF page %page-number%**\n\n",
        )
    )
    # Accept either normal newline escapes parsed by JSON or literal "\\n"
    # sequences supplied by hand-edited configuration files.
    page_separator = page_separator.replace("\\n", "\n")

    kwargs: Dict[str, Any] = {
        "format": "markdown",
        "markdown_with_html": bool(settings.get("markdown_with_html", True)),
        "image_output": str(settings.get("image_output", "external")),
        "image_format": str(settings.get("image_format", "png")),
        "table_method": str(settings.get("table_method", "cluster")),
        "reading_order": str(settings.get("reading_order", "xycut")),
        "markdown_page_separator": page_separator,
        "include_header_footer": bool(settings.get("include_header_footer", False)),
        "keep_line_breaks": bool(settings.get("keep_line_breaks", False)),
        "threads": str(settings.get("threads", "1")),
        "quiet": bool(settings.get("quiet", False)),
    }

    if bool(settings.get("hybrid_enabled", False)):
        kwargs["hybrid"] = str(settings.get("hybrid_backend", "docling-fast"))
        kwargs["hybrid_mode"] = str(settings.get("hybrid_mode", "auto"))

    return kwargs


def _expected_markdown(output_dir: Path, src: Path) -> Path:
    return output_dir / f"{src.stem}.md"


def _find_markdown_for_pdf(output_dir: Path, src: Path) -> Optional[Path]:
    expected = _expected_markdown(output_dir, src)
    if expected.exists() and expected.stat().st_size > 0:
        return expected

    candidates = [
        p for p in output_dir.rglob("*.md")
        if p.stem.lower() == src.stem.lower() and p.stat().st_size > 0
    ]
    return sorted(candidates)[0] if candidates else None


def process_general_pdf(
    src: Path,
    cfg: Any,
    input_root: Optional[Path] = None,
) -> Tuple[int, int, List[str]]:
    """
    Convert one ordinary PDF to detailed Markdown.

    This is used by the main Pretty Printer dispatcher when
    pdf_settings.mode == "general_markdown".
    """
    messages: List[str] = []
    settings = _settings_from_config(cfg)

    output_root = Path(getattr(cfg, "output_dir")).resolve()
    output_subdir = str(
        (getattr(cfg, "pdf_settings", {}) or {}).get(
            "general_output_subdir", "general_pdf_markdown"
        )
    ).strip()

    target_root = output_root / output_subdir if output_subdir else output_root

    if input_root is not None:
        try:
            rel_parent = src.resolve().relative_to(input_root.resolve()).parent
            target_root = target_root / rel_parent
        except ValueError:
            pass

    target_root.mkdir(parents=True, exist_ok=True)
    expected = _expected_markdown(target_root, src)

    overwrite = bool(getattr(cfg, "overwrite", False))
    if expected.exists() and not overwrite:
        messages.append(f"SKIP (exists): {src} -> {expected}")
        return 0, 0, messages

    opendataloader_pdf = _load_opendataloader()

    try:
        opendataloader_pdf.convert(
            input_path=str(src),
            output_dir=str(target_root),
            **_convert_kwargs(settings),
        )
    except Exception as exc:
        messages.append(
            f"FAIL general PDF: {src} ({type(exc).__name__}: {exc})"
        )
        return 0, 1, messages

    md_path = _find_markdown_for_pdf(target_root, src)
    if md_path is None:
        messages.append(
            f"FAIL general PDF: {src} (conversion completed but no non-empty Markdown file was found)"
        )
        return 0, 1, messages

    messages.append(f"OK general PDF: {src} -> {md_path}")
    return 1, 0, messages


def _discover_pdfs(input_dir: Path, recursive: bool) -> List[Path]:
    walker = input_dir.rglob("*.pdf") if recursive else input_dir.glob("*.pdf")
    return sorted(p.resolve() for p in walker if p.is_file())


def _duplicate_stems(pdfs: List[Path]) -> Dict[str, List[Path]]:
    grouped: Dict[str, List[Path]] = {}
    for pdf in pdfs:
        grouped.setdefault(pdf.stem.lower(), []).append(pdf)
    return {stem: paths for stem, paths in grouped.items() if len(paths) > 1}


def _write_manifest(
    input_dir: Path,
    output_dir: Path,
    pdfs: List[Path],
) -> Tuple[Path, Path]:
    rows: List[Dict[str, Any]] = []

    for pdf in pdfs:
        md = _find_markdown_for_pdf(output_dir, pdf)
        rows.append(
            {
                "source_pdf": str(pdf),
                "source_relative_path": str(pdf.relative_to(input_dir)),
                "markdown_file": str(md) if md else "",
                "markdown_exists": bool(md),
                "markdown_bytes": md.stat().st_size if md else 0,
            }
        )

    json_path = output_dir / "conversion_manifest.json"
    csv_path = output_dir / "conversion_manifest.csv"

    json_path.write_text(
        json.dumps(rows, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "source_pdf",
                "source_relative_path",
                "markdown_file",
                "markdown_exists",
                "markdown_bytes",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)

    return json_path, csv_path


def convert_general_pdf_directory(
    input_dir: Path,
    output_dir: Path,
    settings: Optional[Dict[str, Any]] = None,
    *,
    recursive: bool = True,
    clean_output: bool = False,
) -> Tuple[int, int, List[str]]:
    """
    Convert every PDF in a directory using a single OpenDataLoader invocation.

    This is the preferred path for large Vantage/Advantage documentation
    libraries because it avoids starting a new JVM for every PDF.
    """
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    merged_settings = dict(DEFAULT_GENERAL_PDF_SETTINGS)
    merged_settings.update(settings or {})

    messages: List[str] = []

    if not input_dir.exists():
        return 0, 1, [f"FAIL: input directory does not exist: {input_dir}"]

    pdfs = _discover_pdfs(input_dir, recursive)
    if not pdfs:
        return 0, 0, [f"No PDF files found under: {input_dir}"]

    duplicates = _duplicate_stems(pdfs)
    if duplicates:
        messages.append(
            "FAIL: duplicate PDF base names were found. OpenDataLoader can generate "
            "the same Markdown filename for documents with the same base name."
        )
        for stem, paths in sorted(duplicates.items()):
            messages.append(f"  Duplicate stem '{stem}':")
            for path in paths:
                messages.append(f"    {path}")
        return 0, len(duplicates), messages

    if clean_output and output_dir.exists():
        shutil.rmtree(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    opendataloader_pdf = _load_opendataloader()

    messages.append(f"PDF files discovered: {len(pdfs)}")
    messages.append(f"Input directory: {input_dir}")
    messages.append(f"Output directory: {output_dir}")

    try:
        opendataloader_pdf.convert(
            input_path=str(input_dir),
            output_dir=str(output_dir),
            **_convert_kwargs(merged_settings),
        )
    except Exception as exc:
        messages.append(
            f"FAIL batch PDF conversion ({type(exc).__name__}: {exc})"
        )
        return 0, len(pdfs), messages

    succeeded = 0
    missing: List[Path] = []
    for pdf in pdfs:
        if _find_markdown_for_pdf(output_dir, pdf):
            succeeded += 1
        else:
            missing.append(pdf)

    json_path, csv_path = _write_manifest(input_dir, output_dir, pdfs)
    messages.append(f"Manifest JSON: {json_path}")
    messages.append(f"Manifest CSV: {csv_path}")

    if missing:
        messages.append(
            f"WARNING: {len(missing)} PDF(s) did not produce a non-empty Markdown file:"
        )
        messages.extend(f"  {path}" for path in missing)

    failed = len(missing)
    messages.append(
        f"General PDF batch complete: {succeeded} succeeded, {failed} failed."
    )
    return succeeded, failed, messages
