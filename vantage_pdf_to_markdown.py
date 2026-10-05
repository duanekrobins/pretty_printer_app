#!/usr/bin/env python3
"""
Batch-convert Vantage/Advantage documentation PDFs to detailed Markdown.

This entry point intentionally sends the entire input directory to
OpenDataLoader in one conversion call so one JVM can process the full batch.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

from general_pdf import convert_general_pdf_directory


def _resolve(base: Path, value: str) -> Path:
    path = Path(value).expanduser()
    return path.resolve() if path.is_absolute() else (base / path).resolve()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert all Vantage/Advantage PDFs in an input directory to detailed Markdown."
    )
    parser.add_argument(
        "--config",
        default="pretty_print_config_vantage.json",
        help="Configuration JSON. Default: pretty_print_config_vantage.json",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Delete the output directory before conversion.",
    )
    parser.add_argument(
        "--list-only",
        action="store_true",
        help="List discovered PDFs without converting them.",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    config_path = _resolve(script_dir, args.config)

    if not config_path.exists():
        print(f"ERROR: config file not found: {config_path}")
        return 2

    cfg = json.loads(config_path.read_text(encoding="utf-8"))

    input_values = list(cfg.get("input_dirs", []))
    if not input_values:
        print("ERROR: config must contain at least one input_dirs entry.")
        return 2

    input_dir = _resolve(script_dir, str(input_values[0]))
    output_dir = _resolve(script_dir, str(cfg.get("output_dir", "./outputs")))
    recursive = bool(cfg.get("recursive", True))

    pdf_settings = dict(cfg.get("pdf_settings", {}) or {})
    mode = str(pdf_settings.get("mode", "general_markdown")).strip().lower()
    if mode != "general_markdown":
        print(
            "ERROR: pretty_print_config_vantage.json must use "
            'pdf_settings.mode = "general_markdown".'
        )
        return 2

    settings = dict(pdf_settings.get("general_markdown", {}) or {})

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(
        input_dir.rglob("*.pdf") if recursive else input_dir.glob("*.pdf")
    )

    print("=" * 78)
    print("TestSavvy Pretty Printer - Vantage PDF -> Markdown")
    print("=" * 78)
    print(f"Input : {input_dir}")
    print(f"Output: {output_dir}")
    print(f"PDFs  : {len(pdfs)}")
    print()

    for pdf in pdfs:
        try:
            display = pdf.relative_to(input_dir)
        except ValueError:
            display = pdf
        print(f"  - {display}")

    if args.list_only:
        return 0

    if not pdfs:
        print()
        print("No PDFs found. Place PDF files in the input directory and run again.")
        return 0

    print()
    print("Starting OpenDataLoader conversion...")
    started = time.perf_counter()

    try:
        written, failed, messages = convert_general_pdf_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            settings=settings,
            recursive=recursive,
            clean_output=args.clean,
        )
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        return 2

    for message in messages:
        print(message)

    elapsed = time.perf_counter() - started
    print()
    print("=" * 78)
    print("Conversion summary")
    print("=" * 78)
    print(f"Successful: {written}")
    print(f"Failed    : {failed}")
    print(f"Elapsed   : {elapsed:.2f} seconds")

    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
