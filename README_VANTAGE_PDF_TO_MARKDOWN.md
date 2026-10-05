# Vantage/Advantage PDF to Markdown

## Purpose

This workflow converts normal Vantage/Advantage documentation PDFs into detailed Markdown while preserving table structure, images, and source-page provenance.

It is separate from the TestSavvy execution-result PDF parser.

## Directories

Place source PDFs under:

```text
inputs\
```

Generated content is written under:

```text
outputs\
```

Both paths are relative to the Pretty Printer repository root by default.

## First-time setup

```bat
setup_windows.cmd
```

Requirements:

- Python 3.10 or newer
- Java 11 or newer for OpenDataLoader
- packages from `requirements.txt`

Verify Java:

```bat
java -version
```

## Run

```bat
run_vantage_pdf_to_markdown.cmd
```

PowerShell:

```powershell
.\run_vantage_pdf_to_markdown.ps1
```

List inputs without converting:

```bat
run_vantage_pdf_to_markdown.cmd --list-only
```

Delete the existing output directory before conversion:

```bat
run_vantage_pdf_to_markdown.cmd --clean
```

## Output

For each source PDF, OpenDataLoader generates Markdown and, when present, external PNG image assets.

The output directory also receives:

- `conversion_manifest.json`
- `conversion_manifest.csv`

The manifest records each source PDF, its expected Markdown output, whether Markdown was generated, and Markdown file size.

## Default parsing choices

- Markdown output
- HTML-enabled Markdown for complex tables
- cluster-based table detection
- XY-Cut reading order
- external PNG image output
- source PDF page markers
- headers and footers excluded
- hard line wrapping normalized
- one processing thread
- hybrid processing disabled initially

## Hybrid mode

Hybrid mode can be enabled in `pretty_print_config_vantage.json`:

```json
"hybrid_enabled": true,
"hybrid_backend": "docling-fast",
"hybrid_mode": "auto"
```

Use hybrid processing only when the standard parser has difficulty with scans, unusually complex layouts, or difficult tables.

## Duplicate filenames

Do not place two PDFs with the same base filename anywhere under `inputs\`.

For example, these are rejected before conversion:

```text
inputs\finance\guide.pdf
inputs\admin\guide.pdf
```

because both can produce `guide.md`.

## Standard Pretty Printer integration

General PDFs can also be processed by `pretty_printer.py` when `pdf_settings.mode` is `general_markdown`.

The dedicated Vantage runner is preferred for large document libraries because it sends the entire input directory to OpenDataLoader in one invocation.
