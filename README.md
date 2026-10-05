<img src="assets/GovOps_Division_of_Finance_logo.jpg" alt="State of Utah Division of Finance" width="420">

# TestSavvy Pretty Printer V4

TestSavvy Pretty Printer V4 converts TestSavvy execution evidence and general Vantage/Advantage documentation into AI-ready knowledge outputs.

Version 4.1 adds a second PDF path while preserving the specialized execution-result parser.

## Supported inputs

- TestSavvy execution-result PDFs
- General Vantage/Advantage PDFs
- XML
- JSON
- ZIP
- XLSX/XLSM

## PDF processing modes

Pretty Printer now has two deliberately separate PDF engines.

### 1. TestSavvy execution-result intelligence

Use:

```json
"pdf_settings": {
    "mode": "testsavvy_execution"
}
```

This uses `testsavvy_pdf.py` and reconstructs execution boundaries, Test Case -> Iteration -> Scenario -> Step order, Pass/Fail/Not Run results, runtime metadata, screenshots, source PDFs, Markdown, JSON, JSONL, CSV, and visual/event indexes.

### 2. General Vantage/Advantage PDF -> Markdown

Use:

```json
"pdf_settings": {
    "mode": "general_markdown"
}
```

This uses `general_pdf.py` with OpenDataLoader PDF. It is intended for administration guides, financial run sheets, user guides, configuration guides, and other normal Vantage/Advantage documentation.

The default general-document settings use:

- detailed Markdown output
- HTML inside Markdown for complex tables
- cluster table detection
- XY-Cut reading order
- external PNG image extraction
- source-PDF page separators
- normalized line wrapping
- header/footer suppression
- deterministic single-thread processing
- optional hybrid mode

## Windows setup

Run once:

```bat
setup_windows.cmd
```

The setup script creates `.venv`, installs `requirements.txt`, and checks whether Java is available.

TestSavvy execution-result parsing does not require Java. General PDF -> Markdown conversion through OpenDataLoader requires Java 11 or newer.

## TestSavvy execution-result quick start

1. Put TestSavvy execution-result PDFs in `input\`.
2. Confirm `pretty_print_config.json` uses `pdf_settings.mode = "testsavvy_execution"`.
3. Run:

```bat
run_pretty_printer.cmd
```

4. Review `output\testsavvy_execution_reports\`.
5. For a baseline validation run:

```bat
validate_pdf.cmd "C:\path\to\combined_results.pdf"
```

## Vantage/Advantage documentation quick start

A dedicated batch runner is included because OpenDataLoader can process an entire directory with one JVM rather than starting Java once per PDF.

1. Put all Vantage/Advantage PDFs under:

```text
inputs\
```

2. Run:

```bat
run_vantage_pdf_to_markdown.cmd
```

3. Markdown and extracted image assets are written under:

```text
outputs\
```

Useful commands:

```bat
run_vantage_pdf_to_markdown.cmd --list-only
run_vantage_pdf_to_markdown.cmd --clean
```

The Vantage runner uses `pretty_print_config_vantage.json`.

## General PDF configuration

The main general-document options are under:

```json
"pdf_settings": {
    "mode": "general_markdown",
    "general_output_subdir": "",
    "general_markdown": {
        "markdown_with_html": true,
        "image_output": "external",
        "image_format": "png",
        "table_method": "cluster",
        "reading_order": "xycut",
        "markdown_page_separator": "\n\n---\n\n**Source PDF page %page-number%**\n\n",
        "include_header_footer": false,
        "keep_line_breaks": false,
        "threads": "1",
        "quiet": false,
        "hybrid_enabled": false,
        "hybrid_backend": "docling-fast",
        "hybrid_mode": "auto"
    }
}
```

For the dedicated Vantage runner, these values are already supplied in `pretty_print_config_vantage.json`.

## Why the PDF engines remain separate

A TestSavvy execution-result PDF is structured runtime evidence and needs a purpose-built parser. A normal Vantage guide is a document-layout problem and benefits from OpenDataLoader's reading-order, table, and image extraction.

Keeping both paths separate prevents general-document conversion from weakening the validated execution parser.

## Output validation

The Vantage batch runner:

- discovers every PDF recursively
- rejects duplicate PDF base names before conversion
- converts the directory in one OpenDataLoader invocation
- confirms that each input PDF produced a non-empty Markdown file
- writes `conversion_manifest.json`
- writes `conversion_manifest.csv`

## Existing TestSavvy PDF validation

The execution-result parser was validated against the supplied 8,878-page combined TestSavvy report. The current validation artifact contains 137 detected executions and zero validation failures.

See:

- `NOTEBOOKLM_TESTSAVVY_EXECUTION_PDF_GUIDE.md`
- `NOTEBOOKLM_TESTSAVVY_XML_JSON_ORDERING_GUIDE.md`
