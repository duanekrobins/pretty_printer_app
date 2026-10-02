# Pretty Printer App — Latest Complete Package

This package contains the latest TestSavvy XML, JSON, ZIP, and Excel pretty-printer source and its matching settings file.

## Included files

| File | Purpose |
| --- | --- |
| `pretty_printer.py` | Current master application source with TestSavvy JSON execution ordering fixes. |
| `pretty_print_config.json` | Complete 29-setting configuration used by the application. |
| `requirements.txt` | Python dependency declaration. |
| `NOTEBOOKLM_TESTSAVVY_XML_JSON_ORDERING_GUIDE.md` | Guide for interpreting generated TestSavvy Markdown in NotebookLM or another AI knowledge system. |
| `input/` | Place source XML, JSON, ZIP, XLSX, and XLSM files here. |
| `output/` | Generated Markdown is written here by the default configuration. |

## Requirements

- Python 3.9 or newer
- `openpyxl` 3.1 or newer

## Install

From this folder, run:

```bash
python -m pip install -r requirements.txt
```

On Windows, `py` can be used instead of `python`:

```powershell
py -m pip install -r requirements.txt
```

## Run

1. Copy source files into `input/`.
2. Open a terminal in this package folder.
3. Run:

```bash
python pretty_printer.py --config pretty_print_config.json
```

Windows alternative:

```powershell
py pretty_printer.py --config pretty_print_config.json
```

Generated files are written to `output/`.

## Default behavior

The supplied configuration:

- processes XML, JSON, ZIP, XLSX, and XLSM files recursively;
- generates one Markdown output per source file or worksheet;
- applies TestSavvy-aware XML and JSON formatting;
- sorts TestSavvy JSON scenarios by execution sequence instead of raw array position;
- separates active/executable content from excluded or inactive content;
- generates Test Case, Submit-button, custom-function, selector/action-code, Interface Repository, and JSON inventory sections where applicable;
- normalizes generated Markdown filenames;
- overwrites previously generated output; and
- writes the NotebookLM TestSavvy ordering guide into the output directory.

## Settings

Edit `pretty_print_config.json` to change input/output directories, file types, recursive processing, overwrite behavior, output layout, inventory limits, or TestSavvy formatting modes. Paths in the supplied configuration are relative to the directory from which the command is run.

## Current source lineage

This package is based on `pretty_printer_testsavvy_json_order_fixed.py`, the newest verified source. It supersedes the current GitHub `main` version and both earlier backup ZIP versions reviewed on August 4, 2026.
