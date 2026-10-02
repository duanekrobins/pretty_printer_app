<img src="assets/GovOps_Division_of_Finance_logo.jpg" alt="State of Utah Division of Finance" width="420">

# TestSavvy Pretty Printer V4 - execution PDF intelligence

Pretty Printer V4 preserves all existing TestSavvy XML, JSON, ZIP, and Excel capabilities and adds a full execution-result PDF pipeline.

## What the PDF engine produces

For individual or combined TestSavvy execution-result PDFs it detects execution boundaries, reconstructs Test Case -> Iteration -> Scenario -> Step order, extracts all report fields, retains `Pass`, `Fail`, and `Not Run`, extracts original embedded screenshots, and writes queryable Markdown, JSON, JSONL, CSV, source PDFs, and visual indexes.

The canonical step corpus includes Dataset Value, English Text, Dataset, Dataset Header, Test Condition, Element Type, Timestamp, Action Code, Message, screenshot references, event tags, and PDF provenance.

## Windows quick start

1. Run `setup_windows.cmd` once.
2. Put TestSavvy PDFs in `input\`.
3. Run `run_pretty_printer.cmd`.
4. Review `output\testsavvy_execution_reports\`.
5. Run `validate_pdf.cmd "C:\path\to\combined_results.pdf"` when you want a corpus-wide parser validation check.

## Combined PDF filtering

Use `pdf_settings` in `pretty_print_config.json` to process only selected runs while testing:

```json
"pdf_settings": {
    "run_ids": ["20426"],
    "test_case_ids": [],
    "statuses": [],
    "max_runs": 0
}
```

Set the lists back to empty to process every execution.

## Screenshot OCR

Screenshots are always preserved when `extract_images=true`. OCR is optional supplemental evidence and never replaces the original image.

Supported `ocr_mode` values:

- `off`
- `failures`
- `failures_and_key_events`
- `all`

For multimodal AI questions, keep the per-run `source.pdf` and extracted screenshots with the structured corpus.

## Validation status

The parser was validated against the supplied 8,878-page combined TestSavvy report. All 137 detected executions matched their report-level Pass/Fail/Not Run totals in the current validation run.

See `NOTEBOOKLM_TESTSAVVY_EXECUTION_PDF_GUIDE.md` for AI query guidance.
