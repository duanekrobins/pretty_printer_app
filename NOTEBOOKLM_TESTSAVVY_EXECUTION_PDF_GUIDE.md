<img src="assets/GovOps_Division_of_Finance_logo.jpg" alt="State of Utah Division of Finance" width="420">

# TestSavvy execution-result PDF knowledge guide

## Purpose

Use this guide with the execution Markdown, per-run PDFs, CSV/JSONL indexes, and extracted screenshots produced by Pretty Printer. The execution-result PDF is the primary runtime evidence source for questions about what actually happened during a TestSavvy run.

## Evidence hierarchy

For questions about runtime behavior, use evidence in this order:

1. The structured execution step extracted from the execution-result PDF.
2. The step-linked screenshot when the question concerns what appeared on screen.
3. The report `Message` field for execution output or assertions.
4. The report `Action Code` field for the Selenium/WebDriver command that executed or would have executed.
5. The per-run `source.pdf` and source page for verification.

Do not assume a blank `Message` means nothing was visible on screen. The screenshot may contain a validation ribbon, modal, error, warning, transaction status, or other application response that is not represented in the Message field.

## Arbitrary-question rule

The corpus is not limited to predefined questions. Use `execution_steps.csv` or `execution_steps.jsonl` to locate candidate runs/steps, then use the linked screenshot and per-run PDF when visual evidence is needed.

Examples include:

- For failed test cases, what message was visible after `Click: Submit`?
- What actual values were entered into Department across all runs?
- Which Selenium commands targeted `viewActions.submit`?
- What happened immediately before each failure?
- Which steps were `Not Run` after a failure?
- Which screenshots show a particular Advantage error code?
- Which runs used a particular XPath, `data-qa`, `data-qa-id`, or `aria-label` selector?

## Understanding "after" an event

Do not interpret "after Click: Submit" as PDF page +1. Follow `test_case_step_order` and the event indexes. The screenshot on the Submit step is often the post-action screen state; `event_index.csv` also records the next screenshot-bearing step so the AI can inspect both interpretations when necessary.

## Important files

| File | Purpose |
| --- | --- |
| `INDEX.md` | Run-level navigation. |
| `execution_summaries.csv` | One row per execution. |
| `execution_steps.csv` / `.jsonl` | Canonical searchable runtime step corpus. |
| `screenshot_manifest.csv` / `.jsonl` | Screenshot-to-step relationships and optional OCR text. |
| `event_index.csv` / `.jsonl` | Event tags plus event/post-event visual references. |
| `failure_and_not_run_index.csv` | Failed and skipped steps. |
| `TC_<ID>_RUN_<RUN>/execution.md` | Human/AI readable run reconstruction. |
| `TC_<ID>_RUN_<RUN>/execution.json` | Complete structured run object. |
| `TC_<ID>_RUN_<RUN>/source.pdf` | Original visual report pages for that run. |
| `TC_<ID>_RUN_<RUN>/screenshots/` | Original embedded screenshots extracted from the report. |

## Text representations

Code-like and expression fields preserve three forms where applicable:

- `as_displayed`: PDF line wrapping preserved.
- `normalized`: visual whitespace normalized.
- `reconstructed`: conservative repair of obvious PDF token splitting.

When precision matters, use `reconstructed` for searching and `as_displayed` plus `source.pdf` for verification.
