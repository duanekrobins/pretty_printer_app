#!/usr/bin/env python3
"""Validate TestSavvy execution-result PDF parsing against report totals."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import testsavvy_pdf as tspdf  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", type=Path)
    ap.add_argument("--json", type=Path, default=None, help="Optional validation JSON output")
    args = ap.parse_args()

    doc = fitz.open(args.pdf)
    ranges = tspdf.find_execution_ranges(doc)
    rows = []
    failures = []
    for start, end in ranges:
        execution = tspdf.parse_execution(doc, start, end)
        row = {
            "test_case_id": execution.test_case_id,
            "test_case_name": execution.test_case_name,
            "run_id": execution.run_id,
            "run_status": execution.run_status,
            "pages": f"{start+1}-{end+1}",
            **execution.validation,
        }
        rows.append(row)
        if execution.validation.get("status") != "PASS":
            failures.append(row)

    result = {
        "pdf": str(args.pdf),
        "page_count": doc.page_count,
        "execution_count": len(ranges),
        "validation_failures": failures,
        "all_passed": not failures,
        "executions": rows,
    }
    doc.close()

    print(f"PDF pages: {result['page_count']}")
    print(f"Executions: {result['execution_count']}")
    print(f"Validation failures: {len(failures)}")
    for row in failures:
        print(
            f"REVIEW TC={row['test_case_id']} RUN={row['run_id']} STATUS={row['run_status']} "
            f"parsed={row['parsed_pass']}/{row['parsed_fail']}/{row['parsed_not_run']} "
            f"reported={row['reported_pass']}/{row['reported_fail']}/{row['reported_not_run']}"
        )

    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Wrote: {args.json}")
    return 0 if not failures else 2


if __name__ == "__main__":
    raise SystemExit(main())
