#!/usr/bin/env python3
"""
TestSavvy execution-result PDF intelligence engine.

This module extends Pretty Printer with a loss-preserving parser for TestSavvy
execution result PDFs. It supports both individual execution reports and large
combined PDFs containing many execution reports.

Design goals
------------
1. Treat the PDF as evidence, not merely text to summarize.
2. Reconstruct the report hierarchy:
       Execution -> Iteration -> Scenario -> Step
3. Preserve every step field:
       Result, Action Name, Logical, Physical, Dataset Value, English Text,
       Dataset, Dataset Header, Test Condition, Element Type, Timestamp,
       Action Code, Message.
4. Preserve and link embedded screenshots to the exact step that produced them.
5. Preserve Pass, Fail, and Not Run steps.
6. Keep source provenance (PDF page and geometry) for every extracted record.
7. Emit Markdown, JSON/JSONL, CSV, image manifests, failure indexes, and the
   per-run source PDF so an AI system can answer arbitrary questions while
   retaining access to the original visual evidence.

The parser is intentionally stateful and sequence-oriented because TestSavvy
records can cross PDF page boundaries. A step can begin on one page while its
Dataset metadata / Action Code / Message continues on the next page.
"""

from __future__ import annotations

import csv
import hashlib
import json
import re
import shutil
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF

try:
    import pytesseract
    from PIL import Image
except Exception:  # OCR is optional at runtime
    pytesseract = None
    Image = None


# ---------------------------------------------------------------------------
# Constants based on the stable TestSavvy landscape report geometry.
# They are used as column anchors, not as brittle full-page templates.
# ---------------------------------------------------------------------------

STEP_COLUMNS: Tuple[Tuple[str, float, float], ...] = (
    ("result", 150.0, 230.0),
    ("action_name", 230.0, 286.5),
    ("logical", 286.5, 354.5),
    ("physical", 354.5, 434.5),
    ("dataset_value", 434.5, 506.5),
    ("english_text", 506.5, 579.5),
)

META_COLUMNS: Tuple[Tuple[str, float, float], ...] = (
    ("dataset", 286.5, 354.5),
    ("dataset_header", 354.5, 434.5),
    ("test_condition", 434.5, 578.5),
    ("element_type", 578.5, 649.5),
    ("timestamp", 649.5, 760.0),
)

KNOWN_ACTION_NAMES = {
    "Click",
    "Enter Text",
    "Wait",
    "Verify Exists",
    "Smart Wait",
    "Open Browser",
    "END IF",
    "IF",
    "Close Browser",
    "Custom Function",
    "Verify Text",
    "Select",
    "Verify",
    "Enter",
    "Set Runtime Value",
    "While",
    "Next",
    "Explicit Wait",
    "Dropdown Options",
}

RESULT_VALUES = {"Pass", "Fail", "Not Run"}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class TextEvidence:
    as_displayed: str = ""
    normalized: str = ""
    reconstructed: str = ""


@dataclass
class ScreenshotEvidence:
    present: bool = False
    file: str = ""
    width: int = 0
    height: int = 0
    source_page: int = 0  # 1-based page number in source PDF
    run_page: int = 0     # 1-based page number inside this run
    xref: int = 0
    bbox: List[float] = field(default_factory=list)
    sha256: str = ""
    ocr_text: str = ""
    ocr_status: str = "not_requested"


@dataclass
class StepRecord:
    global_step_order: int
    test_case_step_order: int
    scenario_step_order: int
    scenario_order: int
    scenario: str
    result: str
    action_name: str
    logical: str
    physical: str
    dataset_value: TextEvidence
    english_text: TextEvidence
    dataset: str = ""
    dataset_header: str = ""
    test_condition: str = ""
    element_type: str = ""
    timestamp: str = ""
    action_code: TextEvidence = field(default_factory=TextEvidence)
    message: TextEvidence = field(default_factory=TextEvidence)
    screenshot: ScreenshotEvidence = field(default_factory=ScreenshotEvidence)
    action_page: int = 0
    metadata_page: int = 0
    action_y: float = 0.0
    metadata_y: float = 0.0
    previous_step_order: Optional[int] = None
    next_step_order: Optional[int] = None
    event_tags: List[str] = field(default_factory=list)


@dataclass
class ScenarioRecord:
    scenario_order: int
    scenario: str
    steps_passed: Optional[int]
    steps_failed: Optional[int]
    steps_not_run: Optional[int]
    post_verify: Optional[int]
    source_page: int
    y: float


@dataclass
class ExecutionRecord:
    source_pdf: str
    source_start_page: int
    source_end_page: int
    test_case_id: str
    test_case_name: str
    run_status: str
    run_id: str
    machine_name: str
    execution_date: str
    reason: str
    received_date: str
    override_reason: str
    executed_by: str
    reported_problems: str
    iteration: Optional[int]
    reported_steps_passed: Optional[int]
    reported_steps_failed: Optional[int]
    reported_steps_not_run: Optional[int]
    reported_post_verify: Optional[int]
    raw_metadata_record_count: int = 0
    scenarios: List[ScenarioRecord] = field(default_factory=list)
    steps: List[StepRecord] = field(default_factory=list)
    validation: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Text / geometry helpers
# ---------------------------------------------------------------------------

def _clean_ws(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _join_display_words(words: Sequence[Dict[str, Any]]) -> str:
    """Join words preserving PDF line boundaries."""
    if not words:
        return ""
    lines: List[List[Dict[str, Any]]] = []
    for word in sorted(words, key=lambda w: (w["y0"], w["x0"])):
        target = None
        for line in lines:
            if abs(line[0]["y0"] - word["y0"]) <= 2.5:
                target = line
                break
        if target is None:
            target = []
            lines.append(target)
        target.append(word)
    return "\n".join(
        " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))
        for line in lines
    ).strip()


def _normalized_from_display(value: str) -> str:
    return _clean_ws(value)


def reconstruct_testsavvy_text(value: str, code_like: bool = False) -> str:
    """
    Conservatively reconstruct obvious PDF line-wrap artifacts while retaining
    the original displayed representation elsewhere.
    """
    text = _clean_ws(value)
    if not text:
        return ""

    # Known artificial splits seen in narrow report columns.
    replacements = {
        "@advantagefxf x.": "@advantagefxfx.",
        "@advantagefx. ": "@advantagefx.",
        "@cgifx. ": "@cgifx.",
        "@storeResult (": "@storeResult(",
        "${ ": "${",
        " }": "}",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)

    # Join underscore-separated identifiers split by PDF wrapping.
    text = re.sub(r"(?<=\w)\s+_(?=\w)", "_", text)
    text = re.sub(r"(?<=_)\s+(?=\w)", "", text)
    text = re.sub(r"(?<=\w)_\s+(?=\w)", "_", text)

    # Remove visual spaces inside common TestSavvy expression delimiters.
    text = re.sub(r"@storeResult\s*\(\s*", "@storeResult(", text)
    text = re.sub(r"\s*\)\s*\|~\|\s*", ")|~|", text)
    text = re.sub(r"\|~\|\s*@", "|~|@", text)

    if code_like:
        # Avoid changing quoted English strings. Only repair a few syntactic
        # spaces that PDF word extraction commonly inserts around punctuation.
        text = text.replace("data- qa-id", "data-qa-id").replace("data-qa- id", "data-qa-id").replace("data- qa", "data-qa")
        text = text.replace("aria- label", "aria-label")
        text = re.sub(r"\bwd\.\s+", "wd.", text)
        text = re.sub(r"webdriverfx\.\s+", "webdriverfx.", text)
        text = re.sub(r"\bfind_element_by_\s+", "find_element_by_", text)
    return text.strip()


def _evidence(displayed: str, code_like: bool = False) -> TextEvidence:
    return TextEvidence(
        as_displayed=displayed or "",
        normalized=_normalized_from_display(displayed or ""),
        reconstructed=reconstruct_testsavvy_text(displayed or "", code_like=code_like),
    )


def _words(page: fitz.Page) -> List[Dict[str, Any]]:
    out = []
    for w in page.get_text("words"):
        out.append({
            "x0": float(w[0]), "y0": float(w[1]), "x1": float(w[2]), "y1": float(w[3]),
            "text": str(w[4]), "block": int(w[5]), "line": int(w[6]), "word": int(w[7]),
        })
    return out


def _clip_text(page: fitz.Page, rect: fitz.Rect) -> str:
    return page.get_text("text", clip=rect).strip()


def _same_line(words: Sequence[Dict[str, Any]], y: float, tolerance: float = 2.5) -> List[Dict[str, Any]]:
    return [w for w in words if abs(w["y0"] - y) <= tolerance]


def _int_or_none(value: str) -> Optional[int]:
    try:
        return int(str(value).strip())
    except Exception:
        return None


def _safe_component(value: str, max_len: int = 120) -> str:
    text = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip())
    text = re.sub(r"_+", "_", text).strip("_.")
    return (text or "unknown")[:max_len]


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Execution-boundary and summary parsing
# ---------------------------------------------------------------------------

def find_execution_ranges(doc: fitz.Document) -> List[Tuple[int, int]]:
    """Return zero-based inclusive page ranges, one per Execution Summary."""
    starts: List[int] = []
    for pno in range(doc.page_count):
        text = doc[pno].get_text("text")
        if "Execution Summary" in text and "Test Case:" in text and "Run Id:" in text:
            starts.append(pno)
    ranges: List[Tuple[int, int]] = []
    for idx, start in enumerate(starts):
        end = (starts[idx + 1] - 1) if idx + 1 < len(starts) else doc.page_count - 1
        ranges.append((start, end))
    return ranges


def _summary_value(words: Sequence[Dict[str, Any]], y: float, x0: float, x1: float, y_span: float = 14.0) -> str:
    selected = [w for w in words if x0 <= w["x0"] < x1 and y - 1 <= w["y0"] <= y + y_span]
    return _clean_ws(_join_display_words(selected))


def parse_execution_summary(page: fitz.Page) -> Dict[str, str]:
    words = _words(page)

    # Find known left-side labels by their first word. The report columns are
    # stable, while vertical positions can shift when Test Case Name wraps.
    anchors: Dict[str, float] = {}
    for w in words:
        if w["x0"] < 150:
            t = w["text"]
            if t == "Test" and any(x["text"] == "Case:" for x in _same_line(words, w["y0"])):
                anchors["test_case"] = w["y0"]
            elif t == "Run" and any(x["text"] == "Id:" for x in _same_line(words, w["y0"])):
                anchors["run_id"] = w["y0"]
            elif t == "Execution" and any(x["text"] == "Date:" for x in _same_line(words, w["y0"])):
                anchors["execution_date"] = w["y0"]
            elif t == "Received" and any(x["text"] == "Date:" for x in _same_line(words, w["y0"])):
                anchors["received_date"] = w["y0"]
            elif t == "Executed" and any(x["text"] == "By:" for x in _same_line(words, w["y0"])):
                anchors["executed_by"] = w["y0"]

    # Right-side labels.
    for w in words:
        if 400 <= w["x0"] < 520:
            same = _same_line(words, w["y0"])
            line_text = " ".join(x["text"] for x in sorted(same, key=lambda x: x["x0"]))
            if "Run Status:" in line_text:
                anchors["run_status"] = w["y0"]
            elif "Machine Name:" in line_text:
                anchors["machine_name"] = w["y0"]
            elif "Reason:" in line_text and "Override" not in line_text:
                anchors["reason"] = w["y0"]
            elif "Override Reason:" in line_text:
                anchors["override_reason"] = w["y0"]
            elif "Reported Problems:" in line_text:
                anchors["reported_problems"] = w["y0"]

    test_case_value = _summary_value(words, anchors.get("test_case", 42.0), 190, 418, 25)
    m = re.match(r"\s*(\d+)\s*-\s*(.*)", test_case_value)
    tc_id = m.group(1) if m else ""
    tc_name = m.group(2).strip() if m else test_case_value

    return {
        "test_case_id": tc_id,
        "test_case_name": tc_name,
        "run_status": _summary_value(words, anchors.get("run_status", anchors.get("test_case", 42.0)), 545, 760, 14),
        "run_id": _summary_value(words, anchors.get("run_id", 68.0), 190, 418, 14),
        "machine_name": _summary_value(words, anchors.get("machine_name", anchors.get("run_id", 68.0)), 545, 760, 14),
        "execution_date": _summary_value(words, anchors.get("execution_date", 87.0), 190, 418, 14),
        "reason": _summary_value(words, anchors.get("reason", anchors.get("execution_date", 87.0)), 545, 760, 14),
        "received_date": _summary_value(words, anchors.get("received_date", 105.0), 190, 418, 14),
        "override_reason": _summary_value(words, anchors.get("override_reason", anchors.get("received_date", 105.0)), 545, 760, 14),
        "executed_by": _summary_value(words, anchors.get("executed_by", 123.0), 190, 418, 14),
        "reported_problems": _summary_value(words, anchors.get("reported_problems", anchors.get("executed_by", 123.0)), 545, 760, 14),
    }


def parse_iteration_totals(page: fitz.Page) -> Dict[str, Optional[int]]:
    words = _words(page)
    header_y = None
    for w in words:
        if 35 <= w["x0"] < 100 and w["text"] == "Iteration":
            header_y = w["y0"]
            break
    if header_y is None:
        return {"iteration": None, "steps_passed": None, "steps_failed": None, "steps_not_run": None, "post_verify": None}

    # First numeric data row beneath the orange iteration header.
    numeric = [w for w in words if header_y + 10 < w["y0"] < header_y + 45 and re.fullmatch(r"\d+", w["text"])]
    if not numeric:
        return {"iteration": None, "steps_passed": None, "steps_failed": None, "steps_not_run": None, "post_verify": None}
    # Pick row with the most numbers.
    groups: Dict[int, List[Dict[str, Any]]] = {}
    for w in numeric:
        key = int(round(w["y0"]))
        groups.setdefault(key, []).append(w)
    row = max(groups.values(), key=len)

    def number_at(lo: float, hi: float) -> Optional[int]:
        vals = [w for w in row if lo <= w["x0"] < hi]
        return _int_or_none(vals[0]["text"]) if vals else None

    return {
        "iteration": number_at(35, 110),
        "steps_passed": number_at(110, 190),
        "steps_failed": number_at(190, 255),
        "steps_not_run": number_at(255, 320),
        "post_verify": number_at(320, 390),
    }


# ---------------------------------------------------------------------------
# Scenario and step parsing
# ---------------------------------------------------------------------------

def detect_scenarios(page: fitz.Page, source_page: int) -> List[ScenarioRecord]:
    words = _words(page)
    rows: List[ScenarioRecord] = []

    # A scenario data row has four numeric totals in the gray summary columns.
    y_candidates = sorted({w["y0"] for w in words if 210 <= w["x0"] < 450 and re.fullmatch(r"\d+", w["text"])})
    for y in y_candidates:
        same_band = [w for w in words if abs(w["y0"] - y) <= 2.6]
        nums = [w for w in same_band if 210 <= w["x0"] < 450 and re.fullmatch(r"\d+", w["text"])]
        if len(nums) < 4:
            continue
        # Ignore iteration total rows; scenario name lives near x=86.
        left = [w for w in words if 80 <= w["x0"] < 220 and y - 2 <= w["y0"] <= y + 28]
        # Result header can overlap vertically with a wrapped scenario name.
        # Exclude only the known header token rather than truncating by y.
        left = [w for w in left if w["text"] not in {"Result", "Scenario"}]
        name = _clean_ws(_join_display_words(left))
        if not name or name == "Scenario" or name.startswith("Iteration"):
            continue
        # Strip a leading iteration number occasionally carried at x~77.
        name = re.sub(r"^\d+\s+", "", name).strip()

        def val(lo: float, hi: float) -> Optional[int]:
            candidates = [w for w in nums if lo <= w["x0"] < hi]
            return _int_or_none(candidates[0]["text"]) if candidates else None

        rows.append(ScenarioRecord(
            scenario_order=0,
            scenario=name,
            steps_passed=val(210, 265),
            steps_failed=val(265, 330),
            steps_not_run=val(330, 400),
            post_verify=val(400, 450),
            source_page=source_page,
            y=float(y),
        ))

    # Deduplicate rows caused by multiple numeric words sharing nearly same y.
    dedup: List[ScenarioRecord] = []
    for row in sorted(rows, key=lambda r: r.y):
        if dedup and abs(dedup[-1].y - row.y) < 3 and dedup[-1].scenario == row.scenario:
            continue
        dedup.append(row)
    return dedup


def _step_end_marker(words: Sequence[Dict[str, Any]], y: float) -> Optional[Tuple[float, str]]:
    markers: List[Tuple[float, str]] = []
    for w in words:
        if w["y0"] <= y + 2:
            continue
        if 225 <= w["x0"] < 300 and w["text"] == "Screenshot":
            markers.append((w["y0"], "screenshot"))
        if 280 <= w["x0"] < 310 and w["text"] == "Dataset":
            same = _same_line(words, w["y0"])
            if any(x["text"] == "Timestamp" for x in same):
                markers.append((w["y0"], "metadata"))
        if 150 <= w["x0"] < 190 and w["text"] == "Result":
            markers.append((w["y0"], "next_header"))
    return min(markers, key=lambda item: item[0]) if markers else None


def detect_steps(page: fitz.Page, source_page: int) -> List[Dict[str, Any]]:
    words = _words(page)
    found: List[Dict[str, Any]] = []

    for token in words:
        if not (150 <= token["x0"] < 230 and token["text"] in {"Pass", "Fail", "Not"}):
            continue
        y = token["y0"]
        marker = _step_end_marker(words, y)
        if marker is None:
            # A legitimate step value row can begin near the bottom of a page
            # while its Screenshot / Dataset metadata starts on the next page.
            # Use page end as a provisional boundary; the Action Name + English
            # Text validation below prevents carried-over Result cells from
            # becoming false steps.
            y_end, marker_kind = float(page.rect.height), "page_end"
        else:
            y_end, marker_kind = marker
        row_words = [w for w in words if y - 1 <= w["y0"] < y_end - 1]

        fields: Dict[str, str] = {}
        for name, x0, x1 in STEP_COLUMNS:
            display = _join_display_words([w for w in row_words if x0 <= w["x0"] < x1])
            fields[name] = display

        # Anchor Result on the detected token itself. A scenario name can wrap
        # into the Result column at the top of a page (for example a hyphen),
        # producing visual text such as `Pass -`. The authoritative result is
        # still the anchored Pass/Fail/Not token.
        result = "Not Run" if token["text"] == "Not" else token["text"]
        action_name = _clean_ws(fields["action_name"])
        logical = _clean_ws(fields["logical"])
        physical = _clean_ws(fields["physical"])
        dataset_value = fields["dataset_value"].strip()
        english_text = fields["english_text"].strip()

        # This prevents a carried-over Result cell at the top of a continuation
        # page from being mistaken for a new executable step.
        if result not in RESULT_VALUES:
            continue
        if action_name not in KNOWN_ACTION_NAMES:
            continue
        if not _clean_ws(english_text):
            continue

        found.append({
            "result": result,
            "action_name": action_name,
            "logical": logical,
            "physical": physical,
            "dataset_value": dataset_value,
            "english_text": english_text,
            "source_page": source_page,
            "y": float(y),
            "end_y": float(y_end),
            "end_marker": marker_kind,
        })

    return found


# ---------------------------------------------------------------------------
# Metadata / Action Code / Message parsing
# ---------------------------------------------------------------------------

def _global_words(doc: fitz.Document, start: int, end: int) -> Tuple[List[Dict[str, Any]], List[Tuple[float, float]]]:
    """Flatten run words into a continuous vertical coordinate space."""
    words: List[Dict[str, Any]] = []
    page_offsets: List[Tuple[float, float]] = []
    offset = 0.0
    for pno in range(start, end + 1):
        page = doc[pno]
        height = float(page.rect.height)
        page_offsets.append((offset, height))
        for w in _words(page):
            item = dict(w)
            item["source_page"] = pno + 1
            item["local_y0"] = item["y0"]
            item["local_y1"] = item["y1"]
            item["gy0"] = offset + item["y0"]
            item["gy1"] = offset + item["y1"]
            words.append(item)
        offset += height
    return words, page_offsets


def _metadata_headers(global_words: Sequence[Dict[str, Any]]) -> List[float]:
    headers: List[float] = []
    for w in global_words:
        if not (280 <= w["x0"] < 310 and w["text"] == "Dataset"):
            continue
        same = [x for x in global_words if abs(x["gy0"] - w["gy0"]) <= 2.5]
        if any(x["text"] == "Timestamp" for x in same):
            headers.append(w["gy0"])
    return sorted(set(round(x, 3) for x in headers))


def _segment_words(global_words: Sequence[Dict[str, Any]], gy0: float, gy1: float) -> List[Dict[str, Any]]:
    return [w for w in global_words if gy0 <= w["gy0"] < gy1]


def _find_label_y(segment: Sequence[Dict[str, Any]], label: str) -> Optional[float]:
    candidates = [w["gy0"] for w in segment if 225 <= w["x0"] < 290 and w["text"] == label]
    return min(candidates) if candidates else None


def _metadata_value_cutoff(segment: Sequence[Dict[str, Any]], header_y: float, action_y: Optional[float], message_y: Optional[float]) -> float:
    """
    Metadata values appear directly beneath the red Dataset header. Prefer the
    last timestamp line when available; otherwise use a conservative band.
    """
    upper = min([x for x in [action_y, message_y] if x is not None], default=header_y + 95)
    # Limit timestamp detection to the compact metadata band. Long Selenium
    # lines can extend into the timestamp x-range and must not enlarge it.
    timestamps = [
        w for w in segment
        if 649.5 <= w["x0"] < 760
        and header_y + 4 < w["gy0"] < min(upper, header_y + 90)
        and (re.search(r"\d", w["text"]) or w["text"].startswith("("))
    ]
    if timestamps:
        return max(w["gy0"] for w in timestamps) + 3.0
    return min(upper, header_y + 42.0)


def parse_metadata_records(doc: fitz.Document, start: int, end: int) -> List[Dict[str, Any]]:
    gwords, _ = _global_words(doc, start, end)
    headers = _metadata_headers(gwords)
    if not headers:
        return []

    final_end = sum(float(doc[p].rect.height) for p in range(start, end + 1)) + 1
    records: List[Dict[str, Any]] = []
    for idx, hy in enumerate(headers):
        next_h = headers[idx + 1] if idx + 1 < len(headers) else final_end
        seg = _segment_words(gwords, hy + 3, next_h)

        action_y = _find_label_y(seg, "Action")
        message_y = _find_label_y(seg, "Message")
        meta_end = _metadata_value_cutoff(seg, hy, action_y, message_y)
        meta_words = [w for w in seg if w["gy0"] <= meta_end]

        meta: Dict[str, str] = {}
        for name, x0, x1 in META_COLUMNS:
            display = _join_display_words([w for w in meta_words if x0 <= w["x0"] < x1])
            meta[name] = _clean_ws(display)

        # Action Code: all value-column text after metadata values and before
        # Message. The red Action Code label itself is outside x>=287.
        code_end = message_y if message_y is not None else next_h
        code_words = [w for w in seg if w["gy0"] > meta_end and w["gy0"] < code_end and w["x0"] >= 286.5]
        code_display = _join_display_words(code_words)

        # Message: text to the right of the Message label, ending before the
        # next structural scenario/result section. We use only a compact band
        # after Message because TestSavvy messages are normally local to it.
        message_display = ""
        if message_y is not None:
            raw_candidates = [w for w in seg if message_y - 2 <= w["gy0"] < min(next_h, message_y + 120)]
            structural_y: List[float] = []
            # New Result header.
            structural_y.extend(
                w["gy0"] for w in raw_candidates
                if 150 <= w["x0"] < 190 and w["text"] == "Result" and w["gy0"] > message_y + 3
            )
            # New scenario summary row: four numeric totals in the gray columns.
            ys = sorted({round(w["gy0"], 2) for w in raw_candidates if 210 <= w["x0"] < 450 and re.fullmatch(r"\d+", w["text"])})
            for sy in ys:
                same = [w for w in raw_candidates if abs(w["gy0"] - sy) <= 2.6]
                if len([w for w in same if 210 <= w["x0"] < 450 and re.fullmatch(r"\d+", w["text"])]) >= 4:
                    structural_y.append(float(sy))
            cutoff = min(structural_y) if structural_y else min(next_h, message_y + 120)
            message_candidates = [
                w for w in raw_candidates
                if message_y - 2 <= w["gy0"] < cutoff and w["x0"] >= 286.5
            ]
            message_display = _join_display_words(message_candidates)

        first_page = min((w["source_page"] for w in meta_words), default=start + 1)
        records.append({
            **meta,
            "action_code": code_display,
            "message": message_display,
            "source_page": first_page,
            "global_y": hy,
        })
    return records


# ---------------------------------------------------------------------------
# Screenshot extraction / OCR
# ---------------------------------------------------------------------------

def _image_occurrences(page: fitz.Page) -> List[Dict[str, Any]]:
    out = []
    try:
        for info in page.get_image_info(xrefs=True):
            bbox = info.get("bbox")
            if not bbox:
                continue
            out.append({
                "xref": int(info.get("xref") or 0),
                "bbox": [float(v) for v in bbox],
                "width": int(info.get("width") or 0),
                "height": int(info.get("height") or 0),
            })
    except Exception:
        pass
    return out


def _extract_occurrence(doc: fitz.Document, occurrence: Dict[str, Any], dest: Path) -> Tuple[int, int]:
    xref = occurrence.get("xref", 0)
    if xref:
        payload = doc.extract_image(xref)
        ext = payload.get("ext", "png")
        actual = dest.with_suffix("." + ext)
        actual.parent.mkdir(parents=True, exist_ok=True)
        actual.write_bytes(payload["image"])
        if actual != dest:
            # Caller receives actual path by checking generated files.
            pass
        return int(payload.get("width") or occurrence.get("width") or 0), int(payload.get("height") or occurrence.get("height") or 0)
    return 0, 0


def _ocr_image(path: Path, language: str = "eng") -> Tuple[str, str]:
    """
    OCR is supplemental evidence only. Preserve the original screenshot as the
    visual source of truth. Upscaling improves low-resolution screenshots from
    very large combined PDFs.
    """
    if pytesseract is None or Image is None:
        return "", "unavailable"
    try:
        with Image.open(path) as source:
            img = source.convert("RGB")
            scale = 4 if img.width < 1000 else 2
            img = img.resize((img.width * scale, img.height * scale))
            # Sparse and table-heavy application screens generally work better
            # with psm 6 than the default on these report screenshots.
            text = pytesseract.image_to_string(img, lang=language, config="--psm 6")
        return text.strip(), "complete"
    except Exception as exc:
        return "", f"error:{type(exc).__name__}"


def associate_and_extract_screenshots(
    doc: fitz.Document,
    run_start: int,
    steps: List[StepRecord],
    screenshot_dir: Path,
    ocr_mode: str,
    ocr_language: str,
) -> None:
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    # Cache image occurrences per page.
    page_cache: Dict[int, List[Dict[str, Any]]] = {}
    used_occurrences: set[Tuple[int, int, float]] = set()

    for step in steps:
        pno = step.action_page - 1
        if pno < 0 or pno >= doc.page_count:
            continue
        if pno not in page_cache:
            page_cache[pno] = _image_occurrences(doc[pno])
        occurrences = page_cache[pno]

        # Prefer images beneath the step value row. TestSavvy screenshot bboxes
        # normally occupy x~228-559 and start after the Screenshot band.
        candidates = []
        for occ in occurrences:
            bbox = occ["bbox"]
            key = (pno, occ["xref"], bbox[1])
            if key in used_occurrences:
                continue
            if bbox[1] >= step.action_y + 15 and bbox[0] >= 180 and bbox[2] <= 650:
                candidates.append(occ)
        if not candidates:
            continue
        occ = min(candidates, key=lambda o: o["bbox"][1])
        used_occurrences.add((pno, occ["xref"], occ["bbox"][1]))

        base = screenshot_dir / f"step_{step.test_case_step_order:04d}_page_{step.action_page:04d}"
        xref = occ["xref"]
        if xref:
            payload = doc.extract_image(xref)
            ext = payload.get("ext", "png")
            path = base.with_suffix("." + ext)
            path.write_bytes(payload["image"])
            width = int(payload.get("width") or occ.get("width") or 0)
            height = int(payload.get("height") or occ.get("height") or 0)
        else:
            continue

        should_ocr = False
        if ocr_mode == "all":
            should_ocr = True
        elif ocr_mode in {"failures", "failures_and_key_events"}:
            if step.result == "Fail" or "failure" in step.event_tags:
                should_ocr = True
            if ocr_mode == "failures_and_key_events" and any(tag in step.event_tags for tag in {"submit", "validate", "save", "smart_wait", "verification"}):
                should_ocr = True

        ocr_text, ocr_status = ("", "not_requested")
        if should_ocr:
            ocr_text, ocr_status = _ocr_image(path, ocr_language)

        step.screenshot = ScreenshotEvidence(
            present=True,
            file=str(path.name),
            width=width,
            height=height,
            source_page=step.action_page,
            run_page=step.action_page - run_start,
            xref=xref,
            bbox=occ["bbox"],
            sha256=_sha256(path),
            ocr_text=ocr_text,
            ocr_status=ocr_status,
        )


# ---------------------------------------------------------------------------
# Post-processing and validation
# ---------------------------------------------------------------------------

def _event_tags(step: StepRecord) -> List[str]:
    text = " ".join([
        step.scenario or "", step.action_name or "", step.logical or "",
        step.english_text.normalized or "", step.message.normalized or "",
    ]).lower()
    tags: List[str] = []
    mapping = {
        "submit": ["submit"],
        "validate": ["validate"],
        "save": ["save"],
        "smart_wait": ["smart wait"],
        "verification": ["verify", "verification"],
        "custom_function": ["custom function"],
        "sign_in": ["sign in"],
        "sign_out": ["sign out"],
        "search": ["search"],
        "create": ["create"],
    }
    for tag, needles in mapping.items():
        if any(n in text for n in needles):
            tags.append(tag)
    if step.result == "Fail":
        tags.append("failure")
    if step.result == "Not Run":
        tags.append("not_run")
    if step.action_name == "Enter Text":
        tags.append("data_entry")
    if step.action_code.normalized:
        tags.append("action_code")
    return sorted(set(tags))


def _assign_scenarios(steps: List[StepRecord], scenarios: List[ScenarioRecord]) -> None:
    scenarios_sorted = sorted(scenarios, key=lambda s: (s.source_page, s.y))
    for idx, scenario in enumerate(scenarios_sorted, start=1):
        scenario.scenario_order = idx

    per_scenario_count: Dict[int, int] = {}
    for step in steps:
        preceding = [s for s in scenarios_sorted if (s.source_page < step.action_page) or (s.source_page == step.action_page and s.y <= step.action_y)]
        if preceding:
            scenario = preceding[-1]
            step.scenario_order = scenario.scenario_order
            step.scenario = scenario.scenario
        else:
            step.scenario_order = 0
            step.scenario = "Unassigned"
        per_scenario_count[step.scenario_order] = per_scenario_count.get(step.scenario_order, 0) + 1
        step.scenario_step_order = per_scenario_count[step.scenario_order]


def _validate_execution(execution: ExecutionRecord) -> Dict[str, Any]:
    pass_count = sum(1 for s in execution.steps if s.result == "Pass")
    fail_count = sum(1 for s in execution.steps if s.result == "Fail")
    not_run_count = sum(1 for s in execution.steps if s.result == "Not Run")
    screenshot_count = sum(1 for s in execution.steps if s.screenshot.present)
    metadata_count = sum(1 for s in execution.steps if any([s.dataset, s.dataset_header, s.test_condition, s.element_type, s.timestamp, s.action_code.normalized, s.message.normalized]))

    scenario_pass = sum(s.steps_passed or 0 for s in execution.scenarios)
    scenario_fail = sum(s.steps_failed or 0 for s in execution.scenarios)
    scenario_not_run = sum(s.steps_not_run or 0 for s in execution.scenarios)

    checks = {
        "parsed_step_count": len(execution.steps),
        "parsed_pass": pass_count,
        "parsed_fail": fail_count,
        "parsed_not_run": not_run_count,
        "parsed_metadata_records": metadata_count,
        "raw_metadata_record_count": execution.raw_metadata_record_count,
        "screenshot_count": screenshot_count,
        "reported_pass": execution.reported_steps_passed,
        "reported_fail": execution.reported_steps_failed,
        "reported_not_run": execution.reported_steps_not_run,
        "scenario_reported_pass": scenario_pass,
        "scenario_reported_fail": scenario_fail,
        "scenario_reported_not_run": scenario_not_run,
    }
    comparisons = []
    for parsed_key, reported_key in [
        ("parsed_pass", "reported_pass"),
        ("parsed_fail", "reported_fail"),
        ("parsed_not_run", "reported_not_run"),
    ]:
        reported = checks[reported_key]
        if reported is not None:
            comparisons.append(checks[parsed_key] == reported)
    checks["result_totals_match_report"] = all(comparisons) if comparisons else None
    checks["metadata_count_matches_steps"] = metadata_count == len(execution.steps)
    checks["raw_metadata_count_matches_steps"] = execution.raw_metadata_record_count == len(execution.steps)
    checks["scenario_totals_match_steps"] = (
        scenario_pass == pass_count and scenario_fail == fail_count and scenario_not_run == not_run_count
    )
    checks["status"] = "PASS" if (
        checks.get("result_totals_match_report") is not False
        and checks["metadata_count_matches_steps"]
        and checks["raw_metadata_count_matches_steps"]
        and checks["scenario_totals_match_steps"]
    ) else "REVIEW"
    return checks


# ---------------------------------------------------------------------------
# Core run parser
# ---------------------------------------------------------------------------

def parse_execution(doc: fitz.Document, start: int, end: int, global_step_offset: int = 0) -> ExecutionRecord:
    summary = parse_execution_summary(doc[start])
    totals = parse_iteration_totals(doc[start])

    scenarios: List[ScenarioRecord] = []
    raw_steps: List[Dict[str, Any]] = []
    for pno in range(start, end + 1):
        source_page = pno + 1
        scenarios.extend(detect_scenarios(doc[pno], source_page))
        raw_steps.extend(detect_steps(doc[pno], source_page))

    steps: List[StepRecord] = []
    for idx, raw in enumerate(raw_steps, start=1):
        steps.append(StepRecord(
            global_step_order=global_step_offset + idx,
            test_case_step_order=idx,
            scenario_step_order=0,
            scenario_order=0,
            scenario="",
            result=raw["result"],
            action_name=raw["action_name"],
            logical=raw["logical"],
            physical=raw["physical"],
            dataset_value=_evidence(raw["dataset_value"]),
            english_text=_evidence(raw["english_text"]),
            action_page=raw["source_page"],
            action_y=raw["y"],
        ))

    _assign_scenarios(steps, scenarios)

    metadata = parse_metadata_records(doc, start, end)
    for step, meta in zip(steps, metadata):
        step.dataset = meta.get("dataset", "")
        step.dataset_header = meta.get("dataset_header", "")
        step.test_condition = meta.get("test_condition", "")
        step.element_type = meta.get("element_type", "")
        step.timestamp = meta.get("timestamp", "")
        step.action_code = _evidence(meta.get("action_code", ""), code_like=True)
        step.message = _evidence(meta.get("message", ""))
        step.metadata_page = int(meta.get("source_page") or step.action_page)
        step.metadata_y = float(meta.get("global_y") or 0.0)

    for idx, step in enumerate(steps):
        step.previous_step_order = steps[idx - 1].test_case_step_order if idx > 0 else None
        step.next_step_order = steps[idx + 1].test_case_step_order if idx + 1 < len(steps) else None
        step.event_tags = _event_tags(step)

    execution = ExecutionRecord(
        source_pdf="",
        source_start_page=start + 1,
        source_end_page=end + 1,
        test_case_id=summary.get("test_case_id", ""),
        test_case_name=summary.get("test_case_name", ""),
        run_status=summary.get("run_status", ""),
        run_id=summary.get("run_id", ""),
        machine_name=summary.get("machine_name", ""),
        execution_date=summary.get("execution_date", ""),
        reason=summary.get("reason", ""),
        received_date=summary.get("received_date", ""),
        override_reason=summary.get("override_reason", ""),
        executed_by=summary.get("executed_by", ""),
        reported_problems=summary.get("reported_problems", ""),
        iteration=totals.get("iteration"),
        reported_steps_passed=totals.get("steps_passed"),
        reported_steps_failed=totals.get("steps_failed"),
        reported_steps_not_run=totals.get("steps_not_run"),
        reported_post_verify=totals.get("post_verify"),
        raw_metadata_record_count=len(metadata),
        scenarios=scenarios,
        steps=steps,
    )
    execution.validation = _validate_execution(execution)
    return execution