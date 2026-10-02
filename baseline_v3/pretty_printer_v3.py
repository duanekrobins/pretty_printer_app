#!/usr/bin/env python3
"""
TestSavvy XML / Excel / JSON Pretty Printer and Markdown Documentation Generator

Developer: Duane K Robinson
Organization: State of Utah
Date: March 11th 2026

Program Purpose
---------------
This Python program converts machine-generated automation artifacts into
human-readable Markdown documentation suitable for review, troubleshooting,
knowledge-base storage, and AI-assisted ingestion tools such as NotebookLM.

The program is configuration-driven. Runtime behavior is controlled by a JSON
configuration file passed with the --config command-line argument.

Primary Capabilities
--------------------
1. XML Processing
   - Reads XML files from configured input directories.
   - Supports raw XML pretty-print output.
   - Supports enhanced TestSavvy execution report output.
   - Extracts TestSavvy automation_sequence metadata.
   - Maps automation_sequence id to both Test Case ID and Automation Sequence ID.
   - Extracts scenario metadata, scenario attributes, step metadata, runtime
     variables, machine settings, execution summary fields, and raw attributes.
   - Optionally appends a full XML Inventory section listing every parsed node,
     node path, depth, attributes, and text.

2. Markdown Generation
   - Writes Markdown files using the configured output extension.
   - Supports one output Markdown file per input file.
   - Supports combined Markdown output for multiple input files.
   - Structures TestSavvy XML data into sections suitable for NotebookLM,
     repository documentation, and human QA review.

3. Excel Processing
   - Reads .xlsx and .xlsm workbooks.
   - Converts each worksheet/tab into a Markdown table.
   - Supports one Markdown output per worksheet.
   - Supports combined output when configured.

4. JSON Processing
   - Pretty prints JSON.
   - Can wrap pretty-printed JSON in Markdown code fences when output is .md.

Recommended NotebookLM Use
--------------------------
For NotebookLM ingestion, use:
- xml_output_mode: "testsavvy_execution_report"
- include_xml_inventory: true
- output_extension: ".md"

This gives NotebookLM both a human-readable execution report and a complete
node/attribute inventory so XML details are not silently omitted.

Typical Command
---------------
python pretty_printer.py --config pretty_print_config.json

Important Notes
---------------
- This utility is intentionally read-only with respect to source files.
- Output files are written only to the configured output directory.
- overwrite=true allows regenerated output to replace previous output.
- The parser is optimized for TestSavvy execution XML payloads but also retains
  raw XML output support for debugging and exact source review.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import re
import io
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Tuple, Dict, Any
import xml.etree.ElementTree as ET

from openpyxl import load_workbook


@dataclass(frozen=True)
class AppConfig:
    input_dirs: List[Path]
    output_dir: Path
    file_types: Tuple[str, ...]
    recursive: bool
    overwrite: bool
    indent: int
    json_sort_keys: bool
    preserve_xml_declaration: bool
    xml_encoding: str
    output_structure: str
    output_extension: Optional[str]
    exclude_patterns: Tuple[str, ...]
    xml_markdown_mode: str
    combined_output_filename: str
    xml_output_mode: str
    excel_markdown_mode: str
    excel_include_empty_sheets: bool
    excel_max_cell_length: int
    excel_sheet_name_in_filename: bool
    include_xml_inventory: bool
    generate_notebooklm_guide: bool
    notebooklm_guide_filename: str
    notebooklm_enhanced_format: bool
    normalize_markdown_filenames: bool
    json_output_mode: str
    include_json_inventory: bool
    json_inventory_max_rows: int
    json_table_max_rows: int
    process_zip_files: bool


def _normalize_extension(ext: Optional[str]) -> Optional[str]:
    """
    Normalize a configured output extension so values like 'md' and '.md' behave consistently.
    """
    if ext is None:
        return None
    ext = str(ext).strip()
    if ext == "":
        return ""
    if not ext.startswith("."):
        ext = "." + ext
    return ext.lower()



# -----------------------------------------------------------------------------
# Configuration loading and validation
# -----------------------------------------------------------------------------
def load_config(config_path: Path) -> AppConfig:
    """
    Read the JSON configuration file, validate key options, normalize paths and extensions, and return an immutable AppConfig object.
    """
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    raw = json.loads(config_path.read_text(encoding="utf-8"))

    input_dirs = [Path(p).expanduser().resolve() for p in raw.get("input_dirs", [])]
    if not input_dirs:
        raise ValueError("Config must include 'input_dirs' with at least one directory.")

    output_dir = Path(raw.get("output_dir", "./output")).expanduser().resolve()

    file_types_raw = raw.get("file_types", ["xml", "json", "zip", "xlsx", "xlsm"])
    file_types = tuple(ft.lower().strip(".") for ft in file_types_raw)
    allowed = {"json", "xml", "zip", "xlsx", "xlsm"}
    unknown = [ft for ft in file_types if ft not in allowed]
    if unknown:
        raise ValueError(f"Unknown file types in config: {unknown}. Allowed: {sorted(allowed)}")

    recursive = bool(raw.get("recursive", True))
    overwrite = bool(raw.get("overwrite", False))
    indent = int(raw.get("indent", 4))
    json_sort_keys = bool(raw.get("json_sort_keys", True))
    preserve_xml_declaration = bool(raw.get("preserve_xml_declaration", True))
    xml_encoding = str(raw.get("xml_encoding", "utf-8"))

    output_structure = str(raw.get("output_structure", "mirror")).lower()
    if output_structure not in {"mirror", "flat"}:
        raise ValueError("output_structure must be 'mirror' or 'flat'.")

    output_extension = _normalize_extension(raw.get("output_extension", ".md"))
    exclude_patterns = tuple(raw.get("exclude_patterns", []))

    xml_markdown_mode = str(raw.get("xml_markdown_mode", "per_file")).lower()
    if xml_markdown_mode not in {"per_file", "combined"}:
        raise ValueError("xml_markdown_mode must be 'per_file' or 'combined'.")

    xml_output_mode = str(raw.get("xml_output_mode", "testsavvy_execution_report")).lower()
    if xml_output_mode not in {"testsavvy_execution_report", "raw_markdown"}:
        raise ValueError("xml_output_mode must be 'testsavvy_execution_report' or 'raw_markdown'.")

    combined_output_filename = str(raw.get("combined_output_filename", "combined_output.md"))

    excel_markdown_mode = str(raw.get("excel_markdown_mode", "per_sheet")).lower()
    if excel_markdown_mode not in {"per_sheet"}:
        raise ValueError("excel_markdown_mode currently supports only 'per_sheet'.")

    include_xml_inventory = bool(raw.get("include_xml_inventory", True))
    generate_notebooklm_guide = bool(raw.get("generate_notebooklm_guide", True))
    notebooklm_guide_filename = str(raw.get("notebooklm_guide_filename", "NOTEBOOKLM_TESTSAVVY_PAYLOAD_GUIDE.md"))
    notebooklm_enhanced_format = bool(raw.get("notebooklm_enhanced_format", True))
    normalize_markdown_filenames = bool(raw.get("normalize_markdown_filenames", True))

    json_output_mode = str(raw.get("json_output_mode", "testsavvy_smart")).lower()
    if json_output_mode not in {"testsavvy_smart", "raw_markdown"}:
        raise ValueError("json_output_mode must be 'testsavvy_smart' or 'raw_markdown'.")

    include_json_inventory = bool(raw.get("include_json_inventory", True))
    json_inventory_max_rows = int(raw.get("json_inventory_max_rows", 5000))
    json_table_max_rows = int(raw.get("json_table_max_rows", 2000))
    process_zip_files = bool(raw.get("process_zip_files", True))

    return AppConfig(
        input_dirs=input_dirs,
        output_dir=output_dir,
        file_types=file_types,
        recursive=recursive,
        overwrite=overwrite,
        indent=indent,
        json_sort_keys=json_sort_keys,
        preserve_xml_declaration=preserve_xml_declaration,
        xml_encoding=xml_encoding,
        output_structure=output_structure,
        output_extension=output_extension,
        exclude_patterns=exclude_patterns,
        xml_markdown_mode=xml_markdown_mode,
        combined_output_filename=combined_output_filename,
        xml_output_mode=xml_output_mode,
        excel_markdown_mode=excel_markdown_mode,
        excel_include_empty_sheets=bool(raw.get("excel_include_empty_sheets", False)),
        excel_max_cell_length=int(raw.get("excel_max_cell_length", 500)),
        excel_sheet_name_in_filename=bool(raw.get("excel_sheet_name_in_filename", True)),
        include_xml_inventory=include_xml_inventory,
        generate_notebooklm_guide=generate_notebooklm_guide,
        notebooklm_guide_filename=notebooklm_guide_filename,
        notebooklm_enhanced_format=notebooklm_enhanced_format,
        normalize_markdown_filenames=normalize_markdown_filenames,
        json_output_mode=json_output_mode,
        include_json_inventory=include_json_inventory,
        json_inventory_max_rows=json_inventory_max_rows,
        json_table_max_rows=json_table_max_rows,
        process_zip_files=process_zip_files,
    )


def is_excluded(path: Path, patterns: Tuple[str, ...]) -> bool:
    """
    Return True when a source file name matches one of the configured exclude glob patterns.
    """
    return any(fnmatch.fnmatch(path.name, pat) for pat in patterns) if patterns else False


def iter_files(input_dir: Path, recursive: bool, file_types: Tuple[str, ...]) -> Iterable[Path]:
    """
    Yield source files from an input directory according to recursion and configured file extensions.
    """
    if not input_dir.exists():
        return
    walker = input_dir.rglob if recursive else input_dir.glob
    seen = set()
    for ft in file_types:
        for path in walker(f"*.{ft}"):
            if path not in seen:
                seen.add(path)
                yield path


def apply_output_extension(filename: str, forced_ext: Optional[str]) -> str:
    """
    Apply the configured output extension to a file name while preserving the base stem.
    """
    if forced_ext is None or forced_ext == "":
        return filename
    return Path(filename).with_suffix(forced_ext).name


def compute_output_path(src: Path, input_root: Path, output_dir: Path, output_structure: str, forced_ext: Optional[str]) -> Path:
    """
    Build the destination output path for an input file using either mirror or flat output structure.
    """
    if output_structure == "flat":
        return output_dir / apply_output_extension(src.name, forced_ext)
    rel = src.relative_to(input_root)
    return output_dir / rel.with_name(apply_output_extension(rel.name, forced_ext))


def ensure_parent_dir(path: Path) -> None:
    """
    Create the parent directory for an output file if it does not already exist.
    """
    path.parent.mkdir(parents=True, exist_ok=True)


def pretty_print_json(src: Path, indent: int, sort_keys: bool) -> str:
    """
    Load JSON and serialize it back out with configured indentation and optional sorted keys.
    """
    text = src.read_text(encoding="utf-8")
    obj = json.loads(text)
    return json.dumps(obj, indent=indent, ensure_ascii=False, sort_keys=sort_keys) + "\n"


def pretty_print_xml(src: Path, indent: int, encoding: str, preserve_declaration: bool) -> bytes:
    """
    Parse XML, apply indentation, and serialize it back to bytes using the configured XML encoding and declaration preference.
    """
    tree = ET.parse(str(src), parser=ET.XMLParser())
    ET.indent(tree, space=" " * indent)
    buf = io.BytesIO()
    tree.write(
        buf,
        encoding=encoding,
        xml_declaration=preserve_declaration,
        short_empty_elements=True,
    )
    buf.write(b"\n")
    return buf.getvalue()


def wrap_xml_markdown(src: Path, xml_text: str) -> str:
    """
    Wrap pretty-printed XML text inside a Markdown XML code fence.
    """
    return f"# Pretty Printed XML\n\n**Source file:** `{src.name}`\n\n```xml\n{xml_text.rstrip()}\n```\n"


def wrap_json_markdown(src: Path, json_text: str) -> str:
    """
    Wrap pretty-printed JSON text inside a Markdown JSON code fence.
    """
    return f"# Pretty Printed JSON\n\n**Source file:** `{src.name}`\n\n```json\n{json_text.rstrip()}\n```\n"


def sanitize_sheet_name(sheet_name: str) -> str:
    """
    Convert an Excel worksheet name into a safe filename component.
    """
    value = re.sub(r"[^\w\-. ]+", "_", sheet_name.strip())
    value = re.sub(r"\s+", "_", value)
    return value or "Sheet"


def markdown_escape(value: object, max_len: int = 500) -> str:
    """
    Escape Markdown table-sensitive characters and normalize whitespace so cell content renders safely.
    """
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    text = text.replace("|", "\\|")
    text = re.sub(r"\s+", " ", text).strip()
    if max_len > 0 and len(text) > max_len:
        text = text[: max_len - 3] + "..."
    return text


def worksheet_has_data(ws) -> bool:
    """
    Check whether an Excel worksheet contains at least one non-empty cell.
    """
    for row in ws.iter_rows(values_only=True):
        if any(cell not in (None, "") for cell in row):
            return True
    return False



# -----------------------------------------------------------------------------
# Excel worksheet to Markdown conversion
# -----------------------------------------------------------------------------
def worksheet_to_markdown(src: Path, ws, max_cell_length: int) -> str:
    """
    Convert an Excel worksheet into a Markdown table using the first non-empty row as the header row.
    """
    rows = list(ws.iter_rows(values_only=True))
    non_empty_rows = []
    max_cols = 0

    for row in rows:
        trimmed = list(row)
        while trimmed and trimmed[-1] in (None, ""):
            trimmed.pop()
        if trimmed:
            non_empty_rows.append(trimmed)
            max_cols = max(max_cols, len(trimmed))

    if not non_empty_rows:
        return (
            f"# Excel Worksheet Export\n\n"
            f"**Workbook:** `{src.name}`\n\n"
            f"**Worksheet:** `{ws.title}`\n\n"
            f"_This worksheet is empty._\n"
        )

    normalized = []
    for row in non_empty_rows:
        padded = list(row) + [""] * (max_cols - len(row))
        normalized.append([markdown_escape(cell, max_cell_length) for cell in padded])

    header = normalized[0]
    if all(h == "" for h in header):
        header = [f"Column {i+1}" for i in range(max_cols)]
        data_rows = normalized
    else:
        header = [h if h else f"Column {i+1}" for i, h in enumerate(header)]
        data_rows = normalized[1:]

    lines = [
        "# Excel Worksheet Export",
        "",
        f"**Workbook:** `{src.name}`",
        "",
        f"**Worksheet:** `{ws.title}`",
        "",
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]

    for row in data_rows:
        lines.append("| " + " | ".join(row) + " |")

    if not data_rows:
        lines.append("| " + " | ".join([""] * len(header)) + " |")

    lines.append("")
    return "\n".join(lines)


def build_excel_sheet_output_path(
    src: Path,
    input_root: Path,
    output_dir: Path,
    output_structure: str,
    forced_ext: Optional[str],
    sheet_name: str,
    include_sheet_name_in_filename: bool,
) -> Path:
    """
    Compute the output Markdown path for a single Excel worksheet export.
    """
    ext = forced_ext if forced_ext not in (None, "") else ".md"
    safe_sheet = sanitize_sheet_name(sheet_name)

    if output_structure == "flat":
        base_name = src.stem
        out_name = f"{base_name}__{safe_sheet}{ext}" if include_sheet_name_in_filename else f"{safe_sheet}{ext}"
        return output_dir / out_name

    rel = src.relative_to(input_root)
    parent = output_dir / rel.parent / src.stem
    out_name = f"{safe_sheet}{ext}" if include_sheet_name_in_filename else apply_output_extension(src.name, ext)
    return parent / out_name


def process_excel_workbook(src: Path, cfg: AppConfig, input_root: Path):
    """
    Process an Excel workbook and write one Markdown file per worksheet when in per-sheet mode.
    """
    messages = []
    written = 0
    failed = 0

    wb = load_workbook(filename=str(src), data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            if not cfg.excel_include_empty_sheets and not worksheet_has_data(ws):
                messages.append(f"SKIP (empty sheet): {src} [{ws.title}]")
                continue

            out_path = build_excel_sheet_output_path(
                src=src,
                input_root=input_root,
                output_dir=cfg.output_dir,
                output_structure=cfg.output_structure,
                forced_ext=cfg.output_extension,
                sheet_name=ws.title,
                include_sheet_name_in_filename=cfg.excel_sheet_name_in_filename,
            )

            if out_path.exists() and not cfg.overwrite:
                messages.append(f"SKIP (exists): {src} [{ws.title}]")
                continue

            try:
                ensure_parent_dir(out_path)
                md = worksheet_to_markdown(src, ws, cfg.excel_max_cell_length)
                out_path.write_text(md, encoding="utf-8")
                messages.append(f"OK: {src} [{ws.title}] -> {out_path}")
                written += 1
            except Exception as e:
                messages.append(f"FAIL: {src} [{ws.title}] ({type(e).__name__}: {e})")
                failed += 1
    finally:
        wb.close()

    return written, failed, messages


def local_name(tag: str) -> str:
    """
    Return an XML tag name without namespace decoration.
    """
    return tag.split("}", 1)[-1] if "}" in tag else tag


def element_children_map(elem: ET.Element) -> Dict[str, List[ET.Element]]:
    """
    Build a dictionary of child XML elements grouped by local tag name.
    """
    result: Dict[str, List[ET.Element]] = {}
    for child in list(elem):
        result.setdefault(local_name(child.tag), []).append(child)
    return result


def element_text(elem: Optional[ET.Element]) -> str:
    """Fixed to only return immediate node text, preventing recursive concatenation."""
    if elem is None or elem.text is None:
        return ""
    return elem.text.strip()


def first_text_anywhere(root: ET.Element, names: List[str]) -> str:
    """
    Search the XML tree for the first matching tag name and return its text content.
    """
    wanted = {n.lower() for n in names}
    for elem in root.iter():
        if local_name(elem.tag).lower() in wanted:
            txt = element_text(elem)
            if txt:
                return txt
    return ""


def find_execution_nodes(root: ET.Element, names: List[str]) -> List[ET.Element]:
    """
    Find XML elements whose local tag names match the supplied list of candidate names.
    """
    wanted = {n.lower() for n in names}
    found = []
    for elem in root.iter():
        if local_name(elem.tag).lower() in wanted:
            found.append(elem)
    return found


def parse_datetime_text(text: str) -> str:
    """
    Attempt to normalize known datetime formats while preserving unrecognized strings.
    """
    if not text:
        return ""
    candidates = [
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ]
    for fmt in candidates:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%b %-d %Y %-I:%M%p %z") if "%" in fmt else dt.strftime("%b %-d %Y %-I:%M%p")
        except Exception:
            pass
    return text


def normalize_bool_text(value: str) -> str:
    """
    Normalize boolean-like text values to Yes or No when possible.
    """
    v = (value or "").strip().lower()
    if v in {"true", "1", "yes", "y"}:
        return "Yes"
    if v in {"false", "0", "no", "n", ""}:
        return "No"
    return value



# -----------------------------------------------------------------------------
# TestSavvy XML extraction helpers
# -----------------------------------------------------------------------------
def extract_runtime_variables(root: ET.Element) -> List[Tuple[str, str]]:
    """
    Extract runtime variables from TestSavvy runtime variable nodes and attribute-based variable entries.
    """
    vars_out: List[Tuple[str, str]] = []

    candidate_parents = []
    for elem in root.iter():
        name = local_name(elem.tag).lower()
        if name in {"runtimevariables", "runtime_variables", "variables", "runtimevars", "variablelist"}:
            candidate_parents.append(elem)

    def add_pair(k: str, v: str):
        if k:
            vars_out.append((k, v))

    for parent in candidate_parents:
        for elem in parent.iter():
            lname = local_name(elem.tag).lower()
            if lname in {"variable", "runtimevariable", "var", "entry"}:
                key = ""
                val = ""

                for attr_name in ["name", "key"]:
                    if elem.attrib.get(attr_name):
                        key = elem.attrib.get(attr_name, "")
                        break
                for attr_name in ["value"]:
                    if elem.attrib.get(attr_name):
                        val = elem.attrib.get(attr_name, "")
                        break

                if not key or not val:
                    children = element_children_map(elem)
                    for kn in ["name", "key", "variableName", "varName"]:
                        if key:
                            break
                        for c in children.get(kn, []):
                            key = element_text(c)
                            if key:
                                break
                    for vn in ["value", "variableValue", "varValue"]:
                        if val:
                            break
                        for c in children.get(vn, []):
                            val = element_text(c)
                            if val:
                                break

                if key:
                    add_pair(key, val)

    if not vars_out:
        for elem in root.iter():
            lname = local_name(elem.tag).lower()
            if lname in {"variable", "runtimevariable", "var"}:
                key = elem.attrib.get("name") or elem.attrib.get("key") or ""
                val = elem.attrib.get("value") or ""
                if key:
                    add_pair(key, val)

    seen = set()
    deduped = []
    for item in vars_out:
        if item not in seen:
            seen.add(item)
            deduped.append(item)
    return deduped


def extract_step_records(scenario_elem: ET.Element) -> List[Dict[str, str]]:
    """
    Extract TestSavvy step rows from scenario XML, including key attributes and a raw attribute inventory.
    """
    steps: List[Dict[str, str]] = []
    for elem in scenario_elem.iter():
        lname = local_name(elem.tag).lower()
        if lname != "step":
            continue

        dataset = elem.attrib.get("dataset", "") or ""
        header = elem.attrib.get("dataset_header", "") or ""
        dataset_header = f"{dataset} / {header}" if dataset and header else (dataset or header)

        row = {
            "#": elem.attrib.get("count_id") or elem.attrib.get("sequence") or elem.attrib.get("seq") or elem.attrib.get("id") or "",
            "Step ID": elem.attrib.get("step_id", ""),
            "Action": elem.attrib.get("action_name") or elem.attrib.get("actionName") or elem.attrib.get("name") or "",
            "Type": elem.attrib.get("action_type") or elem.attrib.get("actionType") or elem.attrib.get("type") or "",
            "English Text": elem.attrib.get("english_text") or elem.attrib.get("englishText") or "",
            "Dataset / Header": dataset_header,
            "Test Condition": elem.attrib.get("test_condition") or elem.attrib.get("testCondition") or "",
            "Encrypted": "Yes" if str(elem.attrib.get("value_encrypted", "")).strip().lower() in {"1", "true", "yes", "y"} else "No",
            "Value": elem.attrib.get("value", ""),
            "Application": elem.attrib.get("application_name", ""),
            "Interface Map": elem.attrib.get("interface_map", ""),
            "Logical Name": elem.attrib.get("logical_name", ""),
            "Interface Element Type": elem.attrib.get("interface_element_type", ""),
            "Tool": elem.attrib.get("toolplugin") or elem.attrib.get("tool") or "",
            "Comments": elem.attrib.get("comments", ""),
            "Action Code": elem.attrib.get("action_code", ""),
            "Raw Attributes": "; ".join(f"{k}={v}" for k, v in elem.attrib.items()),
        }
        steps.append(row)

    deduped = []
    seen = set()
    for row in steps:
        key = tuple(row.values())
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    return deduped


def extract_automation_sequence(root: ET.Element) -> Optional[ET.Element]:
    """
    Locate the top-level TestSavvy automation_sequence element in the XML document.
    """
    for elem in root.iter():
        if local_name(elem.tag).lower() == "automation_sequence":
            return elem
    return None


def extract_scenarios(root: ET.Element) -> List[Dict[str, Any]]:
    """
    Extract TestSavvy scenario metadata and associated step records from the automation sequence.
    """
    scenarios: List[Dict[str, Any]] = []

    auto_seq = extract_automation_sequence(root)
    candidate_root = auto_seq if auto_seq is not None else root

    for elem in candidate_root.iter():
        lname = local_name(elem.tag).lower()
        if lname != "scenario":
            continue

        title = elem.attrib.get("name") or elem.attrib.get("scenarioName") or elem.attrib.get("title") or "Scenario"
        sid = elem.attrib.get("id") or elem.attrib.get("scenarioId") or ""
        path = elem.attrib.get("path") or ""
        auto_scenario_id = elem.attrib.get("automation_sequence_scenarios_id", "")
        comment = elem.attrib.get("comment_field", "")

        steps = extract_step_records(elem)
        scenarios.append({
            "title": title,
            "scenario_id": sid,
            "automation_sequence_scenarios_id": auto_scenario_id,
            "path": path,
            "comment_field": comment,
            "attributes": dict(elem.attrib),
            "steps": steps,
        })

    deduped = []
    seen = set()
    for s in scenarios:
        key = (s["title"], s["scenario_id"], s["path"], len(s["steps"]))
        if key not in seen:
            seen.add(key)
            deduped.append(s)
    return deduped


def detect_source_guide(root: ET.Element, src: Path) -> str:
    """
    Determine the best human-readable source guide or script name for the XML document.
    """
    auto_seq = extract_automation_sequence(root)
    if auto_seq is not None:
        name = auto_seq.attrib.get("name", "").strip()
        if name:
            return name

    for names in [
        ["sourceGuide", "source_guide", "guideName", "guide"],
        ["automationSequenceName", "sequenceName", "scriptName"],
        ["name", "title"],
    ]:
        val = first_text_anywhere(root, names)
        if val:
            return val
    return src.stem


def extract_execution_summary(root: ET.Element, src: Path, scenario_count: int) -> Dict[str, str]:
    """
    Extract execution-level metadata, including Test Case ID from automation_sequence id.
    """
    auto_seq = extract_automation_sequence(root)

    if auto_seq is not None:
        summary = {
            "Source guide": auto_seq.attrib.get("name", "") or detect_source_guide(root, src),
            "Test Case ID": auto_seq.attrib.get("id", ""),
            "Automation Sequence ID": auto_seq.attrib.get("id", ""),
            "Automation Sequence Name": auto_seq.attrib.get("name", ""),
            "Description": auto_seq.attrib.get("description", ""),
            "Run ID": auto_seq.attrib.get("run_id", "") or src.stem,
            "Status": auto_seq.attrib.get("status_name", "") or auto_seq.attrib.get("status", ""),
            "Status Code": auto_seq.attrib.get("status", ""),
            "Execution Time": auto_seq.attrib.get("execution_time", ""),
            "Execution Sent Date Time": auto_seq.attrib.get("executionsentdatetime", ""),
            "Execution Start Time": auto_seq.attrib.get("executionstarttime", ""),
            "Execution End Time": auto_seq.attrib.get("executionendtime", ""),
            "Machine": auto_seq.attrib.get("machine", ""),
            "Machine ID": auto_seq.attrib.get("machine_id", ""),
            "Executed By": auto_seq.attrib.get("executed_by", ""),
            "Executed By ID": auto_seq.attrib.get("executed_by_id", ""),
            "Last Modified By ID": auto_seq.attrib.get("last_modified_by_id", ""),
            "Client ID": auto_seq.attrib.get("client_id", ""),
            "Iterations": auto_seq.attrib.get("iterations", ""),
            "Verdict": auto_seq.attrib.get("verdict", ""),
            "Reason": auto_seq.attrib.get("reason", ""),
            "Scenario Count": str(scenario_count),
        }
        for key in ["Execution Time", "Execution Sent Date Time", "Execution Start Time", "Execution End Time"]:
            if summary.get(key):
                summary[key] = parse_datetime_text(summary[key])
        return summary

    summary = {
        "Source guide": detect_source_guide(root, src),
        "Test Case ID": "",
        "Automation Sequence ID": "",
        "Automation Sequence Name": "",
        "Description": "",
        "Run ID": first_text_anywhere(root, ["runId", "run_id", "executionId", "execution_id", "id"]) or src.stem,
        "Status": first_text_anywhere(root, ["status", "executionStatus", "runStatus"]),
        "Status Code": "",
        "Execution Time": first_text_anywhere(root, ["executionTime", "runTime", "startTime", "executionStartTime"]),
        "Execution Sent Date Time": "",
        "Execution Start Time": "",
        "Execution End Time": "",
        "Machine": first_text_anywhere(root, ["machine", "machineName", "host", "computerName"]),
        "Machine ID": "",
        "Executed By": first_text_anywhere(root, ["executedBy", "user", "username", "runBy"]),
        "Executed By ID": "",
        "Last Modified By ID": "",
        "Client ID": "",
        "Iterations": "",
        "Verdict": "",
        "Reason": "",
        "Scenario Count": str(scenario_count),
    }
    if summary["Execution Time"]:
        summary["Execution Time"] = parse_datetime_text(summary["Execution Time"])
    return summary



# -----------------------------------------------------------------------------
# Full XML inventory support for completeness / NotebookLM ingestion
# -----------------------------------------------------------------------------
def xml_inventory_rows(root: ET.Element) -> List[Dict[str, str]]:
    """
    Flatten the entire XML tree into rows containing node path, tag name, depth, attributes, and text.
    """
    rows: List[Dict[str, str]] = []

    def visit(elem: ET.Element, path: str, depth: int):
        tag_name = local_name(elem.tag)
        current_path = f"{path}/{tag_name}" if path else tag_name
        # Changed text to only pull immediate node content to avoid unreadable concatenated strings
        text = (elem.text or "").strip()
        rows.append({
            "Path": current_path,
            "Node": tag_name,
            "Depth": str(depth),
            "Attributes": "; ".join(f"{k}={v}" for k, v in elem.attrib.items()),
            "Text": text[:4000],
        })
        for child in list(elem):
            visit(child, current_path, depth + 1)

    visit(root, "", 0)
    return rows


def xml_inventory_markdown(root: ET.Element) -> List[str]:
    """
    Render the flattened XML inventory rows as a Markdown table.
    """
    rows = xml_inventory_rows(root)
    lines = [
        "XML Inventory",
        "| Path | Node | Depth | Attributes | Text |",
        "| --- | --- | --- | --- | --- |",
    ]
    for row in rows:
        lines.append(
            f"| {markdown_escape(row['Path'], 3000)} | "
            f"{markdown_escape(row['Node'], 500)} | "
            f"{markdown_escape(row['Depth'], 50)} | "
            f"{markdown_escape(row['Attributes'], 8000)} | "
            f"{markdown_escape(row['Text'], 4000)} |"
        )
    lines.append("")
    return lines


def scenario_attributes_table_markdown(attributes: Dict[str, str]) -> List[str]:
    """
    Render scenario or automation-sequence attributes as a two-column Markdown table.
    """
    lines = [
        "| Attribute | Value |",
        "| --- | --- |",
    ]
    for key, value in attributes.items():
        lines.append(f"| {markdown_escape(key, 500)} | {markdown_escape(value, 3000)} |")
    lines.append("")
    return lines


def steps_table_markdown(steps: List[Dict[str, str]]) -> List[str]:
    """
    Render extracted TestSavvy step records as a detailed Markdown table.
    """
    header = [
        "#", "Step ID", "Action", "Type", "English Text", "Dataset / Header",
        "Test Condition", "Encrypted", "Value", "Application", "Interface Map",
        "Logical Name", "Interface Element Type", "Tool", "Comments", "Action Code",
        "Raw Attributes"
    ]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]
    for step in steps:
        row = [markdown_escape(step.get(col, ""), 8000) for col in header]
        lines.append("| " + " | ".join(row) + " |")
    if not steps:
        lines.append("| " + " | ".join([""] * len(header)) + " |")
    return lines



# -----------------------------------------------------------------------------
# Enhanced TestSavvy Markdown report generation
# -----------------------------------------------------------------------------
def testsavvy_execution_report_markdown(src: Path, root: ET.Element, include_xml_inventory: bool = True) -> str:
    """
    Build the enhanced NotebookLM-ready TestSavvy Markdown report from parsed XML.
    """
    scenarios = extract_scenarios(root)
    runtime_vars = extract_runtime_variables(root)
    summary = extract_execution_summary(root, src, len(scenarios))
    auto_seq = extract_automation_sequence(root)

    lines: List[str] = []
    lines.append(f"# {markdown_escape(src.name, 500)}")
    lines.append("")
    lines.append("Source guide")
    lines.append(f"{markdown_escape(summary['Source guide'], 1000)}")
    lines.append("")

    lines.append("Execution Summary")
    if summary["Test Case ID"]:
        lines.append(f"Test Case ID: {markdown_escape(summary['Test Case ID'], 1000)}")
    if summary["Automation Sequence ID"]:
        lines.append(f"Automation Sequence ID: {markdown_escape(summary['Automation Sequence ID'], 1000)}")
    if summary["Automation Sequence Name"]:
        lines.append(f"Automation Sequence Name: {markdown_escape(summary['Automation Sequence Name'], 1000)}")
    if summary["Description"]:
        lines.append(f"Description: {markdown_escape(summary['Description'], 1000)}")
    lines.append(f"Run ID: {markdown_escape(summary['Run ID'], 1000)}")
    if summary["Status"]:
        lines.append(f"Status: {markdown_escape(summary['Status'], 1000)}")
    if summary["Status Code"]:
        lines.append(f"Status Code: {markdown_escape(summary['Status Code'], 1000)}")
    if summary["Execution Time"]:
        lines.append(f"Execution Time: {markdown_escape(summary['Execution Time'], 1000)}")
    if summary["Execution Sent Date Time"]:
        lines.append(f"Execution Sent Date Time: {markdown_escape(summary['Execution Sent Date Time'], 1000)}")
    if summary["Execution Start Time"]:
        lines.append(f"Execution Start Time: {markdown_escape(summary['Execution Start Time'], 1000)}")
    if summary["Execution End Time"]:
        lines.append(f"Execution End Time: {markdown_escape(summary['Execution End Time'], 1000)}")
    if summary["Machine"]:
        lines.append(f"Machine: {markdown_escape(summary['Machine'], 1000)}")
    if summary["Machine ID"]:
        lines.append(f"Machine ID: {markdown_escape(summary['Machine ID'], 1000)}")
    if summary["Executed By"]:
        lines.append(f"Executed By: {markdown_escape(summary['Executed By'], 1000)}")
    if summary["Executed By ID"]:
        lines.append(f"Executed By ID: {markdown_escape(summary['Executed By ID'], 1000)}")
    if summary["Last Modified By ID"]:
        lines.append(f"Last Modified By ID: {markdown_escape(summary['Last Modified By ID'], 1000)}")
    if summary["Client ID"]:
        lines.append(f"Client ID: {markdown_escape(summary['Client ID'], 1000)}")
    if summary["Iterations"]:
        lines.append(f"Iterations: {markdown_escape(summary['Iterations'], 1000)}")
    if summary["Verdict"]:
        lines.append(f"Verdict: {markdown_escape(summary['Verdict'], 1000)}")
    if summary["Reason"]:
        lines.append(f"Reason: {markdown_escape(summary['Reason'], 1000)}")
    lines.append(f"Scenario Count: {markdown_escape(summary['Scenario Count'], 1000)}")
    lines.append("")

    if auto_seq is not None:
        lines.append("Automation Sequence Attributes")
        lines.extend(scenario_attributes_table_markdown(dict(auto_seq.attrib)))

        machine_settings = None
        for child in list(auto_seq):
            if local_name(child.tag).lower() == "machine_settings":
                machine_settings = child
                break
        if machine_settings is not None:
            lines.append("Machine Settings Attributes")
            lines.extend(scenario_attributes_table_markdown(dict(machine_settings.attrib)))

    lines.append("Runtime Variables")
    if runtime_vars:
        for key, val in runtime_vars:
            lines.append(f"{markdown_escape(key, 500)}: {markdown_escape(val, 1000)}")
    else:
        lines.append("None")
    lines.append("")

    lines.append("Scenarios and Steps")
    if not scenarios:
        lines.append("No scenarios were detected.")
        lines.append("")
    else:
        for idx, scenario in enumerate(scenarios, start=1):
            lines.append(f"{idx}. {markdown_escape(scenario['title'], 1000)}")
            if scenario.get("scenario_id"):
                lines.append(f"Scenario ID: {markdown_escape(scenario['scenario_id'], 1000)}")
            if scenario.get("automation_sequence_scenarios_id"):
                lines.append(f"Automation Sequence Scenario ID: {markdown_escape(scenario['automation_sequence_scenarios_id'], 1000)}")
            if scenario.get("path"):
                lines.append(f"Path: {markdown_escape(scenario['path'], 1000)}")
            if scenario.get("comment_field"):
                lines.append(f"Comment: {markdown_escape(scenario['comment_field'], 1000)}")
            lines.append("Scenario Attributes")
            lines.extend(scenario_attributes_table_markdown(scenario.get("attributes", {})))
            lines.extend(steps_table_markdown(scenario.get("steps", [])))
            lines.append("")

    if include_xml_inventory:
        lines.extend(xml_inventory_markdown(root))

    return "\n".join(lines)


def xml_to_markdown(src: Path, cfg: AppConfig) -> str:
    """
    Dispatch XML conversion to either raw Markdown mode or enhanced TestSavvy report mode.
    """
    tree = ET.parse(str(src), parser=ET.XMLParser())
    root = tree.getroot()
    if cfg.xml_output_mode == "raw_markdown":
        xml_text = pretty_print_xml(src, cfg.indent, cfg.xml_encoding, cfg.preserve_xml_declaration).decode(cfg.xml_encoding, errors="replace")
        return wrap_xml_markdown(src, xml_text)
    return testsavvy_execution_report_markdown(src, root, cfg.include_xml_inventory)


def process_standard_file(src: Path, out_path: Path, cfg: AppConfig):
    """
    Process JSON or XML source files and write the configured output artifact.
    """
    if out_path.exists() and not cfg.overwrite:
        return True, f"SKIP (exists): {src}"

    try:
        ensure_parent_dir(out_path)
        ext = src.suffix.lower().lstrip(".")

        if ext == "json":
            pretty = pretty_print_json(src, cfg.indent, cfg.json_sort_keys)
            if (cfg.output_extension or "").lower() == ".md":
                pretty = wrap_json_markdown(src, pretty)
            out_path.write_text(pretty, encoding="utf-8")

        elif ext == "xml":
            if (cfg.output_extension or "").lower() == ".md":
                md = xml_to_markdown(src, cfg)
                out_path.write_text(md, encoding="utf-8")
            else:
                out_path.write_bytes(pretty_print_xml(src, cfg.indent, cfg.xml_encoding, cfg.preserve_xml_declaration))

        else:
            return True, f"SKIP (unsupported ext): {src}"

        return True, f"OK: {src} -> {out_path}"
    except Exception as e:
        return False, f"FAIL: {src} ({type(e).__name__}: {e})"


def build_combined_section_for_file(src: Path, cfg: AppConfig) -> str:
    """
    Build one Markdown section for an XML or JSON file when combined output is enabled.
    """
    ext = src.suffix.lower().lstrip(".")
    if ext == "xml":
        body = xml_to_markdown(src, cfg).strip()
        return body
    if ext in {"json"}:
        body = wrap_json_markdown(src, pretty_print_json(src, cfg.indent, cfg.json_sort_keys)).strip()
        return body
    raise ValueError(f"Unsupported combined mode source: {src}")


def build_excel_combined_section(src: Path, sheet_name: str, content: str) -> str:
    """
    Build one Markdown section for an Excel worksheet when combined output is enabled.
    """
    return f"# {src.name} - {sheet_name}\n\n{content.strip()}\n"





def key_value_table_markdown(title: str, values: Dict[str, str], max_value_len: int = 3000) -> List[str]:
    """
    Render a named key/value dictionary as a Markdown table with an explicit heading.
    NotebookLM generally understands named tables better than unlabeled colon-only blocks.
    """
    lines = [title, "", "| Field | Value |", "| --- | --- |"]
    for key, value in values.items():
        if value not in (None, ""):
            lines.append(f"| {markdown_escape(key, 500)} | {markdown_escape(value, max_value_len)} |")
    lines.append("")
    return lines

# -----------------------------------------------------------------------------
# TestSavvy JSON support
# -----------------------------------------------------------------------------
def safe_filename_component(value: object, max_len: int = 120) -> str:
    """
    Convert arbitrary text into a safe filename component that keeps underscores
    and avoids spaces/special characters that can confuse ingestion tools.
    """
    text = "" if value is None else str(value)
    text = re.sub(r"[^\w\-]+", "_", text.strip())
    text = re.sub(r"_+", "_", text).strip("_")
    if not text:
        text = "unnamed"
    return text[:max_len].strip("_") or "unnamed"


def json_type_name(value: Any) -> str:
    """
    Return a compact JSON type name for inventory rows.
    """
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return "array"
    if isinstance(value, int) and not isinstance(value, bool):
        return "integer"
    if isinstance(value, float):
        return "number"
    return "string"


def compact_json_value(value: Any, max_len: int = 1000) -> str:
    """
    Render a JSON scalar or compact collection summary for Markdown tables.
    """
    if isinstance(value, dict):
        return f"object with {len(value)} keys"
    if isinstance(value, list):
        return f"array with {len(value)} items"
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return text[:max_len - 3] + "..." if max_len > 0 and len(text) > max_len else text


def json_inventory_rows(value: Any, path: str = "$", rows: Optional[List[Dict[str, str]]] = None, max_rows: int = 5000) -> List[Dict[str, str]]:
    """
    Flatten JSON into path/type/summary rows so NotebookLM can locate deeply nested values.
    """
    if rows is None:
        rows = []
    if len(rows) >= max_rows:
        return rows

    if isinstance(value, dict):
        rows.append({
            "Path": path,
            "Type": "object",
            "Summary": f"{len(value)} keys",
            "Value": "",
        })
        for key, child in value.items():
            child_path = f"{path}.{key}" if re.match(r"^[A-Za-z_]\w*$", str(key)) else f"{path}[{json.dumps(str(key))}]"
            json_inventory_rows(child, child_path, rows, max_rows)
            if len(rows) >= max_rows:
                break
    elif isinstance(value, list):
        rows.append({
            "Path": path,
            "Type": "array",
            "Summary": f"{len(value)} items",
            "Value": "",
        })
        for idx, child in enumerate(value):
            json_inventory_rows(child, f"{path}[{idx}]", rows, max_rows)
            if len(rows) >= max_rows:
                break
    else:
        rows.append({
            "Path": path,
            "Type": json_type_name(value),
            "Summary": "",
            "Value": compact_json_value(value, 2000),
        })
    return rows


def json_inventory_markdown(data: Any, max_rows: int = 5000) -> List[str]:
    """
    Render a flattened JSON inventory as Markdown.
    """
    rows = json_inventory_rows(data, max_rows=max_rows)
    lines = [
        "## JSON Inventory",
        "",
        "This flattened inventory is the fallback source of truth for deeply nested JSON values not shown in the semantic sections above.",
        "",
        "| JSON Path | Type | Summary | Value |",
        "| --- | --- | --- | --- |",
    ]
    for row in rows:
        lines.append(
            f"| {markdown_escape(row['Path'], 2000)} | {markdown_escape(row['Type'], 100)} | "
            f"{markdown_escape(row['Summary'], 500)} | {markdown_escape(row['Value'], 4000)} |"
        )
    if len(rows) >= max_rows:
        lines.append(f"| INVENTORY_TRUNCATED | notice | Reached configured json_inventory_max_rows={max_rows} | Increase json_inventory_max_rows to include more rows. |")
    lines.append("")
    return lines


def detect_testsavvy_json_kind(data: Any) -> str:
    """
    Detect known TestSavvy JSON export shapes.
    """
    if isinstance(data, dict) and isinstance(data.get("automationSequence"), dict) and isinstance(data.get("sequenceList"), list):
        return "test_case"
    if isinstance(data, dict) and isinstance(data.get("export"), dict) and "interfaceMapName" in data.get("export", {}):
        return "interface_repository"
    if isinstance(data, list):
        return "test_case_group"
    return "generic"


def testsavvy_json_output_filename(data: Any, source_name: str, fallback_index: Optional[int] = None) -> str:
    """
    Build NotebookLM-friendly Markdown filenames for JSON exports.
    """
    kind = detect_testsavvy_json_kind(data)
    suffix = "" if fallback_index is None else f"_{fallback_index:04d}"

    if kind == "test_case":
        auto = data.get("automationSequence", {})
        tc_id = auto.get("id") or "unknown"
        return f"TC_{safe_filename_component(tc_id)}{suffix}.md"

    if kind == "interface_repository":
        export = data.get("export", {})
        name = export.get("interfaceMapName") or export.get("interfaceElementName") or "interface"
        return f"IR_{safe_filename_component(name)}{suffix}.md"

    if kind == "test_case_group":
        stem = safe_filename_component(Path(source_name).stem, 80)
        return f"GROUP_{stem}{suffix}.md"

    stem = safe_filename_component(Path(source_name).stem, 120)
    if "_" not in stem:
        stem = f"JSON_{stem}"
    return f"{stem}{suffix}.md"


# -----------------------------------------------------------------------------
# TestSavvy JSON execution-order helpers
# -----------------------------------------------------------------------------
def coerce_sequence(value: Any, fallback: int = 999999) -> int:
    """
    Convert TestSavvy sequence-like values to integers for stable ordering.
    Missing, blank, or non-numeric values sort after real sequence numbers.
    """
    if value is None or value == "":
        return fallback
    try:
        return int(value)
    except Exception:
        try:
            return int(float(str(value)))
        except Exception:
            return fallback


def flag_is_not_false(value: Any) -> bool:
    """
    Treat missing/null as active unless TestSavvy explicitly exports false/0/no.
    This prevents reusable scenarios with omitted flags from being incorrectly marked excluded.
    """
    if value is None:
        return True
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"false", "0", "no", "n", "excluded", "inactive"}:
        return False
    return True


def flag_label(value: Any) -> str:
    """
    Render a TestSavvy status/include flag in human-friendly form.
    """
    return "Yes" if flag_is_not_false(value) else "No"


def scenario_sequence_value(scenario: Dict[str, Any], fallback: int) -> int:
    """
    Return the scenario sequence TestSavvy uses for authored/runtime order.
    `sequence` is preferred; `original_sequence` is the fallback.
    """
    if scenario.get("sequence") not in (None, ""):
        return coerce_sequence(scenario.get("sequence"), fallback)
    return coerce_sequence(scenario.get("original_sequence"), fallback)


def step_sequence_value(step: Dict[str, Any], fallback: int) -> int:
    """
    Return the step sequence inside a scenario.
    """
    return coerce_sequence(step.get("sequence"), fallback)


def scenario_is_executable(scenario: Dict[str, Any]) -> bool:
    """
    A JSON scenario is runtime-equivalent only when status and include_status are not false.
    """
    return flag_is_not_false(scenario.get("status")) and flag_is_not_false(scenario.get("include_status"))


def step_is_executable(step: Dict[str, Any]) -> bool:
    """
    A JSON step is runtime-equivalent only when step status/include_status and stepAssociation.status are not false.
    """
    assoc = step.get("stepAssociation") or {}
    return (
        flag_is_not_false(step.get("status"))
        and flag_is_not_false(step.get("include_status"))
        and flag_is_not_false(assoc.get("status"))
    )


def scenario_exclusion_reason(scenario: Dict[str, Any]) -> str:
    """
    Explain why a scenario is not expected to execute.
    """
    reasons: List[str] = []
    if not flag_is_not_false(scenario.get("status")):
        reasons.append("scenario.status=false")
    if not flag_is_not_false(scenario.get("include_status")):
        reasons.append("scenario.include_status=false")
    return "; ".join(reasons)


def step_exclusion_reason(step: Dict[str, Any], scenario: Optional[Dict[str, Any]] = None) -> str:
    """
    Explain why a step is not expected to execute.
    """
    reasons: List[str] = []
    if scenario is not None:
        scen_reason = scenario_exclusion_reason(scenario)
        if scen_reason:
            reasons.append(scen_reason)
    if not flag_is_not_false(step.get("status")):
        reasons.append("step.status=false")
    if not flag_is_not_false(step.get("include_status")):
        reasons.append("step.include_status=false")
    assoc = step.get("stepAssociation") or {}
    if not flag_is_not_false(assoc.get("status")):
        reasons.append("stepAssociation.status=false")
    return "; ".join(reasons)


def sorted_scenario_records(scenarios: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Return scenario records with raw JSON index, runtime sequence, and execution classification.
    The raw JSON array order is NOT treated as execution order.
    """
    records: List[Dict[str, Any]] = []
    for raw_idx, scenario in enumerate(scenarios):
        seq = scenario_sequence_value(scenario, raw_idx + 1)
        active = scenario_is_executable(scenario)
        records.append({
            "raw_index": raw_idx,
            "runtime_sequence": seq,
            "scenario": scenario,
            "active": active,
            "exclusion_reason": scenario_exclusion_reason(scenario),
        })
    return sorted(records, key=lambda r: (r["runtime_sequence"], r["raw_index"]))


def sorted_steps_for_scenario(scenario: Dict[str, Any]) -> List[Tuple[int, Dict[str, Any]]]:
    """
    Sort a scenario's steps by TestSavvy step sequence rather than JSON array order.
    """
    steps = scenario.get("scenarioFlowList") or []
    indexed = []
    for raw_idx, step in enumerate(steps):
        indexed.append((raw_idx, step_sequence_value(step, raw_idx), step))
    indexed.sort(key=lambda item: (item[1], item[0]))
    return [(raw_idx, step) for raw_idx, _seq, step in indexed]


def format_json_location(raw_scenario_index: int, raw_step_index: Optional[int] = None) -> str:
    """
    Build a JSON path back to the raw exported file.
    """
    base = f"$.sequenceList[{raw_scenario_index}]"
    if raw_step_index is None:
        return base
    return f"{base}.scenarioFlowList[{raw_step_index}]"


def testsavvy_testcase_step_rows(
    scenario: Dict[str, Any],
    scenario_active: bool = True,
    raw_scenario_index: int = 0,
) -> List[Dict[str, str]]:
    """
    Extract step rows from a TestSavvy test-case JSON scenarioFlowList.
    Steps are sorted by step.sequence and include active/excluded execution status.
    """
    rows: List[Dict[str, str]] = []
    for display_idx, (raw_step_idx, step) in enumerate(sorted_steps_for_scenario(scenario), start=1):
        assoc = step.get("stepAssociation") or {}
        is_active = scenario_active and step_is_executable(step)
        reason = "" if is_active else step_exclusion_reason(step, scenario)
        rows.append({
            "Runtime Step Order": str(display_idx),
            "Step Sequence": str(step_sequence_value(step, raw_step_idx)),
            "Execution Status": "Active / Executable" if is_active else "Excluded / Not Executed",
            "Exclusion Reason": reason,
            "Step Name": step.get("name", ""),
            "Action": step.get("action_name", ""),
            "Type": step.get("action_type", ""),
            "Classification": step.get("action_classification", ""),
            "English Text": step.get("english_text", ""),
            "Element Logical Name": step.get("element_logical_name", ""),
            "Element Type": step.get("element_type", ""),
            "Dataset": assoc.get("dataset_name", ""),
            "Dataset Header": step.get("dataset_header_name", "") or assoc.get("dataset_header_name", ""),
            "Value": assoc.get("value", ""),
            "Encrypted": "Yes" if assoc.get("encrypt_data") else "No",
            "Tool": step.get("auto_profile_name", "") or step.get("test_engine_name", ""),
            "Requires Data": "Yes" if step.get("requires_data") else "No",
            "Uses Override": "Yes" if step.get("uses_override") or assoc.get("uses_override") else "No",
            "SQL Command": step.get("sql_command", ""),
            "Action Code": step.get("action_code", ""),
            "JSON Location": format_json_location(raw_scenario_index, raw_step_idx),
            "Raw Step JSON": json.dumps(step, ensure_ascii=False, sort_keys=True)[:8000],
        })
    return rows


def render_scenario_index_table(title: str, records: List[Dict[str, Any]], note: str = "") -> List[str]:
    """
    Render an execution-aware scenario index.
    """
    lines = [title, ""]
    if note:
        lines.extend([note, ""])
    lines.append("| Runtime Order | Scenario Sequence | Scenario Name | Logical Name | Scenario Type | Reusable | Status | Include Status | Execution Status | Exclusion Reason | Comment | Step Count | JSON Location |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for runtime_order, record in enumerate(records, start=1):
        scenario = record["scenario"]
        active = record["active"]
        lines.append(
            f"| {runtime_order} | {record['runtime_sequence']} | "
            f"{markdown_escape(scenario.get('name', ''), 1000)} | "
            f"{markdown_escape(scenario.get('logical_name', ''), 1000)} | "
            f"{markdown_escape(scenario.get('scenario_type', ''), 500)} | "
            f"{flag_label(scenario.get('reusable'))} | "
            f"{flag_label(scenario.get('status'))} | "
            f"{flag_label(scenario.get('include_status'))} | "
            f"{'Active / Executable' if active else 'Excluded / Not Executed'} | "
            f"{markdown_escape(record.get('exclusion_reason', ''), 1500)} | "
            f"{markdown_escape(scenario.get('comment_field', ''), 1000)} | "
            f"{len(scenario.get('scenarioFlowList') or [])} | `{format_json_location(record['raw_index'])}` |"
        )
    if not records:
        lines.append("|  |  |  |  |  |  |  |  |  |  |  | 0 |  |")
    lines.append("")
    return lines


def submit_like_step(step: Dict[str, Any]) -> bool:
    """
    Return True for actual Submit button/action references, not verification text such as
    "Transaction submitted successfully".
    """
    action_name = str(step.get("action_name") or "").strip().lower()
    step_name = str(step.get("name") or "").strip().lower()
    logical_name = str(step.get("element_logical_name") or "").strip().lower()
    element_type = str(step.get("element_type") or "").strip().lower()
    action_code = str(step.get("action_code") or "").lower()

    if "viewactions.submit" in action_code:
        return True
    if action_name == "click" and logical_name == "submit":
        return True
    if action_name == "click" and step_name.endswith("~ submit"):
        return True
    if logical_name == "submit" and element_type in {"button", "menu item", "link"}:
        return True
    return False


def render_submit_execution_index(records: List[Dict[str, Any]]) -> List[str]:
    """
    Build active/excluded Submit index so NotebookLM can answer coverage questions correctly.
    """
    lines = [
        "## Submit Button Execution Index",
        "",
        "Use this section for questions such as `Which TestSavvy automated test cases actively click Submit?`. Count only rows where Execution Status is `Active / Executable` unless the question explicitly asks for excluded/inactive steps.",
        "",
        "| Runtime Scenario Order | Scenario Sequence | Step Sequence | Execution Status | Exclusion Reason | Scenario Name | Step Name | Element Logical Name | Element Type | Action Code | JSON Location |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    found = 0
    for runtime_order, record in enumerate(records, start=1):
        scenario = record["scenario"]
        scenario_active = record["active"]
        for raw_step_idx, step in sorted_steps_for_scenario(scenario):
            if not submit_like_step(step):
                continue
            found += 1
            active = scenario_active and step_is_executable(step)
            lines.append(
                f"| {runtime_order} | {record['runtime_sequence']} | {step_sequence_value(step, raw_step_idx)} | "
                f"{'Active / Executable' if active else 'Excluded / Not Executed'} | "
                f"{markdown_escape('' if active else step_exclusion_reason(step, scenario), 1500)} | "
                f"{markdown_escape(scenario.get('name', ''), 1000)} | "
                f"{markdown_escape(step.get('name', ''), 1000)} | "
                f"{markdown_escape(step.get('element_logical_name', ''), 1000)} | "
                f"{markdown_escape(step.get('element_type', ''), 500)} | "
                f"{markdown_escape(step.get('action_code', ''), 5000)} | `{format_json_location(record['raw_index'], raw_step_idx)}` |"
            )
    if not found:
        lines.append("|  |  |  |  |  |  |  |  |  |  |  |")
    lines.append("")
    return lines


def testsavvy_testcase_json_markdown(source_name: str, data: Dict[str, Any], cfg: AppConfig) -> str:
    """
    Convert a TestSavvy test-case JSON export into NotebookLM-oriented Markdown.
    IMPORTANT: sequenceList array order is not execution order. Scenarios are sorted by sequence/original_sequence.
    """
    auto = data.get("automationSequence", {}) or {}
    scenarios = data.get("sequenceList", []) or []
    records = sorted_scenario_records(scenarios)
    active_records = [r for r in records if r["active"]]
    excluded_records = [r for r in records if not r["active"]]

    lines: List[str] = []
    title_id = auto.get("id", "")
    title_name = auto.get("name", Path(source_name).stem)

    lines.append(f"# TestSavvy Test Case JSON Payload: {markdown_escape(title_id, 500)} - {markdown_escape(title_name, 1000)}")
    lines.append("")
    lines.append("## Document Identity")
    lines.append("")
    lines.append("| Field | Value |")
    lines.append("| --- | --- |")
    lines.append(f"| Source File | {markdown_escape(source_name, 500)} |")
    lines.append(f"| JSON Export Type | TestSavvy Test Case Export |")
    lines.append(f"| Test Case ID | {markdown_escape(auto.get('id', ''), 500)} |")
    lines.append(f"| Test Case Name | {markdown_escape(auto.get('name', ''), 1000)} |")
    lines.append(f"| Raw JSON Scenario Count | {len(scenarios)} |")
    lines.append(f"| Active Runtime-Equivalent Scenario Count | {len(active_records)} |")
    lines.append(f"| Excluded / Inactive Scenario Count | {len(excluded_records)} |")
    lines.append("")

    lines.append("## Critical Ordering Rule")
    lines.append("")
    lines.append("The exported JSON `sequenceList[]` array order is not reliable execution order. This Markdown file sorts scenarios by `scenario.sequence`, falling back to `scenario.original_sequence`, and preserves raw JSON location separately. Use `Active Runtime-Equivalent Scenario Index` for what the test is expected to execute. Use `Full Raw JSON Scenario Inventory` only for debugging the export file.")
    lines.append("")

    summary_keys = [
        "id", "name", "description", "created_by_id", "last_modified_by_id", "project_id",
        "automation_sequence_test_case_status_id", "case_group_ind", "debug", "continue_on_fail",
        "contains_holds", "displayStepNum", "deleted"
    ]
    lines.append("## Test Case Summary")
    lines.append("")
    lines.append("| Field | Value | JSON Location |")
    lines.append("| --- | --- | --- |")
    for key in summary_keys:
        if key in auto:
            lines.append(f"| {markdown_escape(key, 500)} | {markdown_escape(compact_json_value(auto.get(key), 3000), 3000)} | `$.automationSequence.{markdown_escape(key, 300)}` |")
    lines.append("")

    lines.append("## NotebookLM JSON Field Map")
    lines.append("")
    lines.append("| Business Term | JSON Location | Markdown Location | Meaning |")
    lines.append("| --- | --- | --- | --- |")
    field_rows = [
        ("Test Case ID", "$.automationSequence.id", "## Document Identity; ## Test Case Summary", "The TestSavvy test case / automation sequence identifier."),
        ("Test Case Name", "$.automationSequence.name", "Document title; ## Document Identity; ## Test Case Summary", "Human-readable name of the TestSavvy test case."),
        ("Raw Scenario List", "$.sequenceList[]", "## Full Raw JSON Scenario Inventory", "The exported JSON array. This is not necessarily execution order."),
        ("Execution Scenario Order", "$.sequenceList[].sequence", "## Active Runtime-Equivalent Scenario Index", "Primary scenario ordering field for the authored/runtime-equivalent flow."),
        ("Fallback Scenario Order", "$.sequenceList[].original_sequence", "## Active Runtime-Equivalent Scenario Index", "Fallback ordering field when sequence is missing."),
        ("Scenario Include Status", "$.sequenceList[].include_status", "All scenario indexes", "False means the scenario exists in JSON but should not be counted as executed coverage."),
        ("Scenario Status", "$.sequenceList[].status", "All scenario indexes", "False means the scenario is inactive."),
        ("Scenario Steps", "$.sequenceList[].scenarioFlowList[]", "Scenario step detail tables", "Step/action definitions inside each scenario, sorted by step.sequence."),
        ("Step Include Status", "$.sequenceList[].scenarioFlowList[].include_status", "Scenario step detail tables column Execution Status", "False means the step exists in JSON but is excluded."),
        ("Step Dataset Value", "$.sequenceList[].scenarioFlowList[].stepAssociation.value", "Scenario step detail tables column Value", "Data entered, selected, clicked, or used by the step."),
        ("Step Action Code", "$.sequenceList[].scenarioFlowList[].action_code", "Scenario step detail tables column Action Code", "Executable WebDriver/Selenium/custom automation code."),
    ]
    for term, jloc, mdloc, meaning in field_rows:
        lines.append(f"| {term} | `{jloc}` | {mdloc} | {meaning} |")
    lines.append("")

    lines.extend(render_scenario_index_table(
        "## Active Runtime-Equivalent Scenario Index",
        active_records,
        "This is the primary flow NotebookLM should use for questions about what the test does. It is sorted by `sequence`/`original_sequence` and excludes inactive scenarios.",
    ))
    lines.extend(render_scenario_index_table(
        "## Excluded / Inactive Scenario Index",
        excluded_records,
        "These scenarios exist in the JSON export but should not be counted as executed coverage unless the user explicitly asks about excluded/inactive scenarios.",
    ))

    lines.append("## Full Raw JSON Scenario Inventory")
    lines.append("")
    lines.append("This section preserves the physical JSON array order for debugging only. Do not use it as execution order.")
    lines.append("")
    lines.append("| Raw JSON Array Position | Scenario Sequence | Scenario Name | Include Status | Status | Comment | JSON Location |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- |")
    for raw_idx, scenario in enumerate(scenarios):
        lines.append(
            f"| {raw_idx + 1} | {scenario_sequence_value(scenario, raw_idx + 1)} | "
            f"{markdown_escape(scenario.get('name', ''), 1000)} | "
            f"{flag_label(scenario.get('include_status'))} | "
            f"{flag_label(scenario.get('status'))} | "
            f"{markdown_escape(scenario.get('comment_field', ''), 1000)} | `{format_json_location(raw_idx)}` |"
        )
    lines.append("")

    lines.extend(render_submit_execution_index(records))

    custom_functions = []
    selector_rows = []
    for runtime_order, record in enumerate(records, start=1):
        scenario = record["scenario"]
        scenario_active = record["active"]
        for raw_step_idx, step in sorted_steps_for_scenario(scenario):
            assoc = step.get("stepAssociation") or {}
            value = str(assoc.get("value") or "")
            action_code = str(step.get("action_code") or "")
            active = scenario_active and step_is_executable(step)
            status_text = "Active / Executable" if active else "Excluded / Not Executed"
            if "@" in value or "@" in action_code or str(step.get("action_name", "")).lower() == "custom function":
                custom_functions.append((runtime_order, record["runtime_sequence"], status_text, scenario.get("name", ""), step.get("name", ""), value, action_code, format_json_location(record["raw_index"], raw_step_idx)))
            if any(token in action_code for token in ["xpath", "data-qa", "aria-label", "find_element", "contains(@"]):
                selector_rows.append((runtime_order, record["runtime_sequence"], status_text, scenario.get("name", ""), step.get("name", ""), step.get("element_logical_name", ""), step.get("element_type", ""), action_code, format_json_location(record["raw_index"], raw_step_idx)))

    lines.append("## Custom Functions and Dynamic Values")
    lines.append("")
    lines.append("| Runtime Scenario Order | Scenario Sequence | Execution Status | Scenario | Step | Value | Action Code | JSON Location |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- |")
    for row in custom_functions[:cfg.json_table_max_rows]:
        *cells, json_loc = row
        lines.append("| " + " | ".join(markdown_escape(x, 4000) for x in cells) + f" | `{json_loc}` |")
    if not custom_functions:
        lines.append("|  |  |  |  |  |  |  |  |")
    lines.append("")

    lines.append("## Selector / Action Code Index")
    lines.append("")
    lines.append("Use this table to review XPath, `data-qa`, `data-qa-id`, `aria-label`, and other locator logic. Execution Status tells whether the selector is part of the active runtime-equivalent flow.")
    lines.append("")
    lines.append("| Runtime Scenario Order | Scenario Sequence | Execution Status | Scenario | Step | Element Logical Name | Element Type | Action Code | JSON Location |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for row in selector_rows[:cfg.json_table_max_rows]:
        *cells, json_loc = row
        lines.append("| " + " | ".join(markdown_escape(x, 5000) for x in cells) + f" | `{json_loc}` |")
    if not selector_rows:
        lines.append("|  |  |  |  |  |  |  |  |  |")
    lines.append("")

    def render_scenario_details(section_title: str, detail_records: List[Dict[str, Any]], include_notice: str) -> None:
        lines.append(section_title)
        lines.append("")
        lines.append(include_notice)
        lines.append("")
        for runtime_order, record in enumerate(detail_records, start=1):
            scenario = record["scenario"]
            lines.append(f"### Scenario {runtime_order}: {markdown_escape(scenario.get('name', ''), 1000)}")
            lines.append("")
            meta = {
                "Runtime Scenario Order": runtime_order,
                "Scenario Sequence": record["runtime_sequence"],
                "Execution Status": "Active / Executable" if record["active"] else "Excluded / Not Executed",
                "Exclusion Reason": record.get("exclusion_reason", ""),
                "Logical Name": scenario.get("logical_name", ""),
                "Scenario Type": scenario.get("scenario_type", ""),
                "Reusable": flag_label(scenario.get("reusable")),
                "Status": flag_label(scenario.get("status")),
                "Include Status": flag_label(scenario.get("include_status")),
                "Raw Sequence": scenario.get("sequence", ""),
                "Original Sequence": scenario.get("original_sequence", ""),
                "Comment": scenario.get("comment_field", ""),
                "Dataset Type": scenario.get("dataset_type", ""),
                "Dataset Name": scenario.get("dataset_name", ""),
                "Step Count": len(scenario.get("scenarioFlowList") or []),
                "JSON Location": format_json_location(record["raw_index"]),
            }
            lines.extend(key_value_table_markdown("#### Scenario Metadata", {k: str(v) for k, v in meta.items()}))
            step_rows = testsavvy_testcase_step_rows(scenario, record["active"], record["raw_index"])
            header = [
                "Runtime Step Order", "Step Sequence", "Execution Status", "Exclusion Reason",
                "Step Name", "Action", "Type", "Classification", "English Text",
                "Element Logical Name", "Element Type", "Dataset", "Dataset Header", "Value",
                "Encrypted", "Tool", "Requires Data", "Uses Override", "SQL Command", "Action Code", "JSON Location", "Raw Step JSON"
            ]
            lines.append("#### Scenario Step Details")
            lines.append("")
            lines.append("| " + " | ".join(header) + " |")
            lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            for step_row in step_rows:
                cells = []
                for col in header:
                    val = step_row.get(col, "")
                    if col == "JSON Location":
                        cells.append(f"`{markdown_escape(val, 1000)}`")
                    else:
                        cells.append(markdown_escape(val, 8000))
                lines.append("| " + " | ".join(cells) + " |")
            if not step_rows:
                lines.append("| " + " | ".join([""] * len(header)) + " |")
            lines.append("")

    render_scenario_details(
        "## Active Runtime-Equivalent Scenarios and Steps",
        active_records,
        "These scenarios and steps are sorted by runtime-equivalent sequence and should be used to understand what the TestSavvy test case actually does.",
    )
    render_scenario_details(
        "## Excluded / Inactive Scenarios and Steps",
        excluded_records,
        "These scenarios and steps exist in the JSON export but are not expected to execute. They are useful for troubleshooting design drift and disabled coverage.",
    )

    if cfg.include_json_inventory:
        lines.extend(json_inventory_markdown(data, cfg.json_inventory_max_rows))

    return "\n".join(lines)

def flatten_interface_elements(node: Dict[str, Any], rows: Optional[List[Dict[str, Any]]] = None) -> List[Dict[str, Any]]:
    """
    Flatten a TestSavvy interface repository tree.
    """
    if rows is None:
        rows = []
    actions = node.get("actions") or []
    children = node.get("children") or []
    rows.append({
        "Interface Map": node.get("interfaceMapName", ""),
        "Element Name": node.get("interfaceElementName", ""),
        "Text Path": node.get("textPath", ""),
        "ID Path": node.get("idPath", ""),
        "Depth": node.get("depth", ""),
        "Sequence": node.get("sequence", ""),
        "Interface Type": node.get("interfaceType", ""),
        "Element Type": node.get("interfaceElementTypeName", ""),
        "Default Enter Action": node.get("defaultEnterActionName", ""),
        "Default Verify Action": node.get("defaultVerifyActionName", ""),
        "Override": normalize_bool_text(str(node.get("overrideFlag", ""))),
        "Sync Properties": normalize_bool_text(str(node.get("syncProperties", ""))),
        "Encrypt Data": normalize_bool_text(str(node.get("encryptData", ""))),
        "Template": node.get("templateName", ""),
        "Automation Profile": node.get("automationProfileName", ""),
        "Import Export Key": node.get("importExportKey", ""),
        "Action Count": len(actions),
        "Child Count": len(children),
        "actions": actions,
    })
    for child in children:
        if isinstance(child, dict):
            flatten_interface_elements(child, rows)
    return rows


def testsavvy_interface_json_markdown(source_name: str, data: Dict[str, Any], cfg: AppConfig) -> str:
    """
    Convert a TestSavvy interface repository JSON export into NotebookLM-oriented Markdown.
    """
    export = data.get("export", {}) or {}
    rows = flatten_interface_elements(export)

    lines: List[str] = []
    map_name = export.get("interfaceMapName") or export.get("interfaceElementName") or Path(source_name).stem
    lines.append(f"# TestSavvy Interface Repository JSON Payload: {markdown_escape(map_name, 1000)}")
    lines.append("")
    lines.append("## Document Identity")
    lines.append("")
    lines.append("| Field | Value |")
    lines.append("| --- | --- |")
    lines.append(f"| Source File | {markdown_escape(source_name, 500)} |")
    lines.append("| JSON Export Type | TestSavvy Interface Repository Export |")
    lines.append(f"| Machine Name | {markdown_escape(data.get('machineName', ''), 500)} |")
    lines.append(f"| Client Name | {markdown_escape(data.get('clientName', ''), 500)} |")
    lines.append(f"| Interface Map Name | {markdown_escape(map_name, 500)} |")
    lines.append(f"| Interface Element Count | {len(rows)} |")
    lines.append("")

    lines.append("## NotebookLM JSON Field Map")
    lines.append("")
    lines.append("| Business Term | JSON Location | Markdown Location | Meaning |")
    lines.append("| --- | --- | --- | --- |")
    field_rows = [
        ("Interface Map Name", "$.export.interfaceMapName", "## Document Identity; ## Interface Root Summary", "Top-level interface map, such as FIN."),
        ("Interface Element Name", "$.export.children[].interfaceElementName", "## Interface Element Index", "Human-readable UI element name."),
        ("Text Path", "$.export.children[].textPath", "## Interface Element Index", "Full logical tree path to the UI element."),
        ("ID Path", "$.export.children[].idPath", "## Interface Element Index", "Internal TestSavvy path/id chain."),
        ("Actions", "$.export.children[].actions[]", "## Interface Actions and Properties", "Available enter/verify/click action definitions."),
        ("Selector Properties", "$.export.children[].actions[].properties[]", "## Selector Property Index", "Locator properties such as xpath, id, name, aria-label, data-qa, data-qa-id."),
    ]
    for term, jloc, mdloc, meaning in field_rows:
        lines.append(f"| {term} | `{jloc}` | {mdloc} | {meaning} |")
    lines.append("")

    root_summary = {k: compact_json_value(v, 1000) for k, v in export.items() if k not in {"children", "actions"}}
    lines.extend(key_value_table_markdown("## Interface Root Summary", root_summary))

    lines.append("## Interface Element Index")
    lines.append("")
    lines.append("| # | Text Path | Element Name | ID Path | Depth | Type | Default Enter | Default Verify | Actions | Children | Sync Properties | Encrypt Data |")
    lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for idx, row in enumerate(rows[:cfg.json_table_max_rows], start=1):
        lines.append(
            f"| {idx} | {markdown_escape(row.get('Text Path', ''), 2000)} | "
            f"{markdown_escape(row.get('Element Name', ''), 1000)} | "
            f"{markdown_escape(row.get('ID Path', ''), 1000)} | "
            f"{markdown_escape(row.get('Depth', ''), 100)} | "
            f"{markdown_escape(row.get('Element Type', ''), 500)} | "
            f"{markdown_escape(row.get('Default Enter Action', ''), 500)} | "
            f"{markdown_escape(row.get('Default Verify Action', ''), 500)} | "
            f"{markdown_escape(row.get('Action Count', ''), 50)} | "
            f"{markdown_escape(row.get('Child Count', ''), 50)} | "
            f"{markdown_escape(row.get('Sync Properties', ''), 100)} | "
            f"{markdown_escape(row.get('Encrypt Data', ''), 100)} |"
        )
    if len(rows) > cfg.json_table_max_rows:
        lines.append(f"| TRUNCATED | Increase json_table_max_rows to include all {len(rows)} interface elements. |  |  |  |  |  |  |  |  |  |  |")
    lines.append("")

    action_rows = []
    selector_rows = []
    selector_names = {"xpath", "id", "name", "aria-label", "data-qa", "data-qa-id", "css", "class", "text"}
    for row in rows:
        for action in row.get("actions") or []:
            props = action.get("properties") or []
            prop_text = "; ".join(f"{p.get('propertyName')}={p.get('propertyValue')}" for p in props)
            action_rows.append((row.get("Text Path", ""), row.get("Element Name", ""), action.get("actionName", ""), action.get("stepDefinitionActionId", ""), prop_text))
            selected = []
            for p in props:
                pname = str(p.get("propertyName") or "")
                pval = p.get("propertyValue")
                if pname.lower() in selector_names or pval:
                    selected.append(f"{pname}={pval}")
            if selected:
                selector_rows.append((row.get("Text Path", ""), row.get("Element Name", ""), action.get("actionName", ""), "; ".join(selected)))

    lines.append("## Interface Actions and Properties")
    lines.append("")
    lines.append("| Text Path | Element Name | Action | Step Definition Action ID | Properties |")
    lines.append("| --- | --- | --- | --- | --- |")
    for row in action_rows[:cfg.json_table_max_rows]:
        lines.append("| " + " | ".join(markdown_escape(x, 5000) for x in row) + " |")
    if not action_rows:
        lines.append("|  |  |  |  |  |")
    if len(action_rows) > cfg.json_table_max_rows:
        lines.append(f"| TRUNCATED | Increase json_table_max_rows to include all {len(action_rows)} actions. |  |  |  |")
    lines.append("")

    lines.append("## Selector Property Index")
    lines.append("")
    lines.append("Use this section to analyze locator stability, dynamic `data-qa-id`, XPath, IDs, names, and aria-labels.")
    lines.append("")
    lines.append("| Text Path | Element Name | Action | Selector / Property Values |")
    lines.append("| --- | --- | --- | --- |")
    for row in selector_rows[:cfg.json_table_max_rows]:
        lines.append("| " + " | ".join(markdown_escape(x, 5000) for x in row) + " |")
    if not selector_rows:
        lines.append("|  |  |  |  |")
    if len(selector_rows) > cfg.json_table_max_rows:
        lines.append(f"| TRUNCATED | Increase json_table_max_rows to include all {len(selector_rows)} selector rows. |  |  |")
    lines.append("")

    if cfg.include_json_inventory:
        lines.extend(json_inventory_markdown(data, cfg.json_inventory_max_rows))

    return "\n".join(lines)


def testsavvy_group_json_markdown(source_name: str, data: List[Any], cfg: AppConfig) -> str:
    """
    Render group JSON exports or other top-level JSON arrays.
    """
    lines = [
        f"# TestSavvy JSON Array Payload: {markdown_escape(Path(source_name).stem, 1000)}",
        "",
        "## Document Identity",
        "",
        "| Field | Value |",
        "| --- | --- |",
        f"| Source File | {markdown_escape(source_name, 500)} |",
        "| JSON Export Type | Top-Level Array / Group Export |",
        f"| Item Count | {len(data)} |",
        "",
        "## Array Item Summary",
        "",
        "| Index | Type | Summary |",
        "| --- | --- | --- |",
    ]
    for idx, item in enumerate(data[:cfg.json_table_max_rows]):
        lines.append(f"| {idx} | {json_type_name(item)} | {markdown_escape(compact_json_value(item, 3000), 3000)} |")
    if len(data) > cfg.json_table_max_rows:
        lines.append(f"| TRUNCATED | notice | Increase json_table_max_rows to include all {len(data)} rows. |")
    lines.append("")
    if cfg.include_json_inventory:
        lines.extend(json_inventory_markdown(data, cfg.json_inventory_max_rows))
    return "\n".join(lines)


def json_to_markdown_from_data(source_name: str, data: Any, cfg: AppConfig) -> str:
    """
    Dispatch JSON conversion based on known TestSavvy JSON export shape.
    """
    if cfg.json_output_mode == "raw_markdown":
        pretty = json.dumps(data, indent=cfg.indent, ensure_ascii=False, sort_keys=cfg.json_sort_keys)
        return f"# Pretty Printed JSON\n\n**Source file:** `{source_name}`\n\n```json\n{pretty.rstrip()}\n```\n"

    kind = detect_testsavvy_json_kind(data)
    if kind == "test_case":
        return testsavvy_testcase_json_markdown(source_name, data, cfg)
    if kind == "interface_repository":
        return testsavvy_interface_json_markdown(source_name, data, cfg)
    if kind == "test_case_group":
        return testsavvy_group_json_markdown(source_name, data, cfg)

    lines = [
        f"# Generic JSON Payload: {markdown_escape(Path(source_name).stem, 1000)}",
        "",
        "## Document Identity",
        "",
        "| Field | Value |",
        "| --- | --- |",
        f"| Source File | {markdown_escape(source_name, 500)} |",
        "| JSON Export Type | Generic JSON |",
        f"| Top-Level Type | {json_type_name(data)} |",
        "",
    ]
    if cfg.include_json_inventory:
        lines.extend(json_inventory_markdown(data, cfg.json_inventory_max_rows))
    else:
        pretty = json.dumps(data, indent=cfg.indent, ensure_ascii=False, sort_keys=cfg.json_sort_keys)
        lines.extend(["## Pretty Printed JSON", "", "```json", pretty.rstrip(), "```", ""])
    return "\n".join(lines)


def json_to_markdown(src: Path, cfg: AppConfig) -> str:
    """
    Load a JSON source file and convert it into Markdown.
    """
    data = json.loads(src.read_text(encoding="utf-8-sig"))
    return json_to_markdown_from_data(src.name, data, cfg)


def testsavvy_xml_output_filename(src: Path, cfg: AppConfig) -> str:
    """
    Try to name XML execution Markdown as <TestCaseID>_<RunID>.md.
    Falls back to normalized source stem.
    """
    try:
        tree = ET.parse(str(src), parser=ET.XMLParser())
        root = tree.getroot()
        auto_seq = extract_automation_sequence(root)
        if auto_seq is not None:
            tc_id = safe_filename_component(auto_seq.attrib.get("id", "") or src.stem)
            run_id = safe_filename_component(auto_seq.attrib.get("run_id", "") or "run")
            return f"{tc_id}_{run_id}.md"
    except Exception:
        pass

    stem = safe_filename_component(src.stem)
    stem = re.sub(r"\.dat$", "", stem, flags=re.IGNORECASE)
    if "_" not in stem:
        stem = f"XML_{stem}"
    return f"{stem}.md"


def compute_smart_output_path(src: Path, input_root: Path, cfg: AppConfig) -> Path:
    """
    Compute a NotebookLM-friendly output path for XML/JSON while retaining old behavior for Excel.
    """
    ext = src.suffix.lower().lstrip(".")
    if (cfg.output_extension or "").lower() == ".md" and cfg.normalize_markdown_filenames and ext in {"xml", "json"}:
        if ext == "xml":
            out_name = testsavvy_xml_output_filename(src, cfg)
        else:
            try:
                data = json.loads(src.read_text(encoding="utf-8-sig"))
                out_name = testsavvy_json_output_filename(data, src.name)
            except Exception:
                stem = safe_filename_component(src.stem)
                if "_" not in stem:
                    stem = f"JSON_{stem}"
                out_name = f"{stem}.md"

        if cfg.output_structure == "flat":
            return cfg.output_dir / out_name
        rel_parent = src.relative_to(input_root).parent
        return cfg.output_dir / rel_parent / out_name

    return compute_output_path(src, input_root, cfg.output_dir, cfg.output_structure, cfg.output_extension)


def process_standard_file(src: Path, out_path: Path, cfg: AppConfig):
    """
    Process JSON or XML source files and write the configured output artifact.
    """
    if out_path.exists() and not cfg.overwrite:
        return True, f"SKIP (exists): {src}"

    try:
        ensure_parent_dir(out_path)
        ext = src.suffix.lower().lstrip(".")

        if ext == "json":
            if (cfg.output_extension or "").lower() == ".md":
                out_path.write_text(json_to_markdown(src, cfg), encoding="utf-8")
            else:
                pretty = pretty_print_json(src, cfg.indent, cfg.json_sort_keys)
                out_path.write_text(pretty, encoding="utf-8")

        elif ext == "xml":
            if (cfg.output_extension or "").lower() == ".md":
                md = xml_to_markdown(src, cfg)
                out_path.write_text(md, encoding="utf-8")
            else:
                out_path.write_bytes(pretty_print_xml(src, cfg.indent, cfg.xml_encoding, cfg.preserve_xml_declaration))

        else:
            return True, f"SKIP (unsupported ext): {src}"

        return True, f"OK: {src} -> {out_path}"
    except Exception as e:
        return False, f"FAIL: {src} ({type(e).__name__}: {e})"


def process_zip_file(src: Path, cfg: AppConfig) -> Tuple[int, int, List[str]]:
    """
    Process JSON/XML files contained inside a zip without requiring the user to manually extract it.
    Output is written under output/<zip-stem>/ using NotebookLM-friendly filenames.
    """
    messages: List[str] = []
    written = 0
    failed = 0
    used_names: Dict[str, int] = {}

    if not cfg.process_zip_files:
        return 0, 0, [f"SKIP (zip processing disabled): {src}"]

    try:
        with zipfile.ZipFile(src) as zipf:
            members = [m for m in zipf.infolist() if not m.is_dir()]
            for member in members:
                inner_name = member.filename
                inner_ext = Path(inner_name).suffix.lower().lstrip(".")
                if inner_ext not in {"json", "xml"}:
                    messages.append(f"SKIP (zip member unsupported): {src}!{inner_name}")
                    continue

                try:
                    raw = zipf.read(member)
                    zip_output_root = cfg.output_dir / safe_filename_component(src.stem)

                    if inner_ext == "json":
                        data = json.loads(raw.decode("utf-8-sig"))
                        out_name = testsavvy_json_output_filename(data, Path(inner_name).name)
                        if out_name in used_names:
                            used_names[out_name] += 1
                            out_name = testsavvy_json_output_filename(data, Path(inner_name).name, used_names[out_name])
                        else:
                            used_names[out_name] = 1
                        out_path = zip_output_root / out_name
                        if out_path.exists() and not cfg.overwrite:
                            messages.append(f"SKIP (exists): {src}!{inner_name}")
                            continue
                        ensure_parent_dir(out_path)
                        out_path.write_text(json_to_markdown_from_data(Path(inner_name).name, data, cfg), encoding="utf-8")
                        messages.append(f"OK: {src}!{inner_name} -> {out_path}")
                        written += 1

                    elif inner_ext == "xml":
                        tmp_root = ET.fromstring(raw)
                        out_name = f"{safe_filename_component(Path(inner_name).stem)}.md"
                        out_path = zip_output_root / out_name
                        if out_path.exists() and not cfg.overwrite:
                            messages.append(f"SKIP (exists): {src}!{inner_name}")
                            continue
                        ensure_parent_dir(out_path)
                        md = testsavvy_execution_report_markdown(Path(inner_name), tmp_root, cfg.include_xml_inventory)
                        out_path.write_text(md, encoding="utf-8")
                        messages.append(f"OK: {src}!{inner_name} -> {out_path}")
                        written += 1

                except Exception as e:
                    messages.append(f"FAIL: {src}!{inner_name} ({type(e).__name__}: {e})")
                    failed += 1

    except Exception as e:
        messages.append(f"FAIL: {src} ({type(e).__name__}: {e})")
        failed += 1

    return written, failed, messages


def generate_notebooklm_payload_guide_markdown() -> str:
    """
    Generate a companion Markdown guide explaining TestSavvy XML/JSON payload structures,
    including the critical JSON ordering and active/excluded execution rules.
    """
    return """# NotebookLM Guide: TestSavvy XML and JSON Markdown Payloads

## Purpose

This guide teaches NotebookLM how to read generated TestSavvy Markdown files. Upload this guide together with the generated `.md` files from XML execution payloads, Test Case JSON exports, and Interface Repository JSON exports.

## Supported Payload Types

| Payload Type | Source Shape | Generated Markdown Purpose |
| --- | --- | --- |
| TestSavvy XML Execution Payload | XML containing `automation_sequence`, `scenario`, and `step` nodes | Explains what TestSavvy actually executed in a specific run. This is the closest source to runtime truth. |
| TestSavvy Test Case JSON Export | JSON object containing `automationSequence` and `sequenceList` | Explains the authored test case definition, including active scenarios, excluded scenarios, reusable steps, datasets, custom functions, and selectors. |
| TestSavvy Interface Repository JSON Export | JSON object containing `export.interfaceMapName` and recursive `children` | Explains the interface map tree, logical element paths, actions, selector properties, and locator metadata. |
| TestSavvy Group / Array JSON Export | JSON top-level array | Explains grouped/exported array content and provides a JSON inventory. |

## Most Important Rule: JSON Array Order Is Not Execution Order

For Test Case JSON exports, do not assume `$.sequenceList[]` physical array order is the order TestSavvy executes. The generator sorts scenarios by `sequence`, falling back to `original_sequence`, and preserves the raw JSON array location separately.

Use these sections this way:

| Section | Use For | Do Not Use For |
| --- | --- | --- |
| `## Active Runtime-Equivalent Scenario Index` | Understanding what the test case is expected to execute, sorted by TestSavvy sequence. | Debugging raw export order. |
| `## Active Runtime-Equivalent Scenarios and Steps` | Step-by-step explanation of the active authored flow. | Counting disabled/excluded coverage. |
| `## Excluded / Inactive Scenario Index` | Finding scenarios that exist in the JSON but are excluded from execution. | Answering what the test actually does. |
| `## Excluded / Inactive Scenarios and Steps` | Troubleshooting disabled design content. | Runtime coverage claims. |
| `## Full Raw JSON Scenario Inventory` | Debugging the raw export array and locating JSON paths. | Execution order. |
| `## JSON Inventory` | Finding deeply nested source values. | Human-readable flow analysis unless no other section has the value. |

## Active vs Excluded Execution Rules

A TestSavvy JSON scenario should be considered active/executable only when:

```text
scenario.status is not false
AND scenario.include_status is not false
```

A TestSavvy JSON step should be considered active/executable only when:

```text
parent scenario is active
AND step.status is not false
AND step.include_status is not false
AND stepAssociation.status is not false
```

If any of those values are explicitly false, the generated Markdown marks the row as `Excluded / Not Executed` and records an `Exclusion Reason`, such as `scenario.include_status=false`, `step.include_status=false`, or `stepAssociation.status=false`.

## How To Answer Coverage Questions

When the user asks a coverage question such as:

```text
Which TestSavvy automated test cases click the Submit button?
```

NotebookLM should use `## Submit Button Execution Index` and count only rows where `Execution Status` is `Active / Executable`.

If the user asks:

```text
Which test cases contain Submit, including excluded steps?
```

NotebookLM may include both `Active / Executable` and `Excluded / Not Executed` rows, but it must label them separately.

## Test Case JSON Mapping

| Business Term | JSON Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Test Case ID | `$.automationSequence.id` | `## Document Identity`; `## Test Case Summary` | The TestSavvy test case / automation sequence identifier. |
| Test Case Name | `$.automationSequence.name` | Document title; `## Document Identity`; `## Test Case Summary` | Human-readable name of the test case. |
| Raw Scenario List | `$.sequenceList[]` | `## Full Raw JSON Scenario Inventory` | Physical JSON export array. This is not necessarily execution order. |
| Scenario Execution Order | `$.sequenceList[].sequence` | `## Active Runtime-Equivalent Scenario Index` | Primary scenario ordering field. |
| Fallback Scenario Order | `$.sequenceList[].original_sequence` | `## Active Runtime-Equivalent Scenario Index` | Fallback ordering field when `sequence` is missing. |
| Scenario Include Status | `$.sequenceList[].include_status` | Scenario index and scenario metadata tables | False means the scenario exists in JSON but should not be counted as executed coverage. |
| Scenario Status | `$.sequenceList[].status` | Scenario index and scenario metadata tables | False means the scenario is inactive. |
| Scenario Steps | `$.sequenceList[].scenarioFlowList[]` | Scenario step detail tables | Step/action definitions inside each scenario, sorted by step sequence. |
| Step Include Status | `$.sequenceList[].scenarioFlowList[].include_status` | Scenario step detail table `Execution Status` and `Exclusion Reason` | False means the step exists in JSON but is excluded. |
| Step Data Value | `$.sequenceList[].scenarioFlowList[].stepAssociation.value` | Scenario step detail table column `Value` | Data entered, selected, clicked, or passed into a custom function. |
| Step Dataset | `$.sequenceList[].scenarioFlowList[].stepAssociation.dataset_name` | Scenario step detail table column `Dataset` | Dataset used by the step. |
| Step Dataset Header | `$.sequenceList[].scenarioFlowList[].dataset_header_name` | Scenario step detail table column `Dataset Header` | Dataset field/header used by the step. |
| Action Code | `$.sequenceList[].scenarioFlowList[].action_code` | Scenario step detail table column `Action Code`; `## Selector / Action Code Index` | Selenium/WebDriver/custom code for the step. |
| Custom Function | Step `Value` or `Action Code` containing `@...` | `## Custom Functions and Dynamic Values` | TestSavvy custom functions and dynamic values. |

## Test Case JSON Sections

| Section | Meaning |
| --- | --- |
| `## Critical Ordering Rule` | Explicit reminder that raw JSON array order is not execution order. |
| `## Active Runtime-Equivalent Scenario Index` | Active scenario flow sorted by `sequence` / `original_sequence`. |
| `## Excluded / Inactive Scenario Index` | Disabled scenarios, still sorted by sequence for context. |
| `## Full Raw JSON Scenario Inventory` | Raw export array order and JSON path references. |
| `## Submit Button Execution Index` | Purpose-built coverage table for Submit button questions. |
| `## Custom Functions and Dynamic Values` | Dynamic values, `@storeResult`, `@cgifx`, `@advantagefx`, and other custom function usage. |
| `## Selector / Action Code Index` | Locator/action-code analysis table. |
| `## Active Runtime-Equivalent Scenarios and Steps` | Detailed active scenario and step flow. |
| `## Excluded / Inactive Scenarios and Steps` | Detailed disabled scenario and step content. |
| `## JSON Inventory` | Flattened fallback source of truth. |

## XML Execution Payload Mapping

XML execution Markdown files are already runtime ordered because they come from the execution payload.

| Business Term | XML Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Test Case ID | `automation_sequence/@id` | `## Execution Summary`; `## Automation Sequence Attributes` | TestSavvy test case / automation sequence identifier. |
| Test Case Name | `automation_sequence/@name` | Title; `## Execution Summary`; `## Automation Sequence Attributes` | Human-readable automated test name. |
| Run ID | `automation_sequence/@run_id` | `## Execution Summary`; `## Automation Sequence Attributes` | Specific execution run. |
| Runtime Scenario Order | `scenario/@count_id` and payload order | `## Scenarios and Steps` | Actual execution order in that run. |
| Runtime Step Order | `step/@count_id` and payload order | Scenario step detail table | Actual step order in that run. |
| Action Code | `step/@action_code` | Scenario step detail table column `Action Code` | Runtime automation code/selectors. |

## Interface Repository JSON Mapping

| Business Term | JSON Location | Generated Markdown Location | Meaning |
| --- | --- | --- | --- |
| Interface Map Name | `$.export.interfaceMapName` | `## Document Identity`; `## Interface Root Summary` | Top-level interface map, such as `FIN`. |
| Interface Element Name | Recursive `interfaceElementName` under `$.export.children[]` | `## Interface Element Index` | Human-readable UI element name. |
| Text Path | Recursive `textPath` under `$.export.children[]` | `## Interface Element Index` | Full logical path to the interface element. |
| ID Path | Recursive `idPath` under `$.export.children[]` | `## Interface Element Index` | TestSavvy internal path/id chain. |
| Actions | Recursive `actions[]` under each element | `## Interface Actions and Properties` | Available action definitions such as Click, Enter, Verify. |
| Selector Properties | `actions[].properties[]` | `## Selector Property Index` | Locator properties such as XPath, ID, name, aria-label, `data-qa`, and `data-qa-id`. |

## Selector Analysis Rules

For Test Case JSON, inspect `Action Code`, `Raw Step JSON`, and `## Selector / Action Code Index`. For Interface Repository JSON, inspect `## Selector Property Index` and `## Interface Actions and Properties`. These sections are the best sources for XPath, `data-qa`, `data-qa-id`, `aria-label`, ID, name, and other locator decisions.

## NotebookLM-Friendly File Names

Generated Markdown filenames are normalized so they contain underscores and end in `.md`.

| Source Type | Filename Pattern |
| --- | --- |
| XML execution payload | `<TestCaseID>_<RunID>.md` |
| Test Case JSON export | `TC_<TestCaseID>.md` |
| Interface Repository JSON export | `IR_<InterfaceMapName>.md` |
| Group JSON export | `GROUP_<source>.md` |

Avoid stale filenames like `998_19586.dat.md`; regenerate them as `998_19586.md`.
"""

# -----------------------------------------------------------------------------
# Command-line entry point
# -----------------------------------------------------------------------------
def main() -> int:
    """
    Parse command-line arguments, load configuration, process all source files, and print a summary.
    """
    ap = argparse.ArgumentParser(description="Process TestSavvy XML, JSON, ZIP, and Excel artifacts to Markdown.")
    ap.add_argument("--config", required=True, help="Path to config JSON")
    args = ap.parse_args()

    cfg = load_config(Path(args.config))
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

    total_considered = 0
    written = 0
    failed = 0
    combined_chunks: List[str] = []

    for input_dir in cfg.input_dirs:
        if not input_dir.exists():
            print(f"WARN: input dir does not exist: {input_dir}")
            continue

        for src in iter_files(input_dir, cfg.recursive, cfg.file_types):
            if is_excluded(src, cfg.exclude_patterns):
                print(f"SKIP (excluded): {src}")
                continue

            total_considered += 1
            ext = src.suffix.lower().lstrip(".")

            if ext == "zip":
                z_written, z_failed, z_messages = process_zip_file(src, cfg)
                for msg in z_messages:
                    print(msg)
                written += z_written
                failed += z_failed
                continue

            if ext in {"xlsx", "xlsm"}:
                wb = load_workbook(filename=str(src), data_only=True, read_only=True)
                try:
                    for ws in wb.worksheets:
                        if not cfg.excel_include_empty_sheets and not worksheet_has_data(ws):
                            print(f"SKIP (empty sheet): {src} [{ws.title}]")
                            continue

                        md = worksheet_to_markdown(src, ws, cfg.excel_max_cell_length)

                        if cfg.xml_markdown_mode == "combined":
                            combined_chunks.append(build_excel_combined_section(src, ws.title, md))
                            print(f"OK: {src} [{ws.title}] -> [combined markdown buffer]")
                            written += 1
                            continue

                        out_path = build_excel_sheet_output_path(
                            src=src,
                            input_root=input_dir,
                            output_dir=cfg.output_dir,
                            output_structure=cfg.output_structure,
                            forced_ext=cfg.output_extension,
                            sheet_name=ws.title,
                            include_sheet_name_in_filename=cfg.excel_sheet_name_in_filename,
                        )
                        if out_path.exists() and not cfg.overwrite:
                            print(f"SKIP (exists): {src} [{ws.title}]")
                            continue
                        try:
                            ensure_parent_dir(out_path)
                            out_path.write_text(md, encoding="utf-8")
                            print(f"OK: {src} [{ws.title}] -> {out_path}")
                            written += 1
                        except Exception as e:
                            print(f"FAIL: {src} [{ws.title}] ({type(e).__name__}: {e})")
                            failed += 1
                finally:
                    wb.close()
                continue

            if cfg.xml_markdown_mode == "combined" and ext in {"xml", "json"} and (cfg.output_extension or "").lower() == ".md":
                try:
                    if ext == "xml":
                        combined_chunks.append(build_combined_section_for_file(src, cfg))
                    else:
                        combined_chunks.append(json_to_markdown(src, cfg))
                    print(f"OK: {src} -> [combined markdown buffer]")
                    written += 1
                except Exception as e:
                    print(f"FAIL: {src} ({type(e).__name__}: {e})")
                    failed += 1
                continue

            out_path = compute_smart_output_path(src, input_dir, cfg)
            success, msg = process_standard_file(src, out_path, cfg)
            print(msg)
            if success and msg.startswith("OK:"):
                written += 1
            if not success:
                failed += 1

    if combined_chunks:
        combined_path = cfg.output_dir / cfg.combined_output_filename
        try:
            ensure_parent_dir(combined_path)
            combined_text = "\n\n---\n\n".join(chunk.strip() for chunk in combined_chunks if chunk.strip()) + "\n"
            combined_path.write_text(combined_text, encoding="utf-8")
            print(f"OK: wrote combined markdown -> {combined_path}")
        except Exception as e:
            print(f"FAIL: combined markdown ({type(e).__name__}: {e})")
            failed += 1

    if cfg.generate_notebooklm_guide and (cfg.output_extension or "").lower() == ".md":
        guide_path = cfg.output_dir / cfg.notebooklm_guide_filename
        try:
            ensure_parent_dir(guide_path)
            guide_path.write_text(generate_notebooklm_payload_guide_markdown(), encoding="utf-8")
            print(f"OK: wrote NotebookLM payload guide -> {guide_path}")
            written += 1
        except Exception as e:
            print(f"FAIL: NotebookLM payload guide ({type(e).__name__}: {e})")
            failed += 1

    print("\n=== Summary ===")
    print(f"Total considered: {total_considered}")
    print(f"Outputs written:  {written}")
    print(f"Failed:          {failed}")

    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
