#!/usr/bin/env python3
"""
PUBLIC_INTERFACE
Standalone Script: Read operationId values from a Swagger/OpenAPI JSON text file (swagger.txt),
add/fill a new 'operationId' column in an Excel file (testcases.xlsx) by matching each test row's
endpoint and method, and save the result to a new, visually formatted Excel file
(testcases_with_operationId.xlsx).

Default locations (adjust via CLI args if needed):
- Swagger text file: mock_backend_service/inputs/swagger.txt
- Input Excel file:  mock_backend_service/inputs/testcases.xlsx
- Output Excel file: mock_backend_service/inputs/testcases_with_operationId.xlsx

Matching logic (documented in code):
- The script expects the testcases sheet to describe each API call by at least the HTTP method and path.
  Common column names (case-insensitive) it looks for:
    • method: "method", "http_method"
    • path:   "path", "endpoint", "url", "api_path"
  If these columns are named differently or missing, the script will:
    • Leave the operationId blank for those rows and continue.
    • Print a note indicating what columns were found and which are required for automatic matching.
- If an 'operationId' column already exists, its non-empty values are preserved; only empty cells are auto-filled.
- The operationId is determined by exact method+path match to the swagger paths (e.g., ('POST', '/pet')).

Output formatting:
- The resulting Excel is styled with:
    • Bold headers, thin cell borders, wrapped text for data cells
    • Frozen top row
    • Reasonable column widths
- Only the specified sheet is modified; all other sheets remain unchanged.

Dependencies:
- pandas
- openpyxl

Usage examples:
  python mock_backend_service/scripts/add_operationId_column.py
  python mock_backend_service/scripts/add_operationId_column.py \
    --swagger mock_backend_service/inputs/swagger.txt \
    --input   mock_backend_service/inputs/testcases.xlsx \
    --output  mock_backend_service/inputs/testcases_with_operationId.xlsx \
    --sheet "Standard Template"
"""

import argparse
import json
import os
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def _read_text(path: str) -> str:
    with open(path, "rb") as f:
        return f.read().decode("utf-8", errors="replace")


def _iter_operations_from_paths(paths: Dict[str, Any]) -> Iterable[Tuple[str, str, Dict[str, Any]]]:
    """
    Iterate over HTTP operations from an OpenAPI/Swagger 'paths' object.
    Yields (method_upper, path, operation_object_dict).
    """
    if not isinstance(paths, dict):
        return
    http_methods = {"get", "post", "put", "patch", "delete", "head", "options", "trace"}
    for path, path_item in paths.items():
        if not isinstance(path_item, dict):
            continue
        for method, op_obj in path_item.items():
            if isinstance(method, str) and method.lower() in http_methods and isinstance(op_obj, dict):
                yield method.upper(), path, op_obj


# PUBLIC_INTERFACE
def parse_operation_ids_from_swagger_text(swagger_text: str) -> Dict[Tuple[str, str], str]:
    """
    PUBLIC_INTERFACE
    Parse an OpenAPI/Swagger JSON string and return a mapping of (METHOD, path) -> operationId.

    - Expects swagger_text to be JSON (Swagger 2.0 or OpenAPI 3.x-like with 'paths').
    - Synthesizes an operationId when missing, but real operationId is preferred if present.
    """
    data = json.loads(swagger_text)
    paths = data.get("paths", {}) or {}

    mapping: Dict[Tuple[str, str], str] = {}
    seen_ids: set[str] = set()

    for method, path, op in _iter_operations_from_paths(paths):
        op_id = op.get("operationId")
        if not op_id:
            # Synthesize a readable id when missing
            op_id = f"{method.lower()}_{path}".replace("/", "_").replace("{", "").replace("}", "")
        # If duplicate operationId appears (shouldn't), keep first occurrence
        if (method, path) not in mapping:
            mapping[(method, path)] = op_id
        seen_ids.add(op_id)

    return mapping


def _normalize_colnames(cols: List[str]) -> Dict[str, str]:
    """
    Build a case-insensitive mapping of lowercased names to original names.
    """
    mapping: Dict[str, str] = {}
    for c in cols:
        if isinstance(c, str):
            mapping[c.strip().lower()] = c
    return mapping


def _find_method_and_path_columns(cols_map: Dict[str, str]) -> Tuple[Optional[str], Optional[str], List[str]]:
    """
    Identify likely method and path columns using flexible naming.
    Returns (method_col_name, path_col_name, notes)

    method candidates: method, http_method
    path candidates: path, endpoint, url, api_path
    """
    notes: List[str] = []
    method_col = None
    path_col = None

    for candidate in ("method", "http_method"):
        if candidate in cols_map:
            method_col = cols_map[candidate]
            break
    if not method_col:
        notes.append("No 'method' column found (looked for: method, http_method).")

    for candidate in ("path", "endpoint", "url", "api_path"):
        if candidate in cols_map:
            path_col = cols_map[candidate]
            break
    if not path_col:
        notes.append("No 'path' column found (looked for: path, endpoint, url, api_path).")

    return method_col, path_col, notes


def _ensure_operationid_column(df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    Ensure an 'operationId' column exists (case-insensitive).
    Returns updated DataFrame and the actual column name used.
    """
    cols_map = _normalize_colnames(list(df.columns))
    existing = cols_map.get("operationid")
    if existing:
        return df, existing
    # Create new column at the end
    df["operationId"] = None
    return df, "operationId"


def _apply_openpyxl_formatting(path: str, sheet_name: str) -> None:
    """
    Apply visual formatting to the specified sheet in the given workbook path.

    Formatting includes:
    - Bold header row, centered alignment.
    - Thin borders around all used cells.
    - Increased row height for header and data rows.
    - Reasonable column widths (auto-fit approximation based on content length).
    - Freeze top row for easier scrolling.
    """
    wb = load_workbook(path)
    if sheet_name not in wb.sheetnames:
        wb.save(path)
        return
    ws = wb[sheet_name]

    # Freeze header row
    ws.freeze_panes = "A2"

    # Determine used range
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # Styles
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_align = Alignment(vertical="top", wrap_text=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Apply header styling
    for col in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border

    # Apply borders and alignment to data cells
    for row in range(2, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = wrap_align
            cell.border = border

    # Row heights
    ws.row_dimensions[1].height = 28  # header
    for r in range(2, max_row + 1):
        ws.row_dimensions[r].height = 22

    # Compute column widths based on max content length (approximate)
    # Set a minimum width and cap the maximum.
    min_width = 16
    max_width = 60
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue
            text = str(val)
            # Consider line breaks and longer words
            for part in text.split("\n"):
                max_len = max(max_len, len(part))
        width = min(max(min_width, max_len + 2), max_width)
        ws.column_dimensions[col_letter].width = width

    wb.save(path)


# PUBLIC_INTERFACE
def process_excel_with_operation_ids(
    swagger_path: str,
    input_excel_path: str,
    output_excel_path: str,
    sheet_name: str = "Standard Template",
) -> None:
    """
    PUBLIC_INTERFACE
    Orchestrates the process:
      1) Parse swagger.txt (JSON) to build a mapping of (METHOD, path) -> operationId.
      2) Load testcases.xlsx, ensure an 'operationId' column exists on the target sheet.
      3) For each row, try to match using the method and path columns (flexible naming).
         - If matched, write the operationId to the operationId column.
         - If method/path missing or no match, leave the cell blank.
      4) Save to testcases_with_operationId.xlsx and visually format the target sheet.

    Notes on expected Excel columns (case-insensitive):
      - method: one of ['method', 'http_method']
      - path: one of ['path', 'endpoint', 'url', 'api_path']
      If your sheet uses different names, either rename them or extend the candidates in
      _find_method_and_path_columns. The script will handle missing columns gracefully and leave
      operationId blank when it cannot match.

    Only the specified sheet is modified; other sheets are preserved unchanged.
    """
    if not os.path.isfile(swagger_path):
        raise FileNotFoundError(f"Swagger file not found: {swagger_path}")
    if not os.path.isfile(input_excel_path):
        raise FileNotFoundError(f"Input Excel not found: {input_excel_path}")

    swagger_text = _read_text(swagger_path)
    method_path_to_opid = parse_operation_ids_from_swagger_text(swagger_text)

    # Read Excel with pandas, preserve all sheets
    xls = pd.ExcelFile(input_excel_path)
    sheets: Dict[str, pd.DataFrame] = {name: xls.parse(name) for name in xls.sheet_names}

    if sheet_name not in sheets:
        # If the desired sheet is missing, just write back unchanged to output
        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=name)
        # apply formatting attempt (no-op if sheet missing)
        _apply_openpyxl_formatting(output_excel_path, sheet_name)
        return

    df = sheets[sheet_name].copy()

    # Ensure 'operationId' column exists
    df, op_col = _ensure_operationid_column(df)

    # Find method and path columns with flexible naming
    cols_map = _normalize_colnames(list(df.columns))
    method_col, path_col, discovery_notes = _find_method_and_path_columns(cols_map)

    # Add a clear note as a comment in code (log to console) if columns are missing
    if discovery_notes:
        print("Column discovery notes:")
        for n in discovery_notes:
            print(f"- {n}")

    # Mapping: For each row, if method+path available, try to map to operationId
    for idx in range(len(df)):
        # Preserve non-empty existing opId
        current_val = df.at[idx, op_col]
        if isinstance(current_val, str) and current_val.strip():
            continue

        if not method_col or not path_col:
            # Cannot match without both columns; leave blank
            df.at[idx, op_col] = None
            continue

        method_val = str(df.at[idx, method_col] if method_col in df.columns else "").strip().upper()
        path_val = str(df.at[idx, path_col] if path_col in df.columns else "").strip()

        if not method_val or not path_val:
            df.at[idx, op_col] = None
            continue

        # Exact match required: (METHOD, path)
        opid = method_path_to_opid.get((method_val, path_val))
        df.at[idx, op_col] = opid if opid else None

    # Save back all sheets, replacing only the target one
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            if name == sheet_name:
                df.to_excel(writer, index=False, sheet_name=name)
            else:
                sdf.to_excel(writer, index=False, sheet_name=name)

    # Apply openpyxl-based visual formatting to make the sheet more readable
    _apply_openpyxl_formatting(output_excel_path, sheet_name)


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Parse operationIds from swagger.txt and write them into the 'operationId' column of the Excel."
    )
    parser.add_argument(
        "--swagger",
        default="mock_backend_service/inputs/swagger.txt",
        help="Path to swagger.txt (JSON content). Default: mock_backend_service/inputs/swagger.txt",
    )
    parser.add_argument(
        "--input",
        dest="input_excel",
        default="mock_backend_service/inputs/testcases.xlsx",
        help="Path to input Excel (.xlsx). Default: mock_backend_service/inputs/testcases.xlsx",
    )
    parser.add_argument(
        "--output",
        dest="output_excel",
        default="mock_backend_service/inputs/testcases_with_operationId.xlsx",
        help="Path to output Excel (.xlsx). Default: mock_backend_service/inputs/testcases_with_operationId.xlsx",
    )
    parser.add_argument(
        "--sheet",
        default="Standard Template",
        help='Worksheet name to modify. Default: "Standard Template"',
    )
    return parser


def main() -> int:
    args = _build_arg_parser().parse_args()
    try:
        process_excel_with_operation_ids(
            swagger_path=args.swagger,
            input_excel_path=args.input_excel,
            output_excel_path=args.output_excel,
            sheet_name=args.sheet,
        )
        print(
            f"Success: Filled 'operationId' in sheet '{args.sheet}' and saved to '{args.output_excel}'."
        )
        return 0
    except Exception as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
