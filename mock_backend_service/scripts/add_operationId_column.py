#!/usr/bin/env python3
"""
PUBLIC_INTERFACE
Standalone Script: Read operationId values from a Swagger/OpenAPI JSON text file (swagger.txt),
add a new 'operationId' column to an Excel file (testcases.xlsx), attempt basic mapping,
and save the result to a new Excel file (testcases_with_operationId.xlsx).

Default locations (adjust via CLI args if needed):
- Swagger text file: mock_backend_service/inputs/swagger.txt
- Input Excel file:  mock_backend_service/inputs/testcases.xlsx
- Output Excel file: mock_backend_service/inputs/testcases_with_operationId.xlsx

Mapping logic:
- If the testcases sheet already has an 'operationId' column, it will be preserved. Empty cells
  can be auto-filled using naive heuristics.
- If there is no 'operationId' column, a new one will be created to the right of the last column.
- Naive mapping attempts:
  1) If there are 'method' and 'path' columns (case-insensitive), try to match by HTTP method and path
     against the swagger endpoints that have operationIds.
  2) If none matched or columns missing, leave blank or fill generically if desired.

IMPORTANT:
- This script assumes swagger.txt contains JSON OpenAPI/Swagger (2.0 or 3.x) content.
- The Excel is processed via pandas + openpyxl. Only the first sheet is modified by default,
  which you can override by passing --sheet.
- You may need to adjust the mapping logic section marked with "# MAPPING LOGIC" to fit your testcases.xlsx structure.

Requires:
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
def parse_operation_ids_from_swagger_text(swagger_text: str) -> Tuple[Dict[str, Dict[str, Any]], List[str]]:
    """
    PUBLIC_INTERFACE
    Parse an OpenAPI/Swagger JSON string and return:
      - endpoints: mapping of operationId -> {method, path}
      - operation_ids: list of operationIds in discovery order

    Supports Swagger 2.0 and OpenAPI 3.x with a 'paths' object.
    If an operation lacks an operationId, a synthesized one is generated, but these are still returned
    to maintain row-wise fills if necessary.
    """
    data = json.loads(swagger_text)

    # Locate paths (common to both 2.0 and 3.x at root)
    paths = data.get("paths", {})
    endpoints: Dict[str, Dict[str, Any]] = {}
    operation_ids: List[str] = []
    seen = set()

    for method, path, op in _iter_operations_from_paths(paths):
        op_id = (op.get("operationId") or f"{method.lower()}_{path}").replace("/", "_").replace("{", "").replace("}", "")
        if op_id not in seen:
            endpoints[op_id] = {"method": method, "path": path}
            operation_ids.append(op_id)
            seen.add(op_id)

    return endpoints, operation_ids


def _normalize_colnames(cols: List[str]) -> Dict[str, str]:
    """
    Build a case-insensitive mapping of lowercased names to original names.
    """
    mapping: Dict[str, str] = {}
    for c in cols:
        if isinstance(c, str):
            mapping[c.strip().lower()] = c
    return mapping


def _match_by_method_path(row: pd.Series, endpoints: Dict[str, Dict[str, Any]]) -> Optional[str]:
    """
    Try to match a row to an operationId using method+path if both are present.
    This is a naive exact match on method and path.
    """
    cols_map = _normalize_colnames(list(row.index))
    method_col = cols_map.get("method")
    path_col = cols_map.get("path")
    if not method_col or not path_col:
        return None

    method_val = str(row.get(method_col) or "").strip().upper()
    path_val = str(row.get(path_col) or "").strip()

    if not method_val or not path_val:
        return None

    for op_id, info in endpoints.items():
        if info.get("method") == method_val and info.get("path") == path_val:
            return op_id
    return None


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


# PUBLIC_INTERFACE
def process_excel_with_operation_ids(
    swagger_path: str,
    input_excel_path: str,
    output_excel_path: str,
    sheet_name: str = "Standard Template",
) -> None:
    """
    PUBLIC_INTERFACE
    Main routine:
      1) Load swagger.txt (JSON) and extract operationIds + (method, path) mapping.
      2) Load testcases.xlsx and ensure an 'operationId' column exists on the target sheet.
      3) Apply mapping logic to fill operationId per row, else leave blank.
      4) Save to testcases_with_operationId.xlsx.

    Notes:
      - Only the specified sheet is modified; other sheets are preserved unchanged.
      - Mapping logic is intentionally minimal; update the section marked "# MAPPING LOGIC"
        to suit your testcases.xlsx format (e.g., map by endpoint name columns, tags, etc.).
    """
    if not os.path.isfile(swagger_path):
        raise FileNotFoundError(f"Swagger file not found: {swagger_path}")
    if not os.path.isfile(input_excel_path):
        raise FileNotFoundError(f"Input Excel not found: {input_excel_path}")

    swagger_text = _read_text(swagger_path)
    endpoints, operation_ids = parse_operation_ids_from_swagger_text(swagger_text)

    # Read Excel with pandas, preserve all sheets
    xls = pd.ExcelFile(input_excel_path)
    sheets: Dict[str, pd.DataFrame] = {name: xls.parse(name) for name in xls.sheet_names}

    if sheet_name not in sheets:
        # If the desired sheet is missing, just write back unchanged to output
        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, index=False, sheet_name=name)
        return

    df = sheets[sheet_name].copy()

    # Ensure 'operationId' column exists
    df, op_col = _ensure_operationid_column(df)

    # MAPPING LOGIC:
    # Strategy: for each row, if operationId already present -> leave it.
    # Else, if method+path available -> find a matching operationId from swagger.
    # Else, leave blank. Alternatively, you can cycle operation_ids list to fill generically.
    #
    # To customize mapping:
    # - If your Excel has columns like "endpoint_name", "tag", or "summary", implement
    #   matching rules here (string contains, normalization, etc.) to map to an op_id.
    # - If there is a 1:1 order correspondence, you could assign operation_ids sequentially.

    for idx in range(len(df)):
        current_val = df.at[idx, op_col]
        if isinstance(current_val, str) and current_val.strip():
            continue  # preserve existing
        # Try method+path mapping
        matched = _match_by_method_path(df.loc[idx], endpoints)
        if matched:
            df.at[idx, op_col] = matched
        else:
            # Optional fallback: sequential fill from operation_ids (commented by default)
            # seq_index = idx if idx < len(operation_ids) else None
            # df.at[idx, op_col] = operation_ids[seq_index] if seq_index is not None else None
            df.at[idx, op_col] = None  # Leave blank if no mapping found

    # Save back all sheets, replacing only the target one
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            if name == sheet_name:
                df.to_excel(writer, index=False, sheet_name=name)
            else:
                sdf.to_excel(writer, index=False, sheet_name=name)


def _build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read operationIds from swagger.txt and add an 'operationId' column to testcases.xlsx."
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
            f"Success: Added/updated 'operationId' column on sheet '{args.sheet}' "
            f"and saved to '{args.output_excel}'."
        )
        return 0
    except Exception as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
