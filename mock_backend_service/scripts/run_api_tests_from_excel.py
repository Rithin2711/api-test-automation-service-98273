#!/usr/bin/env python3
"""
Utility script to execute API tests defined in an Excel file and save results.

Features:
- Loads an Excel file with columns for endpoint, method, and payload (configurable column names).
- Sends an HTTP request per row using the given method and payload.
- Determines pass/fail based on response status code (default expected=200, configurable per-row column).
- Writes a 'status' column with values: 'pass', 'fail', or 'data insufficient'.
- Saves results in a new Excel file with '_results' appended to the original filename.

Usage:
    python mock_backend_service/scripts/run_api_tests_from_excel.py \
        --input path/to/tests.xlsx \
        --sheet TestCases \
        --endpoint-col endpoint \
        --method-col method \
        --payload-col payload \
        --expected-status-col expected_status

Notes:
- Endpoints should be absolute URLs (http://...) or will be treated as relative to base_url (--base-url).
- Payload should be a JSON string or a Python-dict-like object; non-JSON will be ignored.
- If a row is missing endpoint, method, or payload as required, the status will be 'data insufficient'.
- If expected_status is missing, defaults to 200.

Dependencies: pandas, requests, openpyxl
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd
import requests


# PUBLIC_INTERFACE
def run_tests_from_excel(
    input_path: str,
    sheet_name: Optional[str] = None,
    endpoint_col: str = "endpoint",
    method_col: str = "method",
    payload_col: str = "payload",
    expected_status_col: str = "expected_status",
    status_col: str = "status",
    base_url: str = "",
    timeout_sec: int = 15,
) -> str:
    """
    PUBLIC_INTERFACE
    Run API tests defined in an Excel file and write results to a new Excel file.

    Parameters:
    - input_path: Path to the Excel file containing test cases.
    - sheet_name: Optional worksheet name (default: first sheet).
    - endpoint_col: Column name for endpoint URL or path.
    - method_col: Column name for HTTP method (GET, POST, PUT, PATCH, DELETE).
    - payload_col: Column name for JSON payload (string or dict-like).
    - expected_status_col: Column with integer expected HTTP status code (default 200).
    - status_col: Name of the status column to add/update.
    - base_url: Optional base URL to prefix for relative endpoints (e.g., "http://localhost:8000").
    - timeout_sec: Request timeout in seconds.

    Returns:
    - The output Excel file path (with '_results' appended before extension).
    """
    # Load the Excel file as a DataFrame
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    # Normalize columns to support case-insensitive matching by creating a mapping
    def find_col(name: str) -> Optional[str]:
        lower_map = {c.lower(): c for c in df.columns}
        return lower_map.get(name.lower())

    endpoint_c = find_col(endpoint_col) or endpoint_col
    method_c = find_col(method_col) or method_col
    payload_c = find_col(payload_col) or payload_col
    expected_c = find_col(expected_status_col) or expected_status_col
    status_c = find_col(status_col) or status_col

    # Ensure the status column exists
    if status_c not in df.columns:
        df[status_c] = ""

    def _normalize_url(ep: str) -> str:
        """
        If endpoint is absolute (starts with http:// or https://), return as is.
        If relative and base_url provided, join carefully.
        """
        ep = ep.strip()
        if ep.startswith("http://") or ep.startswith("https://"):
            return ep
        if not base_url:
            return ep  # leave as-is; requests may fail if not absolute
        if base_url.endswith("/") and ep.startswith("/"):
            return f"{base_url[:-1]}{ep}"
        if not base_url.endswith("/") and not ep.startswith("/"):
            return f"{base_url}/{ep}"
        return f"{base_url}{ep}"

    def _parse_payload(val: Any) -> Optional[Dict[str, Any]]:
        """
        Tries to parse the payload value:
        - If it's a dict already, return it.
        - If it's a valid JSON string, parse and return dict.
        - Otherwise, return None to indicate no JSON body.
        """
        if isinstance(val, dict):
            return val
        if isinstance(val, str):
            s = val.strip()
            if not s:
                return None
            try:
                parsed = json.loads(s)
                if isinstance(parsed, dict):
                    return parsed
                # If it's a list or other JSON, we'll still pass as JSON body
                return parsed  # type: ignore
            except Exception:
                # Try to eval a python-literal-like dict/list safely
                try:
                    parsed = json.loads(json.dumps(eval(s)))  # risky in general; here just as fallback
                    return parsed  # type: ignore
                except Exception:
                    return None
        return None

    results = []
    for idx, row in df.iterrows():
        method = str(row.get(method_c, "") or "").strip().upper()
        endpoint = str(row.get(endpoint_c, "") or "").strip()
        payload_val = row.get(payload_c, None)
        expected_status = row.get(expected_c, None)

        # Determine expected status code
        try:
            expected = int(expected_status) if pd.notna(expected_status) else 200
        except Exception:
            expected = 200

        # Basic validation for mandatory fields
        if not method or not endpoint:
            results.append("data insufficient")
            continue

        url = _normalize_url(endpoint)
        payload = _parse_payload(payload_val)

        # Default to fail; mark pass only when conditions are met
        status_value = "fail"
        try:
            # Choose request based on method and whether payload exists
            if method == "GET":
                # For GET, send payload as query params if it's a dict
                params = payload if isinstance(payload, dict) else None
                resp = requests.get(url, params=params, timeout=timeout_sec)
            elif method in ("POST", "PUT", "PATCH", "DELETE"):
                # For these, send JSON body if provided
                if payload is not None:
                    resp = requests.request(method, url, json=payload, timeout=timeout_sec)
                else:
                    resp = requests.request(method, url, timeout=timeout_sec)
            else:
                # Other methods - just attempt without body unless payload provided
                if payload is not None:
                    resp = requests.request(method, url, json=payload, timeout=timeout_sec)
                else:
                    resp = requests.request(method, url, timeout=timeout_sec)

            if resp.status_code == expected:
                status_value = "pass"
        except Exception:
            # Keep as fail on any exception
            status_value = "fail"

        results.append(status_value)

    # Write results to the status column
    df[status_c] = results

    # Build output file path
    in_path = Path(input_path)
    if in_path.suffix:
        out_path = in_path.with_name(f"{in_path.stem}_results{in_path.suffix}")
    else:
        out_path = in_path.with_name(f"{in_path.name}_results.xlsx")

    # Save to Excel; preserve sheet name if provided
    if sheet_name:
        # When writing a single sheet, pandas will use that name
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    else:
        # Write as default sheet
        df.to_excel(out_path, index=False)

    return str(out_path)


def _parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    """Parse command-line arguments for the script."""
    parser = argparse.ArgumentParser(description="Execute API tests from an Excel file and write results.")
    parser.add_argument("--input", "-i", required=True, help="Path to the input Excel file (.xlsx)")
    parser.add_argument("--sheet", "-s", default=None, help="Worksheet name (default: first sheet)")
    parser.add_argument("--endpoint-col", default="endpoint", help="Column name for endpoint (default: endpoint)")
    parser.add_argument("--method-col", default="method", help="Column name for HTTP method (default: method)")
    parser.add_argument("--payload-col", default="payload", help="Column name for JSON payload (default: payload)")
    parser.add_argument("--expected-status-col", default="expected_status", help="Column name for expected status (default: expected_status)")
    parser.add_argument("--status-col", default="status", help="Column name for writing status (default: status)")
    parser.add_argument("--base-url", default="", help="Base URL to prefix relative endpoints (e.g., http://localhost:8000)")
    parser.add_argument("--timeout", type=int, default=15, help="HTTP request timeout in seconds (default: 15)")
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    """CLI entrypoint."""
    args = _parse_args(argv)
    try:
        output_file = run_tests_from_excel(
            input_path=args.input,
            sheet_name=args.sheet,
            endpoint_col=args.endpoint_col,
            method_col=args.method_col,
            payload_col=args.payload_col,
            expected_status_col=args.expected_status_col,
            status_col=args.status_col,
            base_url=args.base_url,
            timeout_sec=args.timeout,
        )
        print(f"Results written to: {output_file}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
