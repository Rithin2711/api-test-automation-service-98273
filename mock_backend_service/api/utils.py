import io
import json
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple
import requests
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class EndpointSpec:
    """Represents an endpoint discovered in swagger/openapi."""
    method: str
    path: str
    consumes: Optional[List[str]] = None
    produces: Optional[List[str]] = None
    base_url: Optional[str] = None


# PUBLIC_INTERFACE
def parse_swagger_json(swagger_bytes: bytes) -> Dict[str, EndpointSpec]:
    """
    PUBLIC_INTERFACE
    Parse a swagger.json (OpenAPI 2.0) file and return a mapping of operationId (or method+path)
    to EndpointSpec objects. This implementation handles basic swagger 2.0 structures.
    """
    data = json.loads(swagger_bytes.decode("utf-8"))
    host = data.get("host", "")
    schemes = data.get("schemes", ["http"])
    base_path = data.get("basePath", "")
    base_url = None
    if host:
        scheme = schemes[0] if schemes else "http"
        base_url = f"{scheme}://{host}{base_path}"
    elif base_path:
        # If only basePath is provided, use relative base
        base_url = base_path
    else:
        base_url = ""

    out: Dict[str, EndpointSpec] = {}
    paths = data.get("paths", {})
    global_consumes = data.get("consumes")
    global_produces = data.get("produces")

    for path, methods in paths.items():
        for method, op in methods.items():
            if method.lower() not in ("get", "post", "put", "patch", "delete", "options", "head"):
                continue
            op_id = op.get("operationId") or f"{method.lower()}_{path}".replace("/", "_")
            consumes = op.get("consumes", global_consumes)
            produces = op.get("produces", global_produces)
            out[op_id] = EndpointSpec(
                method=method.upper(),
                path=path,
                consumes=consumes,
                produces=produces,
                base_url=base_url,
            )
    return out


def _normalize_url(base_url: str, path: str) -> str:
    if base_url.endswith("/"):
        base_url = base_url[:-1]
    if not path.startswith("/"):
        path = "/" + path
    return f"{base_url}{path}"


def _parse_row_dict(row_values: Dict[str, Any]) -> Tuple[str, str, Dict[str, Any]]:
    """
    Expect at least columns: method, path or operationId, payload (JSON as string).
    If base_url is provided in a column, it overrides swagger host/basePath.
    """
    operation_id = str(row_values.get("operationid") or row_values.get("operation_id") or "").strip()
    method = str(row_values.get("method") or "").strip().upper()
    path = str(row_values.get("path") or "").strip()
    payload_text = row_values.get("payload") or row_values.get("body") or "{}"
    try:
        payload = json.loads(payload_text) if isinstance(payload_text, str) else payload_text
    except Exception:
        payload = {}
    base_url_override = str(row_values.get("base_url") or "").strip()
    expected_status = int(row_values.get("expected_status") or 200)
    return operation_id, base_url_override, {"method": method, "path": path, "payload": payload, "expected_status": expected_status}


# PUBLIC_INTERFACE
def execute_testcases(swagger_bytes: bytes, testcase_bytes: bytes) -> Tuple[bytes, int, int, int]:
    """
    PUBLIC_INTERFACE
    Execute test cases from an Excel against endpoints defined in swagger.
    The Excel must contain a worksheet named 'TestCases' or the first sheet will be used.
    Required columns (case-insensitive headers):
      - method (e.g., GET, POST) or operationId (to match swagger operationId)
      - path (e.g., /health/) when method is specified (ignored if operationId uniquely maps)
      - payload (JSON string) for body where applicable
      - expected_status (optional, default 200)
      - base_url (optional to override swagger base)

    Adds/Updates a 'Status' column with 'Pass' or 'Fail' based on HTTP response status match.
    Returns tuple of (result_excel_bytes, total, passed, failed)
    """
    specs = parse_swagger_json(swagger_bytes)

    wb: Workbook = load_workbook(io.BytesIO(testcase_bytes))
    ws: Worksheet = wb["TestCases"] if "TestCases" in wb.sheetnames else wb[wb.sheetnames[0]]

    # Build header map (case-insensitive)
    headers: Dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        header = str(ws.cell(row=1, column=col).value or "").strip()
        if header:
            headers[header.lower()] = col

    # Ensure Status column exists
    status_col = headers.get("status")
    if not status_col:
        status_col = ws.max_column + 1
        ws.cell(row=1, column=status_col, value="Status")
        headers["status"] = status_col

    total = 0
    passed = 0
    failed = 0

    for row in range(2, ws.max_row + 1):
        total += 1
        row_values = {h: ws.cell(row=1, column=c).value for h, c in headers.items() if h != "status"}
        # pull row value per header
        for h, c in headers.items():
            if h == "status":
                continue
            row_values[h] = ws.cell(row=row, column=c).value

        operation_id, base_url_override, parsed = _parse_row_dict(row_values)
        method = parsed["method"]
        path = parsed["path"]
        payload = parsed["payload"]
        expected_status = parsed["expected_status"]

        url: Optional[str] = None
        if operation_id:
            spec = specs.get(operation_id)
            if spec:
                url = _normalize_url(base_url_override or spec.base_url or "", spec.path)
                method = spec.method
        else:
            # fall back to method+path
            # try to find by method+path match
            found_spec = None
            for sp in specs.values():
                if sp.method == method and sp.path == path:
                    found_spec = sp
                    break
            if found_spec:
                url = _normalize_url(base_url_override or found_spec.base_url or "", found_spec.path)
            else:
                url = _normalize_url(base_url_override or "", path)

        status_value = "Fail"
        try:
            # Send request
            if method == "GET":
                resp = requests.get(url, params=payload if isinstance(payload, dict) else None, timeout=15)
            elif method in ("POST", "PUT", "PATCH", "DELETE"):
                resp = requests.request(method, url, json=payload, timeout=15)
            else:
                resp = requests.request(method, url, timeout=15)

            if resp.status_code == expected_status:
                status_value = "Pass"
                passed += 1
            else:
                failed += 1
        except Exception:
            failed += 1
            status_value = "Fail"

        ws.cell(row=row, column=status_col, value=status_value)

    # Save workbook back to bytes
    out_stream = io.BytesIO()
    wb.save(out_stream)
    out_stream.seek(0)
    return out_stream.read(), total, passed, failed
