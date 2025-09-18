#!/usr/bin/env python3
"""
PUBLIC_INTERFACE
OpenAPI (swagger.json) endpoint and payload extractor.

This standalone script reads an OpenAPI 3.0+ JSON file (commonly called swagger.json)
and prints a structured, human-readable listing of:
- endpoint paths
- HTTP methods
- request payload schemas (from requestBody content or from body-like parameters)

Usage:
  python extract_openapi_payloads.py /path/to/swagger.json

Notes:
- If an endpoint has no payload (e.g., GET with query/path params only), it is still listed,
  and payload is shown as "None".
- The script attempts to resolve $ref references in components/schemas and requestBodies.
- For parameter-based payloads (rare in OpenAPI 3.x; more common in 2.0), the script will
  collect "in: body" or "in: formData" style structures when present for compatibility.
- The script outputs to stdout; you can redirect to a file if needed.

Author: Automation Code Writer
"""
import json
import sys
import os
from typing import Any, Dict, Optional, Union


# Types
JSON = Dict[str, Any]


# PUBLIC_INTERFACE
def load_openapi(path: str) -> JSON:
    """Load and return the OpenAPI JSON from the given file path."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _dict_get(d: Optional[dict], *keys, default=None):
    cur = d or {}
    for k in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(k)
        if cur is None:
            return default
    return cur


def _resolve_ref(doc: JSON, ref: str) -> Optional[JSON]:
    """
    Resolve a JSON reference like '#/components/schemas/MySchema'.
    Returns the resolved object or None if not found.
    """
    if not ref or not isinstance(ref, str):
        return None
    if not ref.startswith("#/"):
        # External refs not supported in this simple script.
        return None
    parts = ref.lstrip("#/").split("/")
    cur: Union[JSON, list, Any] = doc
    for p in parts:
        if isinstance(cur, dict) and p in cur:
            cur = cur[p]
        else:
            return None
    if isinstance(cur, dict):
        return cur
    # If not a dict (e.g., primitive), wrap it or return raw
    return {"value": cur}


def _extract_schema_from_content(doc: JSON, content: JSON) -> Optional[JSON]:
    """
    Extract a schema from a requestBody.content map.
    Prefer application/json, then any other available content type.
    Resolve $ref when present.
    """
    if not isinstance(content, dict):
        return None
    # Prefer JSON-like content types
    preferred = [
        "application/json",
        "application/ld+json",
        "application/*+json",
        "text/json",
    ]
    candidate_media_types = list(content.keys())

    # Choose preferred type if available
    chosen_type = None
    for p in preferred:
        # exact and wildcard match
        for mt in candidate_media_types:
            if mt == p or (p.endswith("/*+json") and mt.endswith("+json")):
                chosen_type = mt
                break
        if chosen_type:
            break
    # Fallback: pick first
    if not chosen_type and candidate_media_types:
        chosen_type = candidate_media_types[0]

    if not chosen_type:
        return None

    mt_obj = content.get(chosen_type) or {}
    schema = mt_obj.get("schema")
    if not schema:
        return None

    # Resolve refs
    if "$ref" in schema:
        resolved = _resolve_ref(doc, schema["$ref"])
        return resolved or schema
    return schema


def _extract_payload_from_parameters(doc: JSON, params: Any) -> Optional[JSON]:
    """
    For compatibility with specs that still use parameters as body (legacy or mixed),
    try to synthesize a payload schema from parameters.

    - Collect 'in: body' schema (OpenAPI 2 style) if present.
    - Collect 'in: formData' (OpenAPI 2 style) as an object with those fields.
    - Otherwise return None.
    """
    if not isinstance(params, list):
        return None

    # Try body parameter (OpenAPI 2.0)
    for p in params:
        if not isinstance(p, dict):
            continue
        # Resolve param $ref if needed
        if "$ref" in p:
            resolved = _resolve_ref(doc, p["$ref"])
            if isinstance(resolved, dict):
                p = resolved

        if p.get("in") == "body":
            schema = p.get("schema")
            if schema:
                if "$ref" in schema:
                    return _resolve_ref(doc, schema["$ref"]) or schema
                return schema

    # Try formData (OpenAPI 2.0)
    form_props = {}
    required_fields = []
    for p in params:
        if not isinstance(p, dict):
            continue
        if "$ref" in p:
            resolved = _resolve_ref(doc, p["$ref"])
            if isinstance(resolved, dict):
                p = resolved
        if p.get("in") in ("formData", "form"):
            name = p.get("name") or "field"
            schema: JSON = {}
            # Map simple type info
            typ = p.get("type")
            fmt = p.get("format")
            if typ:
                schema["type"] = typ
            if fmt:
                schema["format"] = fmt
            form_props[name] = schema or {"type": "string"}
            if p.get("required"):
                required_fields.append(name)

    if form_props:
        obj: JSON = {"type": "object", "properties": form_props}
        if required_fields:
            obj["required"] = required_fields
        return obj

    return None


def _pretty_print_schema(schema: Any, indent: int = 0, max_depth: int = 5):
    """
    Pretty print a JSON schema dict in a simple readable way, limiting depth to avoid verbosity.
    """
    prefix = " " * indent
    if schema is None:
        print(f"{prefix}None")
        return
    if not isinstance(schema, dict):
        print(f"{prefix}{schema}")
        return

    # Limit depth
    if indent // 2 >= max_depth:
        compact = json.dumps(schema, ensure_ascii=False)[:300]
        print(f"{prefix}{compact} ...")
        return

    typ = schema.get("type")
    if "$ref" in schema:
        print(f"{prefix}$ref -> {schema['$ref']}")
        return
    if typ == "object":
        print(f"{prefix}type: object")
        props = schema.get("properties") or {}
        required = schema.get("required") or []
        if props:
            print(f"{prefix}properties:")
            for k, v in props.items():
                req_mark = " (required)" if k in required else ""
                print(f"{prefix}  {k}:{req_mark}")
                _pretty_print_schema(v, indent + 4, max_depth)
        else:
            print(f"{prefix}properties: {{}}")
    elif typ == "array":
        print(f"{prefix}type: array")
        items = schema.get("items")
        print(f"{prefix}items:")
        _pretty_print_schema(items, indent + 2, max_depth)
    else:
        # primitive or unknown
        keys_to_show = ("type", "format", "enum", "description", "default")
        shown = {k: schema[k] for k in keys_to_show if k in schema}
        if shown:
            for k, v in shown.items():
                print(f"{prefix}{k}: {v}")
        else:
            compact = json.dumps(schema, ensure_ascii=False)[:300]
            print(f"{prefix}{compact}")


# PUBLIC_INTERFACE
def extract_endpoints_and_payloads(doc: JSON) -> Dict[str, Dict[str, Any]]:
    """
    PUBLIC_INTERFACE
    Extract endpoint paths, methods, and payload schema from an OpenAPI 3.0+ document.

    Returns:
      A nested dict:
      {
        "<path>": {
            "<METHOD>": {
                "summary": str|None,
                "operationId": str|None,
                "payload_schema": dict|None
            },
            ...
        },
        ...
      }
    """
    out: Dict[str, Dict[str, Any]] = {}

    paths = doc.get("paths") or {}
    for path, path_item in paths.items():
        if not isinstance(path_item, dict):
            continue

        # Check for path-level parameters (fallback/legacy)
        path_level_params = path_item.get("parameters")

        for method in ("get", "post", "put", "patch", "delete", "options", "head", "trace"):
            op = path_item.get(method)
            if not isinstance(op, dict):
                continue

            entry: Dict[str, Any] = {
                "summary": op.get("summary") or op.get("description"),
                "operationId": op.get("operationId"),
                "payload_schema": None,
            }

            # 1) requestBody (OpenAPI 3.x)
            request_body = op.get("requestBody")
            schema_from_body = None
            if isinstance(request_body, dict):
                # Handle $ref on requestBody
                if "$ref" in request_body:
                    resolved_rb = _resolve_ref(doc, request_body["$ref"]) or {}
                else:
                    resolved_rb = request_body
                schema_from_body = _extract_schema_from_content(doc, _dict_get(resolved_rb, "content"))

            # 2) parameters: attempt to synthesize (legacy/mixed or if no requestBody)
            schema_from_params = None
            params = op.get("parameters")
            if not params and path_level_params:
                params = path_level_params
            if params:
                schema_from_params = _extract_payload_from_parameters(doc, params)

            entry["payload_schema"] = schema_from_body or schema_from_params

            # Attach to out
            out.setdefault(path, {})[method.upper()] = entry

    return out


# PUBLIC_INTERFACE
def print_human_readable(result: Dict[str, Dict[str, Any]]) -> None:
    """
    PUBLIC_INTERFACE
    Print a readable representation of the extracted endpoints and payload schemas.
    """
    if not result:
        print("No endpoints found.")
        return

    for path, methods in sorted(result.items(), key=lambda x: x[0]):
        print(f"Path: {path}")
        for method, info in sorted(methods.items(), key=lambda x: x[0]):
            print(f"  Method: {method}")
            if info.get("operationId"):
                print(f"    operationId: {info['operationId']}")
            if info.get("summary"):
                print(f"    summary: {info['summary']}")
            print("    payload schema:")
            _pretty_print_schema(info.get("payload_schema"), indent=6)
        print("")


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Usage: python extract_openapi_payloads.py /path/to/swagger.json")
        return 1

    path = argv[1]
    if not os.path.exists(path):
        print(f"Error: File not found: {path}")
        return 2

    try:
        doc = load_openapi(path)
    except Exception as e:
        print(f"Error: Failed to parse JSON: {e}")
        return 3

    result = extract_endpoints_and_payloads(doc)
    print_human_readable(result)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
