#!/usr/bin/env python3
"""
PUBLIC_INTERFACE
CLI Utility: Extract operationIds from an OpenAPI/Swagger specification.

Usage:
  python extract_operation_ids.py <path_to_openapi_file>

Features:
- Supports JSON (.json) and YAML (.yaml/.yml) formats.
- Handles both OpenAPI 3.x and Swagger (OpenAPI 2.0) structures.
- Prints each operationId with its HTTP method and path.
- If operationId is missing, prints a synthesized one (method_path) and marks it as [no operationId].

Notes:
- No environment variables required.
"""

import argparse
import json
import os
import sys
from typing import Any, Dict, Iterable, Tuple

try:
    import yaml  # PyYAML
except Exception:
    yaml = None  # Will error nicely if YAML is needed but package missing


# PUBLIC_INTERFACE
def load_spec(filepath: str) -> Dict[str, Any]:
    """Load an OpenAPI/Swagger spec from a JSON or YAML file.

    The loader attempts to parse based on file extension first, then tries JSON,
    then YAML as a fallback.

    Args:
        filepath: Path to the spec file (.json, .yaml, .yml)

    Returns:
        Parsed spec as a Python dict.

    Raises:
        RuntimeError: If the file cannot be read or parsed.
    """
    if not os.path.isfile(filepath):
        raise RuntimeError(f"File not found: {filepath}")

    with open(filepath, "rb") as f:
        raw = f.read()

    text = raw.decode("utf-8", errors="replace").strip()
    ext = os.path.splitext(filepath)[1].lower()

    # Try based on extension
    if ext == ".json":
        try:
            return json.loads(text)
        except Exception as e:
            raise RuntimeError(f"Failed to parse JSON: {e}") from e
    elif ext in (".yaml", ".yml"):
        if yaml is None:
            raise RuntimeError("PyYAML is required to parse YAML. Please install PyYAML.")
        try:
            return yaml.safe_load(text)
        except Exception as e:
            raise RuntimeError(f"Failed to parse YAML: {e}") from e

    # Fallback: try JSON first, then YAML if available
    try:
        return json.loads(text)
    except Exception:
        if yaml is None:
            raise RuntimeError(
                "Unknown file extension and JSON parse failed. "
                "Install PyYAML or provide a valid .json file."
            )
        try:
            return yaml.safe_load(text)
        except Exception as e:
            raise RuntimeError(f"Failed to parse file as JSON or YAML: {e}") from e


def _iter_operations_from_paths(paths: Dict[str, Any]) -> Iterable[Tuple[str, str, Dict[str, Any]]]:
    """Iterate over operations from a 'paths' object.

    Yields:
        (method_upper, path, operation_object_dict)
    """
    if not isinstance(paths, dict):
        return

    http_methods = {"get", "post", "put", "patch", "delete", "head", "options", "trace"}
    for path, path_item in paths.items():
        if not isinstance(path_item, dict):
            continue
        for method, op_obj in path_item.items():
            if method.lower() in http_methods and isinstance(op_obj, dict):
                yield method.upper(), path, op_obj


# PUBLIC_INTERFACE
def extract_operations(spec: Dict[str, Any]) -> Iterable[Tuple[str, str, str]]:
    """Extract (operationId, method, path) triples from an OpenAPI/Swagger spec.

    Supports:
      - Swagger 2.0 (OpenAPI 2.0): spec['swagger'] == '2.0'
      - OpenAPI 3.x: spec['openapi'] starts with '3'

    Args:
        spec: Parsed spec dict.

    Yields:
        Tuples of (operation_id_or_synthesized, method_upper, path)
    """
    # Determine version and locate paths
    paths = None
    if isinstance(spec, dict):
        if "paths" in spec and isinstance(spec["paths"], dict):
            paths = spec["paths"]
        elif "components" in spec:
            # Still may have 'paths' at root; nothing else to do.
            paths = spec.get("paths")

    if not paths:
        return

    for method, path, op_obj in _iter_operations_from_paths(paths):
        op_id = op_obj.get("operationId")
        if not op_id:
            # Synthesize a readable operationId when missing
            synthesized = f"{method.lower()}_{path}".replace("/", "_").replace("{", "").replace("}", "")
            op_id = synthesized
            yield f"{op_id} [no operationId]", method, path
        else:
            yield op_id, method, path


def _print_results(ops: Iterable[Tuple[str, str, str]]) -> int:
    """Print operations and return count."""
    count = 0
    print("Discovered operations:")
    print("----------------------")
    for op_id, method, path in ops:
        count += 1
        print(f"- operationId: {op_id}\n  method: {method}\n  path:   {path}\n")
    if count == 0:
        print("No operations found.")
    else:
        print(f"Total: {count} operation(s).")
    return count


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Extract and list operationIds with methods and paths from an OpenAPI/Swagger file (JSON/YAML)."
    )
    parser.add_argument(
        "file",
        help="Path to the OpenAPI/Swagger file (.json, .yaml, .yml)"
    )
    args = parser.parse_args(argv[1:])

    try:
        spec = load_spec(args.file)
        ops = list(extract_operations(spec))
        _print_results(ops)
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
