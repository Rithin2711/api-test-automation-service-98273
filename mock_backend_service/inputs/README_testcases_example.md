# Example testcases.xlsx

This example workbook provides a minimal set of columns for mapping and executing API test cases with the operationId helper scripts.

Columns (first row headers):
- test_case_id: Unique identifier for a test case row (string or integer).
- description: Brief description of the test's intent.
- endpoint: The API path (e.g., /users, /orders/{id}). This is used as the "path" for mapping.
- method: HTTP verb (GET/POST/PUT/DELETE/PATCH).
- expected_result: Freeform expected outcome description (e.g., "200 OK with user list"). For automated status assertions in this project, use the expected_status column described below.
- payload: JSON string for the request body or query params (optional).
- expected_status: Integer HTTP status expected from the call (optional, defaults to 200).
- base_url: Optional override of the base URL to build the full request URL if different from the swagger base.

Notes:
- The operationId scripts look for flexible header names; "endpoint" is treated like "path".
- For execution in this backend, required fields per row are either:
  - operationId, or
  - method + path (endpoint), and optionally payload and expected_status.

This file is intended purely as an example/template.
