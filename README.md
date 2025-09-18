# Project Repository

This Django backend exposes a single combined endpoint to:
- Accept a swagger.json (or .txt with JSON) and a test cases Excel (.xlsx) in one request.
- Parse swagger to discover endpoints and base URL.
- Execute requests for each row from the Excel.
- Compare actual vs expected status and write results into a new "Status" column.
- Return the modified Excel for download.

Endpoint:
- POST /api/execute/combined/ (multipart/form-data)
  Fields (case-insensitive names in Excel rows are supported):
    - swagger: swagger.json or swagger.txt (JSON content)
    - testcases: tests.xlsx (sheet "TestCases" or the first sheet)
  Excel required columns:
    - operationId OR (method + path)
    - payload (JSON string) [optional]
    - expected_status (default 200 if valid integer)
    - base_url (optional override)
  Status values written:
    - Pass: actual HTTP status equals expected_status
    - Fail: actual HTTP status differs or request error with sufficient data
    - Data Insufficient: missing/invalid required fields (e.g., missing method/opId/path or invalid expected_status)

Health:
- GET /api/health/ -> {"message": "Server is up!"}

Notes:
- No files are persisted; the endpoint processes uploads and returns the result .xlsx.
- Swagger must be OpenAPI/Swagger 2.0 JSON.