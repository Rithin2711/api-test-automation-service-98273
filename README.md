# Project Repository

This repository contains a Django backend service that:
- Uploads swagger.json and an Excel of test cases
- Parses swagger to discover endpoints
- Executes requests using rows from the Excel
- Writes Pass/Fail status back to a result Excel
- Serves the result Excel for download

Basic flow:
1) POST /api/upload/swagger/ with form-data file=swagger.json
2) POST /api/upload/testcases/ with form-data file=tests.xlsx (sheet 'TestCases' or the first sheet)
   Required headers (case-insensitive): method or operationId, path (when method used), payload (JSON), expected_status (optional)
3) POST /api/execute/ to run with latest files, or include {"swagger_id": X, "testcase_id": Y}
4) GET /api/results/{file_id}/ to download the generated Excel
5) GET /api/files/ to list uploaded and result files

Standalone utility: Extract endpoints and payload schemas from OpenAPI
- Location: mock_backend_service/scripts/extract_openapi_payloads.py
- Usage:
  python mock_backend_service/scripts/extract_openapi_payloads.py mock_backend_service/interfaces/openapi.json
  python mock_backend_service/scripts/extract_openapi_payloads.py /path/to/swagger.json

This prints a human-readable list of each path, HTTP method, and the request payload schema (if any). It supports OpenAPI 3.0+ and attempts to gracefully handle missing payloads or legacy parameter styles.

Standalone utility: Execute API tests from Excel and write result statuses
- Location: mock_backend_service/scripts/run_api_tests_from_excel.py
- Usage:
  python mock_backend_service/scripts/run_api_tests_from_excel.py \
      --input path/to/tests.xlsx \
      --sheet TestCases \
      --endpoint-col endpoint \
      --method-col method \
      --payload-col payload \
      --expected-status-col expected_status \
      --base-url http://localhost:8000

The script reads an Excel file, sends HTTP requests per row, and writes a 'status' column with pass/fail/data insufficient, saving to a new file with '_results' appended to the filename.