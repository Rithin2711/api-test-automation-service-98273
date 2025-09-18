This directory now contains two sample inputs:
- swagger.txt: OpenAPI 2.0 JSON with:
  - GET /api/hello (operationId: getHello) expected 200
  - POST /api/echo (operationId: postEcho) expected 201
- testcases.xlsx: Excel with a sheet "TestCases" and rows matching those endpoints.

The service's parser (parse_swagger_json) supports Swagger/OpenAPI 2.0 JSON, so swagger.txt is JSON content.

If your environment does not auto-decode Base64 files:
1) Decode "testcases.xlsx.base64" to binary as "testcases.xlsx"
2) Remove the .base64 file if not needed

Expected headers (case-insensitive): operationId OR method+path, payload, expected_status, base_url (optional).
