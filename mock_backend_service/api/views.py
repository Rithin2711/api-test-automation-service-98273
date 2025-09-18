from django.utils import timezone
from rest_framework.decorators import api_view, parser_classes
from rest_framework.response import Response
from rest_framework import status
from rest_framework.parsers import MultiPartParser, FormParser
from django.http import HttpResponse

from .utils import execute_testcases
from .serializers_combined import CombinedExecuteUploadSerializer


@api_view(['GET'])
def health(request):
    """
    PUBLIC_INTERFACE
    Simple health check endpoint.
    Returns JSON {"message": "Server is up!"}
    """
    return Response({"message": "Server is up!"})


# PUBLIC_INTERFACE
@api_view(['POST'])
@parser_classes([MultiPartParser, FormParser])
def execute_combined(request):
    """
    PUBLIC_INTERFACE
    summary: Execute tests with uploaded swagger and testcases in one request
    description: |
      Accepts multipart/form-data containing:
        - swagger: swagger.json or swagger.txt (OpenAPI/Swagger JSON text)
        - testcases: Excel .xlsx file
      Workflow:
        1) Parse the swagger file to resolve base URL and endpoints.
        2) Read the Excel file (sheet "TestCases" or first sheet).
        3) For each row, build and send an HTTP request based on operationId or method+path and payload.
        4) Compare actual status with expected_status from the row.
        5) If data is missing/invalid (e.g., no method/opId/path, invalid expected_status), mark as "Data Insufficient".
        6) Add/Update the "Status" column with Pass/Fail/Data Insufficient.
        7) Return the modified Excel (.xlsx) as the response for download.
      Notes:
        - This endpoint does not store files; it processes the upload and returns the result directly.

    requestBody:
      multipart/form-data:
        fields:
          swagger: file (.json or .txt)
          testcases: file (.xlsx)

    responses:
      200:
        description: Returns the updated Excel file with a "Status" column
        content:
          application/vnd.openxmlformats-officedocument.spreadsheetml.sheet: binary
      400:
        description: Validation or parsing error
    tags:
      - execute
    """
    serializer = CombinedExecuteUploadSerializer(data=request.data)
    if not serializer.is_valid():
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    swagger_file = serializer.validated_data["swagger"]
    testcases_file = serializer.validated_data["testcases"]

    try:
        result_bytes, total, passed, failed = execute_testcases(swagger_file.read(), testcases_file.read())
    except Exception as exc:
        return Response({"detail": f"Failed to execute: {exc}"}, status=status.HTTP_400_BAD_REQUEST)

    filename = f"results_{timezone.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    resp = HttpResponse(
        result_bytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    resp["X-Execution-Summary"] = f"total={total};passed={passed};failed={failed}"
    return resp
