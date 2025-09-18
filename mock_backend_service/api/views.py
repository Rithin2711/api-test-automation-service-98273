import io
from typing import Optional

from django.http import FileResponse, Http404, HttpResponse
from django.utils import timezone
from rest_framework.decorators import api_view, parser_classes
from rest_framework.response import Response
from rest_framework import status
from rest_framework.parsers import MultiPartParser, FormParser

from .models import StoredFile, ExecutionRun
from .serializers import (
    SwaggerUploadSerializer,
    TestCaseUploadSerializer,
    ExecuteRequestSerializer,
    StoredFileInfoSerializer,
)
from .utils import execute_testcases
from .serializers_combined import CombinedExecuteUploadSerializer


@api_view(['GET'])
def health(request):
    """Simple health check."""
    return Response({"message": "Server is up!"})


# PUBLIC_INTERFACE
@api_view(['POST'])
def upload_swagger(request):
    """
    PUBLIC_INTERFACE
    summary: Upload swagger.json
    description: |
      Upload a swagger/openapi JSON file that describes the API endpoints.
    responses:
      201: Swagger stored successfully
      400: Validation error
    """
    serializer = SwaggerUploadSerializer(data=request.data)
    if not serializer.is_valid():
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    f = serializer.validated_data["file"]
    content = f.read()
    stored = StoredFile.objects.create(
        kind=StoredFile.KIND_SWAGGER,
        filename=f.name,
        content=content,
        size_bytes=len(content),
    )
    data = {
        "id": stored.id,
        "kind": stored.kind,
        "filename": stored.filename,
        "uploaded_at": stored.uploaded_at,
        "size_bytes": stored.size_bytes,
    }
    return Response(data, status=status.HTTP_201_CREATED)


# PUBLIC_INTERFACE
@api_view(['POST'])
def upload_testcases(request):
    """
    PUBLIC_INTERFACE
    summary: Upload test cases Excel
    description: |
      Upload an Excel (.xlsx) file containing a worksheet with API test cases.
    responses:
      201: Test cases stored successfully
      400: Validation error
    """
    serializer = TestCaseUploadSerializer(data=request.data)
    if not serializer.is_valid():
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    f = serializer.validated_data["file"]
    content = f.read()
    stored = StoredFile.objects.create(
        kind=StoredFile.KIND_TESTCASE,
        filename=f.name,
        content=content,
        size_bytes=len(content),
    )
    data = {
        "id": stored.id,
        "kind": stored.kind,
        "filename": stored.filename,
        "uploaded_at": stored.uploaded_at,
        "size_bytes": stored.size_bytes,
    }
    return Response(data, status=status.HTTP_201_CREATED)


# PUBLIC_INTERFACE
@api_view(['GET'])
def list_files(request):
    """
    PUBLIC_INTERFACE
    summary: List uploaded files
    description: Return metadata of stored files.
    responses:
      200: List of files
    """
    items = StoredFile.objects.all().values("id", "kind", "filename", "uploaded_at", "size_bytes")
    serializer = StoredFileInfoSerializer(items, many=True)
    return Response(serializer.data)


def _get_latest_or_specific(model_qs, specific_id: Optional[int]):
    if specific_id:
        return model_qs.filter(id=specific_id).first()
    return model_qs.order_by("-uploaded_at").first()


# PUBLIC_INTERFACE
@api_view(['POST'])
def execute(request):
    """
    PUBLIC_INTERFACE
    summary: Execute test cases
    description: |
      Executes the most recent uploaded test cases against the most recent uploaded swagger,
      unless specific IDs are provided in the payload.
    requestBody:
      application/json: { swagger_id?: number, testcase_id?: number }
    responses:
      200: Execution summary with run_id and result file id
      400: Missing files or validation error
    """
    serializer = ExecuteRequestSerializer(data=request.data or {})
    if not serializer.is_valid():
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    swagger_id = serializer.validated_data.get("swagger_id")
    testcase_id = serializer.validated_data.get("testcase_id")

    swagger_obj = _get_latest_or_specific(StoredFile.objects.filter(kind=StoredFile.KIND_SWAGGER), swagger_id)
    testcase_obj = _get_latest_or_specific(StoredFile.objects.filter(kind=StoredFile.KIND_TESTCASE), testcase_id)

    if not swagger_obj or not testcase_obj:
        return Response(
            {"detail": "Swagger and TestCase files are required. Upload them first or specify IDs."},
            status=status.HTTP_400_BAD_REQUEST,
        )

    result_bytes, total, passed, failed = execute_testcases(swagger_obj.content, testcase_obj.content)
    result_file = StoredFile.objects.create(
        kind=StoredFile.KIND_RESULT,
        filename=f"results_{timezone.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        content=result_bytes,
        size_bytes=len(result_bytes),
    )
    run = ExecutionRun.objects.create(
        swagger=swagger_obj,
        testcase=testcase_obj,
        result=result_file,
        total=total,
        passed=passed,
        failed=failed,
        notes="",
    )
    return Response(
        {
            "run_id": run.id,
            "result_file_id": result_file.id,
            "summary": {"total": total, "passed": passed, "failed": failed},
        },
        status=status.HTTP_200_OK,
    )


# PUBLIC_INTERFACE
@api_view(['POST'])
@parser_classes([MultiPartParser, FormParser])
def execute_combined(request):
    """
    PUBLIC_INTERFACE
    summary: Execute tests with uploaded swagger and testcases in one request
    description: |
      Accepts multipart/form-data containing:
        - swagger: swagger.json or swagger.txt (OpenAPI/Swagger JSON)
        - testcases: Excel .xlsx file
      The service parses the swagger, executes requests for each row in the Excel,
      adds a 'Status' column (Pass/Fail), and directly returns the updated Excel file.

      Response is a binary .xlsx file and is NOT stored as a result in the database.
    requestBody:
      multipart/form-data:
        fields:
          swagger: file (.json or .txt)
          testcases: file (.xlsx)
    responses:
      200: Returns the updated Excel file
      400: Validation or parsing error
    """
    serializer = CombinedExecuteUploadSerializer(data=request.data)
    if not serializer.is_valid():
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    swagger_file = serializer.validated_data["swagger"]
    testcases_file = serializer.validated_data["testcases"]

    # Read bytes
    swagger_bytes = swagger_file.read()
    testcases_bytes = testcases_file.read()

    # If .txt, still treat as JSON text content
    try:
        # Trigger JSON parsing validation via execute_testcases -> parse_swagger_json
        result_bytes, total, passed, failed = execute_testcases(swagger_bytes, testcases_bytes)
    except Exception as exc:
        return Response({"detail": f"Failed to execute: {exc}"}, status=status.HTTP_400_BAD_REQUEST)

    # Return the updated Excel directly
    filename = f"results_{timezone.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    resp = HttpResponse(
        result_bytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    resp["X-Execution-Summary"] = f"total={total};passed={passed};failed={failed}"
    return resp


# PUBLIC_INTERFACE
@api_view(['GET'])
def download_result(request, file_id: int):
    """
    PUBLIC_INTERFACE
    summary: Download result Excel
    description: Download a previously generated result Excel by its file id.
    responses:
      200: Returns the .xlsx file
      404: Not found
    """
    obj = StoredFile.objects.filter(id=file_id, kind=StoredFile.KIND_RESULT).first()
    if not obj:
        raise Http404("Result not found")
    return FileResponse(
        io.BytesIO(bytes(obj.content)),
        as_attachment=True,
        filename=obj.filename or "results.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
