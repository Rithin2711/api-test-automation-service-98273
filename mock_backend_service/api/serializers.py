from rest_framework import serializers


# PUBLIC_INTERFACE
class SwaggerUploadSerializer(serializers.Serializer):
    """
    PUBLIC_INTERFACE
    Serializer to validate swagger.json uploads.
    """
    file = serializers.FileField(help_text="Swagger/OpenAPI JSON file (swagger.json)")

    def validate_file(self, value):
        if not value.name.lower().endswith(".json"):
            raise serializers.ValidationError("Only .json files are accepted for swagger.")
        return value


# PUBLIC_INTERFACE
class TestCaseUploadSerializer(serializers.Serializer):
    """
    PUBLIC_INTERFACE
    Serializer to validate Excel test cases uploads.
    """
    file = serializers.FileField(help_text="Excel file containing test cases (.xlsx)")

    def validate_file(self, value):
        allowed = (".xlsx",)
        if not any(value.name.lower().endswith(ext) for ext in allowed):
            raise serializers.ValidationError("Only .xlsx files are accepted for test cases.")
        return value


# PUBLIC_INTERFACE
class ExecuteRequestSerializer(serializers.Serializer):
    """
    PUBLIC_INTERFACE
    Serializer to specify which stored swagger and testcase to execute.
    If omitted, uses the latest uploaded files.
    """
    swagger_id = serializers.IntegerField(required=False, help_text="ID of stored swagger to use.")
    testcase_id = serializers.IntegerField(required=False, help_text="ID of stored testcase to use.")


# PUBLIC_INTERFACE
class StoredFileInfoSerializer(serializers.Serializer):
    """
    PUBLIC_INTERFACE
    Serializer to return stored file information.
    """
    id = serializers.IntegerField()
    kind = serializers.CharField()
    filename = serializers.CharField()
    uploaded_at = serializers.DateTimeField()
    size_bytes = serializers.IntegerField()
