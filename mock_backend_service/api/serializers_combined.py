from rest_framework import serializers


# PUBLIC_INTERFACE
class CombinedExecuteUploadSerializer(serializers.Serializer):
    """
    PUBLIC_INTERFACE
    Serializer to validate a combined multipart upload containing swagger (JSON or .txt)
    and testcases (Excel .xlsx).
    """
    swagger = serializers.FileField(help_text="Swagger/OpenAPI as .json or .txt")
    testcases = serializers.FileField(help_text="Test cases Excel (.xlsx)")

    def validate_swagger(self, value):
        name = (value.name or "").lower()
        if not (name.endswith(".json") or name.endswith(".txt")):
            raise serializers.ValidationError("Swagger must be .json or .txt")
        return value

    def validate_testcases(self, value):
        name = (value.name or "").lower()
        if not name.endswith(".xlsx"):
            raise serializers.ValidationError("Testcases must be .xlsx")
        return value
