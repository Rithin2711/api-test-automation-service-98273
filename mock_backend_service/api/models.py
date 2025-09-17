from django.db import models


class StoredFile(models.Model):
    """
    Model to store uploaded files and reference them for processing.
    """
    KIND_SWAGGER = "swagger"
    KIND_TESTCASE = "testcase"
    KIND_RESULT = "result"
    KIND_CHOICES = (
        (KIND_SWAGGER, "Swagger JSON"),
        (KIND_TESTCASE, "Test Case Excel"),
        (KIND_RESULT, "Result Excel"),
    )

    kind = models.CharField(max_length=20, choices=KIND_CHOICES)
    filename = models.CharField(max_length=255)
    content = models.BinaryField()  # keep small files in DB for simplicity
    uploaded_at = models.DateTimeField(auto_now_add=True)
    size_bytes = models.IntegerField(default=0)

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self) -> str:  # pragma: no cover
        return f"{self.kind}:{self.filename}"


class ExecutionRun(models.Model):
    """
    Model to represent an execution of test cases against a swagger spec.
    """
    swagger = models.ForeignKey(StoredFile, on_delete=models.CASCADE, related_name="runs_as_swagger")
    testcase = models.ForeignKey(StoredFile, on_delete=models.CASCADE, related_name="runs_as_testcase")
    result = models.ForeignKey(StoredFile, on_delete=models.SET_NULL, null=True, blank=True, related_name="runs_as_result")
    created_at = models.DateTimeField(auto_now_add=True)
    total = models.IntegerField(default=0)
    passed = models.IntegerField(default=0)
    failed = models.IntegerField(default=0)
    notes = models.TextField(blank=True, default="")

    class Meta:
        ordering = ["-created_at"]

    def __str__(self) -> str:  # pragma: no cover
        return f"Run {self.id} - {self.created_at}"
