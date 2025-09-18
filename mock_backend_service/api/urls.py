from django.urls import path
from .views import (
    health,
    upload_swagger,
    upload_testcases,
    list_files,
    execute,
    execute_combined,
    download_result,
)

urlpatterns = [
    path('health/', health, name='Health'),
    path('upload/swagger/', upload_swagger, name='UploadSwagger'),
    path('upload/testcases/', upload_testcases, name='UploadTestcases'),
    path('files/', list_files, name='ListFiles'),
    path('execute/', execute, name='Execute'),
    path('execute/combined/', execute_combined, name='ExecuteCombined'),
    path('results/<int:file_id>/', download_result, name='DownloadResult'),
]
