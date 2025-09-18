from django.urls import path
from .views import (
    health,
    execute_combined,
)

urlpatterns = [
    path('health/', health, name='Health'),
    path('execute/combined/', execute_combined, name='ExecuteCombined'),
]
