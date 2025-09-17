from django.apps import AppConfig


class ApiConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'api'

    def ready(self):
        # Placeholder for signals if needed in future.
        # Ensures import path exists without heavy init.
        return
