from django.apps import AppConfig


class ElectionStatisticsConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "election_statistics"

    def ready(self):
        import election_statistics.signals
