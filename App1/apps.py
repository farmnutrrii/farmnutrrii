from django.apps import AppConfig


class App1Config(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'App1'

    def ready(self):
        from App1 import scheduler
        scheduler.start_scheduler()

