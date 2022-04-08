import os

from celery import Celery
from celery.schedules import crontab

# Set the default Django settings module for the 'celery' program.
from kombu import Queue, Exchange

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'webparsermini.settings')

app = Celery('webparsermini')

# Using a string here means the worker doesn't have to serialize
# the configuration object to child processes.
# - namespace='CELERY' means all celery-related configuration keys
#   should have a `CELERY_` prefix.
app.config_from_object('django.conf:settings', namespace='CELERY')

# Load task modules from all registered Django apps.
app.autodiscover_tasks()

app.conf.beat_schedule = {
    'run_pars': {
        'task': 'main.tasks.run_pars',
        'schedule': crontab(minute='*/15'),
    },
}
app.conf.timezone = 'UTC'


@app.task(bind=True)
def debug_task(self):
    print(f'Request: {self.request!r}')
