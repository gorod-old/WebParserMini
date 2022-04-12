import os

from aiogram import Bot, Dispatcher
from aiogram.utils import executor
from django.apps import AppConfig


class MainConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'main'
