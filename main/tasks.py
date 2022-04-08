import os
import shutil

from celery import group

from main.parser import pars_data
from webparsermini.celery import app


@app.task
def run_pars():
    pars_data()


@app.task(queue='for_parsers', routing_key='for_parsers')
def run_parser():
    pass


