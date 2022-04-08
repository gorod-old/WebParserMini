from django.http import HttpResponse
from django.shortcuts import render

# Create your views here.
from main.parser import pars_data


def index(request):
    # pars_data()
    return render(request, 'main/index.html')
