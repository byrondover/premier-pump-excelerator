#!/usr/bin/env python3

import os
import sys

from django.conf import settings
from django.conf.urls import url
from django.core.management import execute_from_command_line
from django.http import HttpResponse
from django.shortcuts import render_to_response
from django.views.decorators.csrf import csrf_exempt

from excelerator import Excelerator

SITE_ROOT = os.path.dirname(os.path.realpath(__file__))

settings.configure(
    DEBUG=True,
    SECRET_KEY='A-random-secret-key!',
    ROOT_URLCONF=sys.modules[__name__],
    STATIC_URL = '/static/',
    STATICFILES_DIRS = [
        '.',
        os.path.join(SITE_ROOT, 'static/'),
    ],
    TEMPLATES = [
        {
            'BACKEND': 'django.template.backends.django.DjangoTemplates',
            'DIRS': ['.']
        }
    ],
    INSTALLED_APPS = ['django.contrib.staticfiles']
)


def index(request):
    return render_to_response('templates/index.html')

@csrf_exempt
def file_upload(request):
    original_file = request.FILES['file'].file
    original_filename = request.FILES['file'].name

    excelerator = Excelerator()
    wb = excelerator.excelerate(original_file)

    filename_components = [
        '.'.join(original_filename.split('.')[:-1]),
        '_RENDERED.',
        original_filename.split('.')[-1]
    ]
    excelerated_filename = ' '.join(filename_components)

    response = HttpResponse(
        content_type='/'.join([
            'application',
            'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ])
    )
    wb.save(response)

    return response


urlpatterns = [
    url(r'^$', index),
    url(r'^file-upload$', file_upload),
]

if __name__ == '__main__':
    execute_from_command_line(sys.argv)
