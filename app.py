#!/usr/bin/env python3

import logging
import os
import re
import sys
import uuid
from base64 import b64decode

import requests
from flask import Flask, redirect, render_template, request, send_file, url_for
from flask_cors import cross_origin
from flask_sslify import SSLify

from excelerator import Excelerator

app = Flask(__name__)
app.config['PREFERRED_URL_SCHEME'] = 'https'

# Google Cloud environment variables are defined in app.yaml
APP_ENV = os.environ.get('APP_ENV', 'development')
GA_TRACKING_ID = os.environ.get('GA_TRACKING_ID')
MAILGUN_DOMAIN = os.environ.get('MAILGUN_DOMAIN')
MAILGUN_SVALUE = os.environ.get('MAILGUN_SVALUE')

EMAIL_ADMINS='byrondover+ppp-e@gmail.com'
EMAIL_RECIPIENTS='bill@premier-pump.net'
CORS_ALLOWED_DOMAINS = [
    'https://excelerator.premier-pump.io',
    'https://premier-pump-excelerator.appspot.com'
]
global ORDER_NUMBER
ORDER_NUMBER=4000
VERSION = '2.1.1'
YEAR_IN_SECS = 31536000


def calculate_order_number(raw_order_number):
    numbers_and_spaces = re.compile(r'[^\d\s]+')
    stripped_order_number = numbers_and_spaces.sub('', str(raw_order_number))
    numbers = list(filter(None, stripped_order_number.split(' ')))

    if len(numbers) > 0:
        return int(numbers[0]) + 1

    return ORDER_NUMBER


def get_filename(_file):
    original_filename = _file.filename
    filename, extension = os.path.splitext(original_filename)

    return filename


def get_form(request):
    form = dict()

    form['multiplier'] = int(request.form.get('multiplier', 1))
    form['order_number'] = request.form.get('order-number', '').strip()
    form['primary_color'] = request.form.get('primary-color', '').strip()
    form['secondary_color'] = request.form.get('secondary-color', '').strip()
    form['file'] = request.files.get('file')

    return form


def send_email(to, bcc='', filename='', multiplier='None', order_number='None',
               primary_color='None', secondary_color='None'):
    if MAILGUN_DOMAIN and MAILGUN_SVALUE:
        url = 'https://api.mailgun.net/v3/{}/messages'.format(MAILGUN_DOMAIN)
        svalue = (
            'api',
            b64decode(b64decode(
                str(MAILGUN_SVALUE + '==').encode('UTF-8')
            )).decode('UTF-8')
        )
        data = {
            'from': 'PPP Bot <noreply@premier-pump.io>',
            'to': to,
            'bcc': bcc,
            'subject': '[PPP-E] File Uploaded',
            'text': """File Uploaded: {filename}

            Order Number: {order_number}
            Multiplier: {multiplier}
            Primary Color: {primary_color}
            Secondary Color: {secondary_color}
            """.format(**locals())
        }

        response = requests.post(url, auth=svalue, data=data)
        response.raise_for_status()


def track_event(category, action, label=None, value=0, ip_addr=None):
    if GA_TRACKING_ID:
        data = {
            'v': '1',  # API Version.
            'tid': GA_TRACKING_ID,  # Tracking ID / Property ID.
            # Anonymous Client Identifier. Ideally, this should be a UUID that
            # is associated with particular user, device, or browser instance.
            'cid': str(ip_addr) if ip_addr else uuid.uuid4(),
            't': 'event',  # Event hit type.
            'ec': category,  # Event category.
            'ea': action,  # Event action.
            'el': label,  # Event label.
            'ev': value,  # Event value, must be an integer
        }

        response = requests.post(
            'https://www.google-analytics.com/collect', data=data)

        # If the request fails, this will raise a RequestException. Depending
        # on your application's needs, this may be a non-error and can be caught
        # by the caller.
        response.raise_for_status()


class SSLifyImproved(SSLify):

    def __init__(self, app=None, age=YEAR_IN_SECS, subdomains=False,
                 permanent=False, skips=None):
        super().__init__(app, age, subdomains, permanent, skips)

    @property
    def hsts_header(self):
        """Returns the proper HSTS policy."""
        hsts_policy = 'max-age={0}'.format(self.hsts_age)

        if self.hsts_include_subdomains:
            hsts_policy = '; includeSubDomains'

        hsts_policy += '; preload'

        return hsts_policy


if APP_ENV == 'production':
    sslify = SSLifyImproved(app, permanent=True, subdomains=True)


@app.route('/')
def index():
    base_url = str()

    if APP_ENV == 'production':
        base_url = 'https://excelerator.premier-pump.io'

    return render_template('index.html', base_url=base_url, version=VERSION)


@app.route('/favicon.ico')
def favicon():
    if APP_ENV == 'production':
        return redirect("https://excelerator.premier-pump.io/static/img/favicon.ico")
    else:
        # Fails Chrome browser HTTPS security verification. ):
        return redirect(url_for('static', filename='img/favicon.ico'))


@app.route('/file-upload', methods=['OPTIONS', 'POST'])
@cross_origin(CORS_ALLOWED_DOMAINS)
def get_tasks():
    form = get_form(request)
    filename = get_filename(form['file'])

    excelerator = Excelerator(
        form['file'],
        form['multiplier'],
        form['order_number'],
        form['primary_color'],
        form['secondary_color']
    )

    global ORDER_NUMBER
    ORDER_NUMBER = calculate_order_number(form['order_number'])
    workbook = excelerator.get_workbook_stream()

    # Limit base filename to 64 characters
    excelerated_filename = filename[:64] + ' PPP-E' + '.xlsx'

    if APP_ENV == 'production':
        try:
            send_email(EMAIL_RECIPIENTS,
                       bcc=EMAIL_ADMINS,
                       filename=form['file'].filename,
                       multiplier=form['multiplier'] or 'None',
                       order_number=form['order_number'] or 'None',
                       primary_color=form['primary_color'] or 'None',
                       secondary_color=form['secondary_color'] or 'None')
            track_event(category='File', action='uploaded', label=filename,
                        value=form['multiplier'], ip_addr=request.remote_addr)
        except Exception as e:
            # Email call or Google Analyics call fails? No big deal.
            # Log out the error and move on.
            print(e, file=sys.stderr)
            pass

    return send_file(
        workbook,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=excelerated_filename,
        as_attachment=True
    )


@app.route('/order-number')
@cross_origin(CORS_ALLOWED_DOMAINS)
def order_number():
    return str(ORDER_NUMBER)


@app.route('/error')
@app.errorhandler(500)
def server_error(error='Unknown'):
    logging.exception('An error occurred during a request.')
    return """
    An internal error occurred: <pre>{}</pre>
    See logs for full stacktrace.
    """.format(error), 500


if __name__ == '__main__':
    if APP_ENV == 'production':
        app.run()
    else:
        # For local debugging only
        app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)), debug=True)
