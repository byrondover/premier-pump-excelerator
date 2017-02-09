#!/usr/bin/env python3

import logging
import os

from flask import Flask, redirect, render_template, request, send_file, url_for

from excelerator import Excelerator

app = Flask(__name__)
app.config['PREFERRED_URL_SCHEME'] = 'https'


def get_filename(_file):
    original_filename = _file.filename
    filename, extension = os.path.splitext(original_filename)

    return filename


def get_form(request):
    form = dict()

    form['multiplier'] = int(request.form.get('multiplier', 1))
    form['order_number'] = request.form.get('order-number', str()).strip()
    form['primary_color'] = request.form.get('primary-color', str()).strip()
    form['secondary_color'] = request.form.get('secondary-color', str()).strip()
    form['file'] = request.files.get('file')

    return form


@app.route('/')
def index():
    return render_template('index.html')


# Fails Chrome browser HTTPS security verification. ):
#@app.route('/favicon.ico')
#def favicon():
#    return redirect(url_for('static', filename='img/favicon.ico'))


@app.route('/file-upload', methods=['POST'])
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

    workbook = excelerator.get_workbook_stream()

    # Limit base filename to 64 characters
    excelerated_filename = filename[:64] + ' PPP-E' + '.xlsx'

    return send_file(
        workbook,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=excelerated_filename,
        as_attachment=True
    )


@app.route('/error')
@app.errorhandler(500)
def server_error(error='Unknown'):
    logging.exception('An error occurred during a request.')
    return """
    An internal error occurred: <pre>{}</pre>
    See logs for full stacktrace.
    """.format(error), 500


if __name__ == '__main__':
    # For local debugging only
    # app.run('0.0.0.0', debug=True)
    app.run()
