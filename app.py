#!/usr/bin/env python3

import logging
import os

from flask import Flask, redirect, render_template, request, send_file, url_for

from excelerator import Excelerator

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/favicon.ico')
def favicon():
    return redirect(url_for('static', filename='img/favicon.ico'))


@app.route('/file-upload', methods=['POST'])
def get_tasks():
    multiplier = int(request.form.get('multiplier', 1))

    original_file = request.files.get('file')
    original_filename = original_file.filename
    filename, extension = os.path.splitext(original_filename)

    excelerator = Excelerator(original_file, multiplier)
    workbook = excelerator.get_workbook_stream()

    filename_components = [
        filename,
        '-Excelerated',
        '.xlsx'
    ]
    excelerated_filename = ''.join(filename_components)

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
    app.run(debug=True)
