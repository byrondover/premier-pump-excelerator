#!/usr/bin/env python3

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
    original_file = request.files['file']
    original_filename = original_file.filename

    excelerator = Excelerator(original_file)
    workbook = excelerator.get_workbook_stream()

    filename_components = [
        '.'.join(original_filename.split('.')[:-1]),
        '-Excelerated.',
        original_filename.split('.')[-1]
    ]
    excelerated_filename = ''.join(filename_components)

    return send_file(
        workbook,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=excelerated_filename,
        as_attachment=True
    )


if __name__ == '__main__':
    app.run(debug=True)
