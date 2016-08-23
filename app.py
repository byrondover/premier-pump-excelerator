#!/usr/bin/env python3

import io

from flask import Flask, make_response, render_template, request, send_file, send_from_directory

from excelerator import Excelerator

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/file-upload', methods=['POST'])
def get_tasks():
    original_file = request.files['file']
    original_filename = original_file.filename

    excelerator = Excelerator()
    wb = excelerator.excelerate(original_file)
    workbook = excelerator.get_workbook()


    from io import BytesIO
    #buffer = BytesIO()
    #buffer.write(workbook.decode('latin-1').encode())
    #buffer.seek(0)
    #wb.save(buffer)

    import pdb; pdb.set_trace()
    wb.save('tmp.xlsx')

    filename_components = [
        '.'.join(original_filename.split('.')[:-1]),
        '_RENDERED.',
        original_filename.split('.')[-1]
    ]
    excelerated_filename = ' '.join(filename_components)

    #response = make_response(workbook)
    #content_disposition = "attachment; filename={filename}".format(
    #    filename=excelerated_filename)
    #response.headers['Content-Disposition'] = content_disposition

    return send_from_directory('.',
        'tmp.xlsx'
        #mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        #attachment_filename=excelerated_filename,
        #as_attachment=False
    )

if __name__ == '__main__':
    app.run(debug=True)
