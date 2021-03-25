import os

import xlrd
import openpyxl
import pandas as pd

from flask import render_template, request, redirect, send_from_directory, abort
from googletrans import Translator
from xlwt import Workbook

from app import app


@app.route('/')
def index():
    return render_template('public/index.html')


@app.route('/about')
def about():
    return "<h1 style='color:red'>About Us</h1>"


app.config['FILE_UPLOADS'] = "/home/belbet/app/app/static/files/uploads"
app.config['ALLOWED_FILE_EXTENSIONS'] = ["XLS", "XLSX"]
app.config['CLIENT_EXCELS'] = "/home/belbet/app/app/static/client/csv"


def allowed_file(filename):
    if not "." in filename:
        return False

    # ext = filename.rsplit(".", 1)[1]

    # if ext.upper() in app.config['ALLOWED_FILE_EXTENSIONS']:
    #     return True

    # else:
    #     return False


@app.route('/translator', methods=['GET', 'POST'])
def translator():
    if request.method == "POST":
        if request.files:  # filename = secure_filename.filename(excel.filename)

            excel = request.files['excel']

            if excel.filename == "":
                print("Empty file name")
                return redirect(request.url)

            # if not allowed_file(excel.filename):
            #     print("Invalid file. Kindly upload an excel file")
            #     return redirect(request.url)

            else:
                # filename = secure_filename.filename(excel.filename)
                excel.save(os.path.join(app.config['FILE_UPLOADS'], excel.filename))

                translator = Translator(service_urls=['translate.googleapis.com'])

                location = (os.path.join(app.config['FILE_UPLOADS'], excel.filename))

                print("Lost here")

                # Writing to file
                wb_w = Workbook()
                sheet1 = wb_w.add_sheet('sheet 1')
                # Reading file
                # wb_r = openpyxl.load_workbook(location, False)
                # wb_r = pd.read_excel(location, engine='openpyxl')
                wb_r = xlrd.open_workbook(location)
                sheet = wb_r.sheet_by_index(0)
                sheet.cell_value(0, 0)

                for column in range(sheet.nrows):

                    for row in range(sheet.ncols):
                        value = sheet.cell_value(column, row)
                        if type(value) == str:
                            value = translator.translate(value, dest='es').text
                        sheet1.write(column, row, value)
                wb_w.save(r'' + os.path.join(app.config['CLIENT_EXCELS'], excel.filename))
                print("Translated")



            return redirect(request.url)
    return render_template('public/translator.html')


"""
string:
int:
float:
uuid:
"""
@app.route('/get-file/<file_name>')
def get_file(file_name):
    try:
        return send_from_directory(app.config['CLIENT_EXCELS'], filename=file_name,
        as_attachment=True)
    except FileNotFoundError:
        abort(404)
