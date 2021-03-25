import os

import xlrd
import openpyxl
import pandas as pd

from flask import render_template, request, redirect, send_from_directory, abort
from googletrans import Translator
from xlwt import Workbook

from app import app

# @app.route('/')
# def index():
#     return render_template('public/index.html')

#
# @app.route('/about')
# def about():
#     return "<h1 style='color:red'>About Us</h1>"


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


@app.route('/', methods=['GET', 'POST'])
def translator():
    names = [
        'Abkhazian',
        'Afar',
        'Afrikaans',
        'Akan',
        'Albanian',
        'Amharic',
        'Arabic',
        'Aragonese',
        'Armenian',
        'Assamese',
        'Avaric',
        'Avestan',
        'Aymara',
        'Azerbaijani',
        'Bambara',
        'Bashkir',
        'Basque',
        'Belarusian',
        'Bengali',
        'Bihari languages',
        'Bislama',
        'Bosnian',
        'Breton',
        'Bulgarian',
        'Burmese',
        'Catalan, Valencian',
        'Central Khmer',
        'Chamorro',
        'Chechen',
        'Chichewa, Chewa, Nyanja',
        'Chinese',
        'Church Slavonic, Old Bulgarian, Old Church Slavonic',
        'Chuvash',
        'Cornish',
        'Corsican',
        'Cree',
        'Croatian',
        'Czech',
        'Danish',
        'Divehi, Dhivehi, Maldivian',
        'Dutch, Flemish',
        'Dzongkha',
        'English',
        'Esperanto',
        'Estonian',
        'Ewe',
        'Faroese',
        'Fijian',
        'Finnish',
        'French',
        'Fulah',
        'Gaelic, Scottish Gaelic',
        'Galician',
        'Ganda',
        'Georgian',
        'German',
        'Gikuyu, Kikuyu',
        'Greek (Modern)',
        'Greenlandic, Kalaallisut',
        'Guarani',
        'Gujarati',
        'Haitian, Haitian Creole',
        'Hausa',
        'Hebrew',
        'Herero',
        'Hindi',
        'Hiri Motu',
        'Hungarian',
        'Icelandic',
        'Ido',
        'Igbo',
        'Indonesian',
        'Interlingua (International Auxiliary Language Association)',
        'Interlingue',
        'Inuktitut',
        'Inupiaq',
        'Irish',
        'Italian',
        'Japanese',
        'Javanese',
        'Kannada',
        'Kanuri',
        'Kashmiri',
        'Kazakh',
        'Kinyarwanda',
        'Komi',
        'Kongo',
        'Korean',
        'Kwanyama, Kuanyama',
        'Kurdish',
        'Kyrgyz',
        'Lao',
        'Latin',
        'Latvian',
        'Letzeburgesch, Luxembourgish',
        'Limburgish, Limburgan, Limburger',
        'Lingala',
        'Lithuanian',
        'Luba-Katanga',
        'Macedonian',
        'Malagasy',
        'Malay',
        'Malayalam',
        'Maltese',
        'Manx',
        'Maori',
        'Marathi',
        'Marshallese',
        'Moldovan, Moldavian, Romanian',
        'Mongolian',
        'Nauru',
        'Navajo, Navaho',
        'Northern Ndebele',
        'Ndonga',
        'Nepali',
        'Northern Sami',
        'Norwegian',
        'Norwegian Bokmål',
        'Norwegian Nynorsk',
        'Nuosu, Sichuan Yi',
        'Occitan (post 1500)',
        'Ojibwa',
        'Oriya',
        'Oromo',
        'Ossetian, Ossetic',
        'Pali',
        'Panjabi, Punjabi',
        'Pashto, Pushto',
        'Persian',
        'Polish',
        'Portuguese',
        'Quechua',
        'Romansh',
        'Rundi',
        'Russian',
        'Samoan',
        'Sango',
        'Sanskrit',
        'Sardinian',
        'Serbian',
        'Shona',
        'Sindhi',
        'Sinhala, Sinhalese',
        'Slovak',
        'Slovenian',
        'Somali',
        'Sotho, Southern',
        'South Ndebele',
        'Spanish, Castilian',
        'Sundanese',
        'Swahili',
        'Swati',
        'Swedish',
        'Tagalog',
        'Tahitian',
        'Tajik',
        'Tamil',
        'Tatar',
        'Telugu',
        'Thai',
        'Tibetan',
        'Tigrinya',
        'Tonga (Tonga Islands)',
        'Tsonga',
        'Tswana',
        'Turkish',
        'Turkmen',
        'Twi',
        'Uighur, Uyghur',
        'Ukrainian',
        'Urdu',
        'Uzbek',
        'Venda',
        'Vietnamese',
        'Volap_k',
        'Walloon',
        'Welsh',
        'Western Frisian',
        'Wolof',
        'Xhosa',
        'Yiddish',
        'Yoruba',
        'Zhuang, Chuang',
        'Zulu'
    ]
    languages = ['ab', 'aa', 'af', 'ak', 'sq', 'am', 'ar', 'an', 'hy', 'as', 'av', 'ae', 'ay', 'az', 'bm',
                 'ba', 'eu', 'be', 'bn', 'bh', 'bi', 'bs', 'br', 'bg', 'my', 'ca', 'km', 'ch', 'ce', 'ny',
                 'zh', 'cu', 'cv', 'kw', 'co', 'cr', 'hr', 'cs', 'da', 'dv', 'nl', 'dz', 'en', 'eo', 'et',
                 'ee', 'fo', 'fj', 'fi', 'fr', 'ff', 'gd', 'gl', 'lg', 'ka', 'de', 'ki', 'el', 'kl', 'gn',
                 'gu', 'ht', 'ha', 'he', 'hz', 'hi', 'ho', 'hu', 'is', 'io', 'ig', 'id', 'ia', 'ie', 'iu',
                 'ik', 'ga', 'it', 'ja', 'jv', 'kn', 'kr', 'ks', 'kk', 'rw', 'kv', 'kg', 'ko', 'kj', 'ku',
                 'ky', 'lo', 'la', 'lv', 'lb', 'li', 'ln', 'lt', 'lu', 'mk', 'mg', 'ms', 'ml', 'mt', 'gv',
                 'mi', 'mr', 'mh', 'ro', 'mn', 'na', 'nv', 'nd', 'ng', 'ne', 'se', 'no', 'nb', 'nn', 'ii',
                 'oc', 'oj', 'or', 'om', 'os', 'pi', 'pa', 'ps', 'fa', 'pl', 'pt', 'qu', 'rm', 'rn', 'ru',
                 'sm', 'sg', 'sa', 'sc', 'sr', 'sn', 'sd', 'si', 'sk', 'sl', 'so', 'st', 'nr', 'es', 'su',
                 'sw', 'ss', 'sv', 'tl', 'ty', 'tg', 'ta', 'tt', 'te', 'th', 'bo', 'ti', 'to', 'ts', 'tn',
                 'tr', 'tk', 'tw', 'ug', 'uk', 'ur', 'uz', 've', 'vi', 'vo', 'wa', 'cy', 'fy', 'wo', 'xh', 'yi', 'yo',
                 'za', 'zu']

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
                            value = translator.translate(value, dest="sw").text
                        sheet1.write(column, row, value)
                wb_w.save(r'' + os.path.join(app.config['CLIENT_EXCELS'], excel.filename))
                print("Translated")
                print("Done")

            return redirect(request.url)
    return render_template('public/translator.html', languages=languages)


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
