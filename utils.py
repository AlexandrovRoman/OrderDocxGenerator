import sys
import subprocess
from os.path import split


def month2str(month_index: int,
              months=('январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                      'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь')):
    if month_index not in range(1, 13):
        raise ValueError("Incorrect month index")
    return months[month_index - 1]


def remove_row(table, index):
    row = table.rows[index]
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)


def convert_in_linux(doc_name, pdf_name):
    args = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', split(pdf_name)[0], doc_name]
    subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def convert_in_windows(doc_name, pdf_name):
    from docx2pdf import convert
    convert(doc_name, pdf_name)


def doc2pdf(doc_name, pdf_name):
    if sys.platform.lower().startswith('linux'):
        convert_in_linux(doc_name, pdf_name)
    elif sys.platform.lower().startswith('win'):
        convert_in_windows(doc_name, pdf_name)
