import os
import textract
import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
import pandas as pd
import win32com.client as win32
import re
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from win32com.client import constants
from docx import Document
import pytesseract
from pdf2image import convert_from_path
import PyPDF2


# __________________TODO__________________
#     Доделать обработку сканов pdf

# Список обрабатываемых textract'ом типов документов
file_types = ('.docx', '.pdf', '.xlsx', '.ppt', '.xls')


def has_text(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                return True
    return False


def convert_xls_to_xlsx(input_files):
    for filename in os.listdir(input_files):
        if filename.endswith('.xls'):
            xls_path = os.path.join(input_files, filename)
            wb = xlrd.open_workbook(xls_path)

            xlsx_path = os.path.join(input_files, filename[:-4] + '.xlsx')
            new_wb = Workbook()

            ws = new_wb.active

            for sheet_name in wb.sheet_names():
                sheet = wb.sheet_by_name(sheet_name)

                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        col_letter = get_column_letter(col + 1)
                        cell_value = sheet.cell_value(row, col)
                        ws[f"{col_letter}{row + 1}"] = cell_value

            new_wb.save(xlsx_path)
            os.remove(xls_path)


def convert_doc_to_docx(input_files):
    for filename in os.listdir(input_files):
        if filename.endswith(".doc"):
            input_path = os.path.join('D:\\Projects\\tsiars_gpt\\input_files', filename)
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(input_path)
            print(doc)
            doc.Activate()

            new_file_abs = os.path.abspath(input_path)
            new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
            word.ActiveDocument.SaveAs(
                new_file_abs, FileFormat=constants.wdFormatXMLDocument
            )
            doc.Close(False)
            os.remove(input_path)


def metadata_extracter(input_files, output_txt):
    df = pd.DataFrame(columns=['Meta', 'Path'])

    for filename in os.listdir(input_files):
        if filename.endswith(".docx"):
            doc = Document(os.path.join(input_files, filename))
            metadata = doc.core_properties
            metadata_dict = {"Title": metadata.title,
                             "Author": metadata.author,
                             "Subject": metadata.subject,
                             "Keywords": metadata.keywords,
                             "Category": metadata.category,
                             "Comments": metadata.comments}

            df = df._append({'Meta': metadata_dict,
                            'Path': os.path.join(output_txt, f'{filename[:-5]}.txt')},
                            ignore_index=True)

        elif filename.endswith(".pdf"):
            with open(os.path.join(input_files, filename), 'rb') as pdf_file:
                parser = PDFParser(pdf_file)
                doc = PDFDocument(parser)
                metadata_dict = doc.info[0]

                df = df._append({'Meta': metadata_dict,
                                'Path': os.path.join(output_txt, f'{filename[:-5]}.txt')},
                                ignore_index=True)

        elif filename.endswith(".xlsx") or filename.endswith(".xlsm") or filename.endswith(".xltx") or filename.endswith(".xltm"):
            wb = openpyxl.load_workbook(os.path.join(input_files, filename))
            metadata_dict = {}
            metadata_dict['authors'] = wb.properties.creator
            metadata_dict['theme'] = wb.properties.title
            metadata_dict['tags'] = wb.properties.keywords
            metadata_dict['category'] = wb.properties.category
            metadata_dict['comments'] = wb.properties.description
            df = df._append({'Meta': metadata_dict,
                            'Path': os.path.join(output_txt, f'{filename[:-5]}.txt')},
                            ignore_index=True)

    df.to_csv('dataframe.csv', index=False)


def textract_converter(input_files, output_txt):
    for filename in os.listdir(input_files):
        if filename.endswith('.pdf'):
            if has_text(filename):
                pass
                # Перевод PDF -> Docx -> txt
            else:
                pages = convert_from_path(os.path.join(input_files, filename), 500)
                text = ""
                for pageNum, imgBlob in enumerate(pages):
                    text += pytesseract.image_to_string(imgBlob, lang='rus') + '\n'
                with open(f'{filename[:-4]}.txt', 'w') as the_file:
                    the_file.write(text)

        if filename.endswith(file_types):
            text = textract.process(os.path.join(input_files, filename)).decode('utf-8')
            new_filename = os.path.splitext(filename)[0] + '.txt'
            with open(os.path.join(output_txt, new_filename), 'w', encoding='utf-8') as f:
                f.write(text)


def lines_editor(output_txt):
    for filename in os.listdir(output_txt):
        if filename.endswith('.txt'):
            filepath = os.path.join(output_txt, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            lines = [line for line in lines if line.strip()]
            with open(filepath, 'w', encoding='utf-8') as f:
                f.writelines(lines)