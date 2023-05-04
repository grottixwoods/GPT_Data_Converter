import os
import textract
import csv
import PyPDF2
import pytesseract
import win32com.client as win32
import re

from pdf2image import convert_from_path
from pdf2img2txt import has_text
from win32com.client import constants
from docx import Document

# Список обрабатываемых textract'ом типов документов
file_types = ('.docx', '.pdf', '.xlsx', '.ppt', '.xls')

def convert_doc_to_docx(input_files):
       
    for filename in os.listdir(input_files):
        # Opening MS Word
        input_path = os.path.join('D:\\Projects\\tsiars_gpt\\input_files', filename)
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(input_path)
        print(doc)
        doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(input_path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
        os.remove(input_path)


def metadata_extracter(input_files, output_txt):
    
    csv_file = open(os.path.join(output_txt, "metadata.csv"), 'w', newline='')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(["Metadata", "Path to file"])

    for filename in os.listdir(input_files):
        if filename.endswith('.pdf'):
            if has_text(filename):
                print(f'PDF-файл {filename} содержит текст')
            else:
                print(f'PDF-файл {filename} не содержит текст')
                pages = convert_from_path(f'input_files/{filename}', 500)
                text = ""
                print(f'PDF-файл {filename} не содержит текст2')
                for pageNum, imgBlob in enumerate(pages):
                    text += pytesseract.image_to_string(imgBlob, lang='rus') + '\n'
                    print(f'PDF-файл {filename} не содержит текст3')
                with open(f'{filename[:-4]}.txt', 'w') as the_file:
                    the_file.write(text)

        if filename.endswith(file_types):
            if filename.endswith(".docx"):
                # Word Meta
                doc = Document(os.path.join(input_files, filename))
                metadata = doc.core_properties
                metadata_dict = {"Title": metadata.title,
                                "Author": metadata.author,
                                "Subject": metadata.subject,
                                "Keywords": metadata.keywords,
                                "Category": metadata.category,
                                "Comments": metadata.comments}
                csv_writer.writerow([metadata_dict, os.path.join(output_txt, f'{filename[:-5]}.txt')])

            elif filename.endswith(".pdf"):
                # Pdf Meta
                pdf_file = open(os.path.join(input_files, filename), 'rb')
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                metadata_dict = pdf_reader.metadata
                csv_writer.writerow([metadata_dict, os.path.join(output_txt, f'{filename[:-5]}.txt')])
    csv_file.close()

def textract_converter(input_files, output_txt):
    
    csv_file = open(os.path.join(output_txt, "metadata.csv"), 'w', newline='')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(["Metadata", "Path to file"])
    
    for filename in os.listdir(input_files):

        # if filename.endswith('.pdf'):
        #     if has_text(filename):
        #         print(f'PDF-файл {filename} содержит текст')
        #     else:
        #         print(f'PDF-файл {filename} не содержит текст')
        #         pages = convert_from_path(f'input_files/{filename}', 500)
        #         text = ""
        #         print(f'PDF-файл {filename} не содержит текст2')
        #         for pageNum, imgBlob in enumerate(pages):
        #             text += pytesseract.image_to_string(imgBlob, lang='rus') + '\n'
        #             print(f'PDF-файл {filename} не содержит текст3')
        #         with open(f'{filename[:-4]}.txt', 'w') as the_file:
        #             the_file.write(text)

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