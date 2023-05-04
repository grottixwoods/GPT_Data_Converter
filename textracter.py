import os
import textract
import platform
import csv
from docx import Document
import PyPDF2
import pytesseract
from pdf2image import convert_from_path


# Исходная директория
input_files = 'input_files'
# Конечная директория
output_txt = 'output_txt'

# Список обрабатываемых textract'ом типов документов
file_types = ('.docx', '.pdf', '.xlsx', '.ppt', '.xls')

def has_text(pdf_file_path):
    with open(f'input_files/{pdf_file_path}', 'rb') as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        for i in range(2):
            page = pdf_reader.getPage(i)
            text = page.extractText()
            if text and len(text) > 100:
                return True
    return False

def converter():
    # Пересохраняем .doc в .docx (Только для Windows)
    if platform.system() == "Windows":
        import aspose.words as aw
        for filename in os.listdir(input_files):
            if filename.endswith('.doc'):
                input_path = os.path.join(input_files, filename)
                output_path = os.path.join(input_files,
                                           filename.split('.')[0] + '.docx')
                doc = aw.Document(input_path)
                doc.save(output_path)
                os.remove(input_path)

    csv_file = open(os.path.join(output_txt, "metadata.csv"), 'w', newline='')
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(["Metadata", "Path to file"])

    # Проходимся по директории с условием окончания документов на file_types
    # (P.S. Antiword работает только на MacOS)
    for filename in os.listdir(input_files):
        if filename.endswith(file_types):
            # Достаем данные из файлов
            input_path = os.path.join(input_files, filename)
            if filename.endswith(".docx"):
                # Extract metadata from a Word document
                doc = Document(os.path.join(input_files, filename))
                metadata = doc.core_properties
                metadata_dict = {"Title": metadata.title,
                                 "Author": metadata.author,
                                 "Subject": metadata.subject,
                                 "Keywords": metadata.keywords,
                                 "Category": metadata.category,
                                 "Comments": metadata.comments}
                # Write the metadata to the CSV file
                csv_writer.writerow
                (
                    [
                        metadata_dict, os.path.join(
                                                 output_txt,
                                                 f'{filename[:-5]}.txt'
                                                )
                    ]
                )
            elif filename.endswith(".pdf"):
                # Extract metadata from a PDF document
                pdf_file = open(os.path.join(input_files, filename), 'rb')
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                metadata_dict = pdf_reader.metadata
                # Write the metadata to the CSV file
                csv_writer.writerow
                (
                    [
                        metadata_dict, os.path.join(
                                                 output_txt,
                                                 f'{filename[:-5]}.txt'
                                                )
                    ]
                )
            text = textract.process(os.path.join(input_files,
                                                 filename)).decode('utf-8')
            # Создаем новое имя с расширением txt для файла
            new_filename = os.path.splitext(filename)[0] + '.txt'
            # Сохраняем файл
            with open(os.path.join(output_txt, new_filename), 'w',
                      encoding='utf-8') as f:
                f.write(text)
    csv_file.close()

    # Проходимся по директории с условием окончания документов на .txt
    for filename in os.listdir(output_txt):
        if filename.endswith('.txt'):
            filepath = os.path.join(output_txt, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            lines = [line for line in lines if line.strip()]
            # Убираем вотермарки
            if 'Created with an evaluation copy of Aspose.Words.' in lines[-1]:
                lines.pop()
                lines.pop(0)
            # Убираем пустые cтроки
            with open(filepath, 'w', encoding='utf-8') as f:
                f.writelines(lines)

    os.remove(input_path)