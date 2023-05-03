import PyPDF2
import pytesseract
from pdf2image import convert_from_path
import glob


def has_text(pdf_file_path):
    with open(pdf_file_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        for page in pdf_reader.pages:
            text = page.extract_text()
            if text:
                return True
    return False


pdfs = glob.glob(r"input_files/*.pdf")
for pdf in pdfs:
    if has_text(pdf):
        print(f'PDF-файл {pdf} содержит текст')
    else:
        pages = convert_from_path(pdf, 500)
        text = ""
        for pageNum, imgBlob in enumerate(pages):
            text += pytesseract.image_to_string(imgBlob, lang='rus') + '\n'
        with open(f'{pdf[:-4]}.txt', 'w') as the_file:
            the_file.write(text)
        print(f'PDF-файл {pdf} не содержит текст')
