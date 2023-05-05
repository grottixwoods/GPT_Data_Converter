import pytesseract
from pdf2image import convert_from_path


import os

pdfs = ['tsiars_gpt/input_files/Аргументы и факты 202102.pdf', 'tsiars_gpt/input_files/scan.pdf']
i = 0
for pdf in pdfs:
    pages = convert_from_path(pdf, 100)
    print(i, 'картинки')
    i += 1
    text = ""
    for pageNum, imgBlob in enumerate(pages):
        print(i, 'Текстики')
        i += 1
        text += pytesseract.image_to_string(imgBlob, lang='rus')+'\n'
    with open(f'{pdf[:-4]}.txt', 'w') as the_file:
        the_file.write(text)
        os.remove(pdf)