#вывод с текстом картинок и без начала
import os
from pdfminer.high_level import extract_text

# Путь к директории входных файлов
input_dir = "C:/Users/marin/PycharmProjects/tsiars_gpt/input_files"

# Путь к директории выходных файлов
output_dir = "C:/Users/marin/PycharmProjects/tsiars_gpt/output_txt"

# Проверяем, существует ли директория, и если нет, то создаем ее
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Открываем PDF-файл
with open(os.path.join(input_dir, "Аргументы и факты 2021'01.pdf"), "rb") as pdf_file:
    # Извлекаем текст из PDF-файла с помощью pdfminer.six
    text = extract_text(pdf_file)

    # Разбиваем текст на абзацы и сохраняем в список
    paragraphs = text.split("\n\n")


    # Открываем файл для записи
    with open(os.path.join(output_dir, "document.txt"), "w", encoding="utf-8") as txt_file:
        # Записываем каждый абзац в отдельную строку в текстовом файле
        for paragraph_num in range(len(paragraphs)):
            paragraph = paragraphs[paragraph_num]
            if paragraph_num == 0:
                txt_file.write(paragraph)
            else:
                txt_file.write("\n\n" + paragraph)
