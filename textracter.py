import os
import textract

# Исходная директория
input_directory = "./docx_files/"

# Конечная директория
output_directory = "./database_txt/"

# список обрабатываемых textract'ом типов документов
file_types = (".doc", ".docx", ".rtf", ".pdf", ".xlsx", ".ppt")

# Проходимся по директории с условием окончания документов на file_types
for filename in os.listdir(input_directory):
    if filename.endswith(file_types):
        # достаем данные из файлов
        text = textract.process(os.path.join(input_directory, filename)).decode("utf-8")
        # создаем новое имя с расширением txt для файла
        new_filename = os.path.splitext(filename)[0] + ".txt"
        # сохраняем файл
        with open(os.path.join(output_directory, new_filename), "w", encoding="utf-8") as f:
            f.write(text)