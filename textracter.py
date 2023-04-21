import os
import textract
import aspose.words as aw

# Исходная директория
input_directory = "D:/Projects/tsiars_gpt/docx_files"

# Конечная директория
output_directory = "D:/Projects/tsiars_gpt/database_txt"

# список обрабатываемых textract'ом типов документов
file_types = (".docx", ".pdf", ".xlsx", ".ppt")

#пересохраняем .doc в .docx (Только для Win версии)
for filename in os.listdir(input_directory):
    if filename.endswith(".doc"):
        input_path = os.path.join(input_directory, filename)
        output_path = os.path.join(input_directory, filename.split(".")[0] + ".docx")
        doc = aw.Document(input_path)
        doc.save(output_path)
        os.remove(input_path)

# Проходимся по директории с условием окончания документов на file_types (P.S. Antiword работает только на Linux)
for filename in os.listdir(input_directory):
    if filename.endswith(file_types):
        # достаем данные из файлов
        text = textract.process(os.path.join(input_directory, filename)).decode("utf-8")
        # создаем новое имя с расширением txt для файла
        new_filename = os.path.splitext(filename)[0] + ".txt"
        # сохраняем файл
        with open(os.path.join(output_directory, new_filename), "w", encoding="utf-8") as f:
            f.write(text)

for filename in os.listdir(output_directory):
    if filename.endswith('.txt'):
        filepath = os.path.join(output_directory, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        lines = [line for line in lines if line.strip()] 
        with open(filepath, 'w', encoding='utf-8') as f:
            f.writelines(lines)