import os
import textract
import platform
from web import input_files, output_txt

# Список обрабатываемых textract'ом типов документов
file_types = ('.docx', '.pdf', '.xlsx', '.ppt', '.xls')

def converter():
    # Пересохраняем .doc в .docx (Только для Win версии)
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

    # Проходимся по директории с условием окончания документов на file_types
    # (P.S. Antiword работает только на Linux)
    for filename in os.listdir(input_files):
        if filename.endswith(file_types):
            # Достаем данные из файлов
            input_path = os.path.join(input_files, filename)
            text = textract.process(os.path.join(input_files,
                                                 filename)).decode('utf-8')
            # Создаем новое имя с расширением txt для файла
            new_filename = os.path.splitext(filename)[0] + '.txt'
            # Сохраняем файл
            with open(os.path.join(output_txt, new_filename), 'w',
                      encoding='utf-8') as f:
                f.write(text)
            os.remove(input_path)

    # Проходимся по директории с условием окончания документов на .txt
    for filename in os.listdir(output_txt):
        if filename.endswith('.txt'):
            filepath = os.path.join(output_txt, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            # Убираем пустые cтроки
            lines = [line for line in lines if line.strip()]
            with open(filepath, 'w', encoding='utf-8') as f:
                f.writelines(lines)
