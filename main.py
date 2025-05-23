"""Основной модуль для конвертации и обработки документов."""
from typing import NoReturn
import platform
from textracter import (
    convert_doc_to_docx,
    extract_metadata,
    extract_text_from_documents,
    clean_text_files,
    convert_xls_to_xlsx,
    move_files_to_root
)

# Пути к директориям
INPUT_FILES: str = 'input_files'
OUTPUT_TXT: str = 'output_txt'

def main() -> NoReturn:
    """Выполняет основной процесс обработки документов."""
    # Перемещение содержимого из подпапок в основную директорию
    move_files_to_root(INPUT_FILES)
    
    # Конвертация XLS файлов в формат XLSX
    convert_xls_to_xlsx(INPUT_FILES)
    
    # Конвертация DOC файлов в формат DOCX (только для Windows)
    if platform.system() == "Windows":
        convert_doc_to_docx(INPUT_FILES)
    
    # Обработка документов и извлечение текста
    extract_text_from_documents(INPUT_FILES, OUTPUT_TXT)
    
    # Очистка и форматирование извлеченного текста
    clean_text_files(OUTPUT_TXT)

if __name__ == '__main__':
    main()