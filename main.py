from textracter import convert_doc_to_docx, metadata_extracter, textract_converter, lines_editor

# Исходная директория
input_files = 'input_files'
# Конечная директория
output_txt = 'output_txt'

if __name__ == '__main__':
    convert_doc_to_docx(input_files)
    metadata_extracter(input_files, output_txt)
    textract_converter(input_files, output_txt)
    lines_editor(output_txt)