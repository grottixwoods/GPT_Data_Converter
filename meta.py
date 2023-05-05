import csv
import PyPDF2
from docx import Document

import os

input_files = 'input_files'
# Конечная директория
output_txt = 'output_txt'

fieldnames = ["Metadata", "Path to file"]
csv_file = open(os.path.join(output_txt, "metadata.csv"), 'w', newline='')
writer = csv.writer(csv_file)
for filename in os.listdir(input_files):
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
        writer.writerow([metadata_dict, os.path.join(output_txt,
                                                     f'{filename[:-5]}.txt')])
    if filename.endswith(".pdf"):
        # Extract metadata from a PDF document
        pdf_file = open(os.path.join(input_files, filename), 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        metadata_dict = pdf_reader.metadata
        # Write the metadata to the CSV file
        writer.writerow([metadata_dict, os.path.join(output_txt,
                                                     f'{filename[:-5]}.txt')])

csv_file.close()
