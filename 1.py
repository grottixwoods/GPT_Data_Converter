import pytesseract
from pdf2image import convert_from_path
import glob

pdfs = glob.glob("scan.pdf")
i=0

pages = convert_from_path("scan.pdf", 300)
print(i)
text = ""
for pageNum,imgBlob in enumerate(pages):
    i+=1
    print(i)
    text+=pytesseract.image_to_string(imgBlob,lang='rus')+'\n'
with open(f'{"scan.pdf"[:-4]}.txt', 'w') as the_file:
    the_file.write(text)