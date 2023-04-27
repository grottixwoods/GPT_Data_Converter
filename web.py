from bs4 import BeautifulSoup
import urllib3
import wget
from textracter import converter, input_files, output_txt
import os

# Проверяем и создаем исходную директорию
if not os.path.exists(input_files):
    os.makedirs(input_files)
# Проверяем и создаем конечную директорию
if not os.path.exists(output_txt):
    os.makedirs(output_txt)


url_root = 'https://www.pravovik24.ru'
url = "https://www.pravovik24.ru/documents"

http = urllib3.PoolManager()
links = []

response = http.request('GET', url)

soup = BeautifulSoup(
    response.data.decode('utf-8'),
    'html.parser'
)
for link in soup.findAll('a'):
    buf = link.get('href')
    if '/documents/dogovory/' in buf:
        response1 = http.request(
            'GET',
            url_root+buf
        )
        soup2 = BeautifulSoup(
            response1.data.decode('utf-8'),
            'html.parser'
        )
        if 'documents' in buf:
            response2 = http.request(
                'GET',
                url_root+buf
            )
            soup2 = BeautifulSoup(
                response1.data.decode('utf-8'),
                'html.parser'
            )
            for link in soup2.findAll('a'):
                buf = link.get('href')
                if 'upload' in buf:
                    response = wget.download(
                        url_root+buf,
                        out=input_files
                    )
                    converter()
    elif 'documents' in buf:
        response1 = http.request(
            'GET',
            url_root+buf
        )
        soup2 = BeautifulSoup(
            response1.data.decode('utf-8'),
            'html.parser'
        )
        for link in soup2.findAll('a'):
            buf = link.get('href')
            if 'upload' in buf:
                response = wget.download(
                    url_root+buf,
                    out=input_files
                )
                converter()
