from bs4 import BeautifulSoup
import urllib3
import wget
from textracter import converter, input_directory

url_root = 'https://www.pravovik24.ru'
url = "https://www.pravovik24.ru/documents"

http = urllib3.PoolManager()
links = []

response = http.request('GET', url)

soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')
for link in soup.findAll('a'):
    buf = link.get('href')
    if 'documents' in buf:
        response1 = http.request('GET', url_root+buf)
        soup2 = BeautifulSoup(response1.data.decode('utf-8'), 'html.parser')
        for link in soup2.findAll('a'):
            buf = link.get('href')
            if 'upload' in buf:
                response = wget.download(url_root+buf, out=input_directory)
                converter()
