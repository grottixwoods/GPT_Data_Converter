from bs4 import BeautifulSoup
import urllib3
import wget
from textracter import converter

url_root = 'https://www.pravovik24.ru'
url = "https://www.pravovik24.ru/documents/akty/"

http = urllib3.PoolManager()
links = []

response = http.request('GET', url)

soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')
for link in soup.findAll('a'):
    buf = link.get('href')
    if 'upload' in buf:
        response = wget.download(url_root+buf)
        converter()
