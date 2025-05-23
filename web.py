"""Модуль для скачивания документов с веб-сайта."""
from typing import List, NoReturn
import os
import urllib3
import wget
from bs4 import BeautifulSoup
from textracter import extract_text_from_documents, input_files, output_txt

# Константы
URL_ROOT: str = 'https://www.pravovik24.ru'
URL: str = f"{URL_ROOT}/documents"

# Инициализация HTTP клиента
HTTP_CLIENT = urllib3.PoolManager()

def ensure_directories() -> None:
    """Создает входную и выходную директории, если они не существуют."""
    for directory in [input_files, output_txt]:
        if not os.path.exists(directory):
            os.makedirs(directory)

def download_documents() -> None:
    """Скачивает документы с веб-сайта."""
    response = HTTP_CLIENT.request('GET', URL)
    soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')

    for link in soup.find_all('a'):
        href = link.get('href')
        if not href:
            continue

        if '/documents/dogovory/' in href:
            process_dogovory_page(href)
        elif 'documents' in href:
            process_documents_page(href)

def process_dogovory_page(href: str) -> None:
    """Обрабатывает страницу договоров и скачивает документы.
    
    Args:
        href: URL страницы договоров
    """
    response = HTTP_CLIENT.request('GET', f"{URL_ROOT}{href}")
    soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')
    
    if 'documents' in href:
        response = HTTP_CLIENT.request('GET', f"{URL_ROOT}{href}")
        soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')
    
    download_files_from_soup(soup)

def process_documents_page(href: str) -> None:
    """Обрабатывает страницу документов и скачивает файлы.
    
    Args:
        href: URL страницы документов
    """
    response = HTTP_CLIENT.request('GET', f"{URL_ROOT}{href}")
    soup = BeautifulSoup(response.data.decode('utf-8'), 'html.parser')
    download_files_from_soup(soup)

def download_files_from_soup(soup: BeautifulSoup) -> None:
    """Скачивает файлы из объекта BeautifulSoup.
    
    Args:
        soup: Объект BeautifulSoup с HTML-контентом
    """
    for link in soup.find_all('a'):
        href = link.get('href')
        if href and 'upload' in href:
            wget.download(f"{URL_ROOT}{href}", out=input_files)

def main() -> NoReturn:
    """Выполняет основной процесс скачивания документов."""
    ensure_directories()
    download_documents()
    extract_text_from_documents(input_files, output_txt)

if __name__ == '__main__':
    main()
