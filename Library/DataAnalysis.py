import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import json
import zipfile
import tempfile
import urllib.request
import PySimpleGUI as sg


keys = []
dir_list = []


def forming_dict():
    """Функция чтение CSV файла, формирование словаря, в котором хранится список всех компаний
     и ссылки на сайт с их отчетностями. Функция возвращает сформированный словарь. Владимир Никитин."""
    df = pd.read_csv(os.path.abspath('./Data/Data.csv'), encoding='1251')
    pd.set_option('display.max_columns', None)
    df = df[['EMITENT_FULL_NAME', 'DISCLOSURE_RF_INFO_PAGE']].drop_duplicates()

    df['DISCLOSURE_RF_INFO_PAGE'].fillna(0, inplace=True)
    inform_dict = df.groupby(['EMITENT_FULL_NAME'])['DISCLOSURE_RF_INFO_PAGE'].apply(list).to_dict()
    url_dict = {}
    for elem in inform_dict.keys():
        if inform_dict[elem][0] != 0 and "www.e-disclosure.ru/" in inform_dict[elem][0] and \
                (inform_dict[elem][0] != 'https://www.e-disclosure.ru/' and
                 inform_dict[elem][0] != 'www.e-disclosure.ru' and
                 inform_dict[elem][0] != 'http://www.e-disclosure.ru/'):
            url_dict[elem] = inform_dict[elem][0]
    return url_dict


def saving(url):
    """Функция скачивания архивов и их разархивации. На вход функции подается словрь с компаниями
    и ссылками на страницу сайта с их отчетностями, в процессе скачивания появляетс окно с полоской прогресса.
     Владимир Никитин и Матанов Кирилл."""
    k = 0
    progress = 0
    if not os.path.exists(os.path.abspath('./Output')):
        os.mkdir(os.path.abspath('./Output'))
    progressbar = [[sg.ProgressBar(len(url), orientation='h', size=(51, 10), key='progressbar')]]
    layout = [[sg.Frame('Прогресс', layout=progressbar)], [sg.Button('Старт')]]
    window = sg.Window('Скачивание архивов', layout)
    progress_bar = window['progressbar']
    event, values = window.read()
    if event == 'Старт':
        for elem in url.keys():
            name = elem
            while '"' in name:
                name = elem.replace('"', '')
            name = name.lstrip()
            name = name.rstrip()
            if os.path.exists(os.path.abspath('./Output') + '/' + name):
                pass
            else:
                os.mkdir(os.path.abspath('./Output') + '/' + name)
            k = k + 1
            link = url[elem]
            response = requests.get(link)
            soup = BeautifulSoup(response.text, 'lxml')
            if "e-disclosure" in url[elem]:
                quotes = soup.findAll("a", {"class": "file-link"})
                for elements in quotes:
                    string_united = ''
                    string = str(elements)
                    for i in string:
                        string_united = string_united + i
                    beg_url = string_united.find('href="http')+6
                    end_url = string_united.find('">')
                    response = requests.get(string_united[beg_url:end_url])
                    file = tempfile.TemporaryFile()
                    file.write(response.content)
                    try:
                        fzip = zipfile.ZipFile(file)
                        while '"' in elem:
                            elem = elem.replace('"', ' ')
                        fzip.extractall(os.path.abspath('./Output') + "/" + name)
                        file.close()
                        fzip.close()
                    except zipfile.BadZipFile:
                        pass
            progress = progress + 1
            progress_bar.UpdateBar(progress)
    window.close()


def site_parsing(urls):
    """Функция формирования словаря, содержащего только компании, имеющие страницу с отчетностью
     и ссылки на эти страницы. На вход подается словарь всех компаний и ссылок на главную страницу их сайта
     Владимир Никитин и Матанов Кирилл."""
    progressbar = [[sg.ProgressBar(len(urls), orientation='h', size=(51, 10), key='progressbar')]]
    layout = [[sg.Frame('Прогресс', layout=progressbar)], [sg.Button('Старт')]]
    window = sg.Window('Обработка списка компаний', layout)
    progress_bar = window['progressbar']
    event, values = window.read()
    otch_dict = {}
    k = 0
    progress = 0
    if event == 'Старт':
        for elem in urls.keys():
            progress = progress + 1
            progress_bar.UpdateBar(progress)
            if ';' in urls[elem]:
                list_of_urls = urls[elem].split(';')
                for i in list_of_urls:
                    if 'www.e-disclosure.ru/' in i:
                        urls[elem] = i
            k = k + 1
            if " " in urls[elem]:
                urls[elem] = urls[elem].replace(" ", '')
            if "https:" not in urls[elem] and "http:" not in urls[elem]:
                urls[elem] = "https://" + urls[elem]
            url = urls[elem]
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'lxml')
            quotes = soup.find_all('a')
            for elements in quotes:
                string_united = ''
                string = str(elements)
                for i in string:
                    string_united = string_united + i
                if "Год" in string_united:
                    first = string_united.find('"')
                    second = string_united.find('">')
                    first_amp = string_united.find('amp')
                    second_amp = string_united.find(';type')
                    string_united = string_united.replace(string_united[first_amp:second_amp+1], '')
                    otch_dict[elem] = "https://e-disclosure.ru/"+string_united[first+1:second-4]
    window.close()
    return otch_dict


def make_file(data):
    """Функция создания json файла, содержащего словарь с компаниями и ссылками на страницу их сайта с отчетностью.
     На вход подается слвоарь, содержащий все эти данные. Владимир Никитин"""
    with open(os.path.abspath('./Data/List.json'), "w") as f:
        json.dump(data, f)


def read_file():
    """""Функция чтения json файла, содержащего словарь с компаниями и ссылками на страницу их сайта с отчетностью.
    Матанов Кирилл"""
    with open(os.path.abspath('./Data/List.json'), "r") as f:
        list_info = json.load(f)
    return list_info


def make_dir_file():
    """Функция создания файла Folders.txt, содержащий список папок Владимир Никитин"""
    dir_list = os.listdir(os.path.abspath('./Output'))
    open(os.path.abspath('./Output/Folders.txt'), "w").close()
    with open(os.path.abspath('./Output/Folders.txt'), "w") as f:
        for elem in dir_list:
            if elem != 'Folders.txt':
                f.write(elem + '\n')


def download_listing():
    """Матанов Кирилл"""
    url = "https://www.moex.com/ru/listing/securities-list.aspx"
    link = ""
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    quotes = soup.findAll("a")
    for elem in quotes:
        string_united = ''
        string = str(elem)
        for i in string:
            string_united = string_united + i
        if ">CSV (разделители - запятые)" in string_united:
            first = string_united.find('="') + 2
            last = string_united.find(">CSV (разделители - запятые)") - 1
            link = string_united[first:last]
            link = "https://www.moex.com/ru/listing/" + link
    destination = os.path.abspath('./Data/Data.csv')
    urllib.request.urlretrieve(link, destination)
