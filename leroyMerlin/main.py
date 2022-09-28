import requests
import json
import re
import sys
import time
from random import randint
from fake_useragent import UserAgent
import pandas as pd
import datetime as dt
from typing import Union
import PySimpleGUI as sg
from pathlib import Path
from config import PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS, MIN_DELAY, MAX_DELAY
from threading import Thread

from seleniumwire import undetected_chromedriver as uc

# sg.theme_previewer()
# exit(0)

sg.theme('Default1')
input_filename = ''
output_filename = ''


def get_qrator_id(proxy: Union[None, str] = None) -> str:
    if proxy:
        chrome_options = {
            'proxy': {
                'http': f'http://{proxy}',
                'https': f'https://{proxy}',
                'no_proxy': 'localhost,127.0.0.1'
            }
        }
    else:
        chrome_options = {}

    driver = uc.Chrome(use_subprocess=True, seleniumwire_options=chrome_options)
    driver.get('https://leroymerlin.ru/')
    time.sleep(5)
    cookies = driver.get_cookies()
    driver.quit()

    qrator_jsid = ''
    for cookie in cookies:
        if cookie['name'] == 'qrator_jsid':
            qrator_jsid = cookie['value']
            print('Qrator: ' + qrator_jsid)

    if qrator_jsid == '':
        print('Cannot parse qrator_jsid cookie from site')
        sys.exit()

    return str(qrator_jsid)


def get_regions() -> dict:
    with open('nregions.json', 'r', encoding='utf-8') as f:
        regions = json.load(f)
    return regions


def convert_excel_input_to_dict(filename: str = 'input.xlsx') -> dict:
    excel = pd.ExcelFile(filename)
    sheets = [sheet for sheet in excel.sheet_names if sheet in get_regions().keys()]

    excel_dfs = {}
    for i in sheets:
        df = pd.read_excel(filename, sheet_name=i, header=None)
        df.columns = [i]
        try:
            df[i] = df[i].astype(str).str.extract(r'\D*(\d+)').astype(int)
        except Exception as ex:
            print(ex)
        excel_dfs.update(df.to_dict(orient='list'))

    return excel_dfs


def parse_item_page(page: str) -> tuple:
    parsed_price = re.findall('"main_price":\s*(\d+[.]?\d*)', page)
    parsed_name = re.findall('"displayedName":\s*"([^\"]*)', page)
    not_found = re.findall('ничего не найдено', page)

    try:
        price = float(parsed_price[0])
    except ValueError:
        price = parsed_price[0]
    except IndexError:
        if not_found:
            price = 'Не найден'
        else:
            price = "Не удалось спарсить"

    try:
        name = parsed_name[0]
    except IndexError:
        if not_found:
            name = 'Не найден'
        else:
            name = "Не удалось спарсить"

    return price, name


def create_headers(qrator_jsid: str) -> dict:
    ua = UserAgent()
    # u_agents = [random.choice((ua.chrome, ua.firefox, ua.safari, ua.opera)) for x in range(10)]

    headers = {'user-agent': ua.chrome, #'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:104.0) Gecko/20100101 Firefox/104.0',
           'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
           'sec-ch-ua': 'Google Chrome";v="105", "Not)A;Brand";v="8", "Chromium";v="105',
           'sec-ch-ua-mobile': '?0',
           'sec-ch-ua-platform': 'Windows',
           'sec-fetch-dest': 'document',
           'sec-fetch-mode': 'navigate',
           'sec-fetch-site': 'none',
           'sec-fetch-user': '?1',
           'cookie': 'qrator_jsid='+qrator_jsid}

    return headers


def get_proxy_dict() -> dict:
    proxies = {'https': f'http://{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{PROXY_PORT}'}
    return proxies


def process_inputs_dict(inputs: dict, window, session, regions: dict) -> list:
    global PARSING_IS_STOPPED
    sg.cprint_set_output_destination(window, '-ML-'+sg.WRITE_ONLY_KEY)
    all_inputs_len = sum([len(inputs[region]) for region in inputs])

    output_records = []
    counter_total = 0

    for key in inputs:
        api_url = regions[key]['site']
        counter_in_region = 0
        sg.cprint('', key='-ML-'+sg.WRITE_ONLY_KEY)

        for article in inputs[key]:
            url = 'https://' + api_url + '/search/?q=' + str(article)

            resp = session.get(url)

            price, name = parse_item_page(resp.text)
            output_records.append((article, key, name, price, str(dt.datetime.today().date())))
            counter_in_region += 1
            counter_total += 1

            screen_text = f'{counter_in_region:3d}. {key}: {article} - {price} - {name}'
            sg.cprint(screen_text, key='-ML-'+sg.WRITE_ONLY_KEY)
            progress = round(counter_total/all_inputs_len*100)
            print(progress)
            window['-PBAR-'].update(progress)

            delay = randint(MIN_DELAY, MAX_DELAY)/10
            finish = time.time() + delay

            while time.time() < finish:
                if PARSING_IS_STOPPED:
                    return output_records
                time.sleep(0.2)

    return output_records


def requesting(inputs, window, regions):
    qrator_jsid = get_qrator_id(f'{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{PROXY_PORT}')
    headers = create_headers(qrator_jsid)
    proxies = get_proxy_dict()

    with requests.Session() as session:
        session.headers.update(headers)
        session.proxies.update(proxies)
        st = time.time()
        output_records = process_inputs_dict(inputs, window, session, regions)
        sg.cprint('', key='-ML-'+sg.WRITE_ONLY_KEY)
        sg.cprint(f'Время выполнения: {round((time.time()-st)/60, 1)} минут', key='-ML-'+sg.WRITE_ONLY_KEY)
        sg.cprint('', key='-ML-'+sg.WRITE_ONLY_KEY)
        print(f'Время выполнения: {time.time()-st}')

        df = pd.DataFrame(output_records, columns=['Артикул', 'Город', "Название", "Цена", 'Дата запроса'])
        df.to_excel(output_filename, index=False)
        sg.cprint(f'Файл: {output_filename} сохранен', key='-ML-'+sg.WRITE_ONLY_KEY)


def get_window():
    column_to_be_centered = [
        [sg.Frame("Данные для обработки", element_justification = "center", expand_x=True, layout = [
        [sg.Text("Файл для парсинга: "), sg.Input(size=(10, 1), enable_events=True, key='-FILENAME-', expand_x=True), sg.FileBrowse(button_text='Выбрать', key="-IN-", file_types=(("Excel", "*.xlsx"), ))],
        [sg.T("Количество артикулов по регионам")],
        [sg.Table([], col_widths=[15, 10], num_rows=6,  headings=['Регион', 'Кол-во SKU'], key='-TABLE-', auto_size_columns=False)]
            ])],
        [sg.B("Запуск парсинга", disabled=True), sg.B('СТОП', disabled=True)],
        [sg.Frame('Статус обработки', element_justification = "center", layout = [
        [sg.T('Прогресс: '), sg.ProgressBar(100, orientation='h', expand_x=True, size=(10, 12), key='-PBAR-', bar_color=('blue', 'white'), relief='RELIEF_FLAT', border_width=1)],
        [sg.Multiline(key='-ML-'+sg.WRITE_ONLY_KEY, size=(145, 25), auto_refresh=True),]])]]
    # ]

    layout = [[sg.VPush()],
              [sg.Push(), sg.Column(column_to_be_centered, element_justification='c'), sg.Push()],
              [sg.VPush()]
        ]

    icon = b'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'
    window = sg.Window(title='Поиск цен на Leroymerlin', layout=layout, size=(650, 460), icon=icon)

    return window


def main():
    window = get_window()
    global input_filename
    global output_filename
    global PARSING_IS_STOPPED

    while True:
        event, values = window.read(timeout=100)
        if event in (sg.WIN_CLOSED, "Exit", None):
            break
        elif event == 'Запуск парсинга':
            PARSING_IS_STOPPED = False
            window['Запуск парсинга'].update(disabled=True)
            window['СТОП'].update(disabled=False)
            thread = Thread(target=requesting, args=(inputs, window, regions))
            thread.start()

        elif event == 'СТОП':
            PARSING_IS_STOPPED = True
            window['Запуск парсинга'].update(disabled=False)
            window['СТОП'].update(disabled=True)

        elif event == '-FILENAME-':
            input_filename = values['-FILENAME-']
            output_filename = (Path(input_filename).parent) / (Path(input_filename).stem + ' - output' + Path(input_filename).suffix)

            inputs = convert_excel_input_to_dict(filename=input_filename)
            if inputs:
                regions = get_regions()
                tabs = [[region, len(inputs[region])] for region in inputs]
                all_inputs_len = sum([len(inputs[region]) for region in inputs])
                window['-TABLE-'].update(values=tabs)
                window['Запуск парсинга'].update(disabled=False)
            else:
                window['Запуск парсинга'].update(disabled=True)
                window['-TABLE-'].update(values=[])
                sg.popup_error('Данный файл не содержит листов с названиями регионов!')

    print('Done')


if __name__ == '__main__':
    main()
