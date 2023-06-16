from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import pandas as pd
import random
import lxml
from random import randint
from selenium.webdriver.common.by import By
import time
import openpyxl

# browser = webdriver.Firefox()
browser = webdriver.Chrome()
time.sleep(5)  # долго грузится - делаем задержку
browser.get('https://www.list-org.com/')

for x in range(2, 10):  # 2 - A2 ячейка, 187 - A186 ячейка.
    wb = openpyxl.load_workbook('ИНН.xlsx')
    # sheet=wb.get_active_sheet()
    sheet = wb['Лист1']
    # sheet = wb.get_sheet_by_name('Лист1')

    # получаем кортеж из ОГРН в ячейке A2
    a = tuple(str(sheet.cell(row=x,
                             column=1).value).strip())

    # поиск строки с ИНН
    act = browser.find_element(By.NAME, "val")
    # перевод курсора в сроку ИНН
    act.click()
    # пауза 3 секунды
    time.sleep(3)
    # вводим посимвольно в строку ИНН, т.к. ввод сразу всего ИНН не корректно
    # обрабатывается
    i = 0
    for i in range(10):
        act.send_keys(a[i])
        time.sleep(round(random.uniform(0.2, 0.5), 2))
        i += 1
    # поиск строки с кнопкой "поиск"
    act = browser.find_element(By.CSS_SELECTOR, '.btn-primary')
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # переход по ссылке
    act.click()
    # пауза от 3 до 4 секунд
    time.sleep(randint(3, 4))
    act = browser.find_element(By.XPATH,
                               "/html/body/div[2]/div[1]/div[1]/div/p/label/a")

    time.sleep(randint(1, 3))
    act.click()
    # пауза от 3 до 4 секунд
    time.sleep(randint(3, 4))

    act = browser.find_element(By.XPATH,
                               "/html/body/div[2]/div[1]/div[1]/a[3]")
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # переход по ссылке
    act.click()

    # Create an URL object
    url = browser.current_url
    headers1 = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_4) AppleWebKit/537.36 '
                      '(KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36',
        'Referer': 'https://www.yandex.ru/'}
    # Create object page
    page = requests.get(url, headers=headers1)

    # parser-lxml = Change html to Python friendly format
    # Obtain page's information
    soup = BeautifulSoup(page.text, 'lxml')
    # # Obtain information from tag <table>
    table1 = soup.find("table", id="rep_table")

    # # Obtain every title of columns with tag <th>
    headers = []
    for i in table1.find_all("td"):
        print('I = ', i)
        title = i.text
        print(title)
        headers.append(title)
    break

    # # пауза от 3 до 4 секунд
    # time.sleep(randint(3, 4))
    # browser.back()
    # # пауза от 1 до 3 секунд
    # time.sleep(randint(1, 3))
    # browser.back()
    # time.sleep(randint(1, 3))
    # # пауза от 1 до 3 секунд
    # browser.back()
    # time.sleep(randint(1, 3))
    # # пауза от 1 до 3 секунд

    # x += 1
browser.quit()
