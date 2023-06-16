import clickable as clickable
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.keys import Keys
import requests
from bs4 import BeautifulSoup
import pandas as pd
import random
from random import randint
from selenium.webdriver.common.by import By
import time
import openpyxl

# Переменная browser для подключения к Хрому
browser = webdriver.Chrome()
# долго грузится - делаем задержку
time.sleep(2)
# открываем ссылку Налоговых отчетов
browser.get('https://bo.nalog.ru/')


# цикл по заданному списку ИНН
# 2 - A2 ячейка, 10 - A9 ячейка.
for x in range(2, 10):
    # открываем файл с ИНН
    wb = openpyxl.load_workbook('ИНН.xlsx')
    # указываем активный лист
    sheet = wb['Лист1']
    # получаем кортеж из ИНН в ячейке по циклу начиная с А2
    a = tuple(str(sheet.cell(row=x,
                             column=1).value).strip())
    # пауза 5 секунд после подключения к сайту
    time.sleep(5)
    # ищем информационное окно "Ресурс сформирован на основании информации,
    #                                  представленной составителями отчетности"
    act = browser.find_element(By.XPATH,
                               '//div[@class = "modal-content short-info-bottom"]'
                               '/child::button')
    # закрываем окно
    act.click()
    # поиск строки с ИНН
    act = browser.find_element(By.ID, 'search')
    # пауза 3 секунду
    time.sleep(3)
    # вводим посимвольно в строку ИНН, т.к. ввод сразу всего ИНН не корректно
    # обрабатывается
    i = 0
    for i in range(10):
        act.send_keys(a[i])
        time.sleep(round(random.uniform(0.2, 0.5), 2))
        i += 1
    # поиск строки с кнопкой "поиск"
    act = browser.find_element(By.CSS_SELECTOR, '.button_search')
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # переход по ссылке
    act.click()
    # пауза от 3 до 4 секунд
    time.sleep(randint(1, 3))
    # поиск ссылки на выбранной компании
    act = browser.find_element(By.XPATH, '//div[@class = "results-search-tbody"]/child::a')
    # переходим по ссылке выбранной компании
    act.click()
    # пауза от 3 до 4 секунд
    time.sleep(randint(10, 11))
    # поиск 2019 года
    act = browser.find_element(By.XPATH, '//button[text() = "2019"][1]')
    # пауза от 1 до 3 секунд
    time.sleep(randint(8, 10))
    # переход по ссылке
    act.click()
    #
    # get_url = browser.current_url
    #
    #
    #
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
    #

    x += 1
browser.quit()
