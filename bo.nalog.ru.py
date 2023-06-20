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
    # пауза 3-4 секунд после подключения к сайту
    time.sleep(randint(3, 4))
    # ищем информационное окно "Ресурс сформирован на основании информации,
    #                                  представленной составителями отчетности"
    act = browser.find_element(By.XPATH,
                               '//div[@class = "modal-content short-info-bottom"]'
                               '/child::button')
    # закрываем окно
    act.click()
    # поиск строки с ИНН
    act = browser.find_element(By.ID, 'search')
    # пауза 2-3 секунду
    time.sleep(randint(2, 3))
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
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # поиск ссылки на выбранной компании
    act = browser.find_element(By.XPATH,
                               '//div[@class = "results-search-tbody"]/child::a')
    # переходим по ссылке выбранной компании
    act.click()
    # пауза от 8 до 10 секунд
    time.sleep(randint(8, 10))
    # поиск Скачать таблицей или текстом года
    act = browser.find_element(By.XPATH,
                               '//button[text() = "Скачать таблицей или текстом"]')
    # пауза от 2 до 4 секунд
    time.sleep(randint(2, 4))
    # переход по ссылке
    act.click()
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # находим запись Выбрать все
    act = browser.find_element(By.XPATH, '//button[text() = "Выбрать все"]')
    # отмечаем все отчеты
    act.click()
    # пауза от 1 до 3 секунд
    time.sleep(randint(1, 3))
    # переменная год для цикла
    year = [2019, 2020, 2021, 2022]
    for i_year in year:
        # выбираем 2019 год
        act = browser.find_element(By.XPATH, '//button[@data-year = "{0}"]'.format(i_year))
        # нажимаем по 2019
        act.click()
        # пауза от 1 до 3 секунд
        time.sleep(randint(1, 3))
        # поиск кнопки скачать архив
        act = browser.find_element(By.XPATH, '//span[text() = "Скачать архив"]')
        act.click()
        # пауза от 8 до 10 секунд
        time.sleep(randint(8, 10))

    x += 1
browser.quit()
