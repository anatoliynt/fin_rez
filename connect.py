from selenium import webdriver
from selenium.webdriver.common.keys import Keys
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
    #sheet = wb.get_sheet_by_name('Лист1')

    a = tuple(str(sheet.cell(row=x,
                             column=1).value).strip())  # получаем кортеж из ОГРН в ячейке A2
    # act = browser.find_element_by_id('query')
    act = browser.find_element(By.NAME, "val")
    act.click()
    time.sleep(3)
    # вводим посимвольно в строку ОГРН, т.к. ввод сразу всего ОГРН не корректно обрабатывается
    i = 0
    for i in range(10):
        act.send_keys(a[i])
        time.sleep (0.1)
        i += 1
    act = browser.find_element(By.CSS_SELECTOR, '.btn-primary')
    time.sleep(0.5)
    act.click()
    time.sleep(2)
    act = browser.find_element(By.XPATH,
                               "/html/body/div[2]/div[1]/div[1]/div/p/label/a")
    time.sleep(0.5)
    act.click()
    time.sleep(2)
    act = browser.find_element(By.XPATH,
                               "/html/body/div[2]/div[1]/div[1]/a[3]")
    time.sleep(0.5)
    act.click()
    time.sleep(2)
    browser.back()
    time.sleep(0.5)
    browser.back()
    time.sleep(0.5)
    browser.back()
    time.sleep(0.5)
    # act = browser.find_element(By.XPATH, '//*[@title="В начало"]')
    # time.sleep(0.5)
    # act.click()
    # удаляем старый ОГРН
    # i = 0
    # for i in range(10):
    #     act.send_keys(Keys.BACK_SPACE)
    #     i += 1

    x += 1
browser.quit()
