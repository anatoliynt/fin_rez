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



url = "https://www.list-org.com/company/7086383/report"
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_4) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36',
    'Referer': 'https://www.yandex.ru/'}
# Create object page
page = requests.get(url, headers=headers)

# parser-lxml = Change html to Python friendly format
# Obtain page's information
soup = BeautifulSoup(page.text, 'lxml')
print(soup.h1)