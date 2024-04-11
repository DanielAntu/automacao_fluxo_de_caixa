from pathlib import Path
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
import os
from openpyxl import load_workbook

load_dotenv()

USER = os.environ.get('USERS', '')
PASSWORD = os.environ.get('PASSWORD', '')

ROOT_PATCH = Path(__file__).parent
CHROMEDRIVER_NAME = 'chromedriver.exe'
CHROMEDRIVER_PATH = ROOT_PATCH / 'bin' / CHROMEDRIVER_NAME

def make_browser(*options):
    chrome_options = webdriver.ChromeOptions()

    if options is not None:
        for option in options:
            chrome_options.add_argument(option)

    chrome_service = Service(executable_path=CHROMEDRIVER_PATH)
    browser = webdriver.Chrome(service=chrome_service, options=chrome_options)

    return browser

def preacher_form(browser, xpath, value):
    browser.find_element(By.XPATH, xpath).send_keys(value)

# openpyxl
workbook = load_workbook('dados_registro.xlsx')
sheet = workbook['Sheet1']

browser = make_browser()
browser.get('http://127.0.0.1:8000/')

# login
preacher_form(browser, '//*[@id="id_username"]', USER)
sleep(1)
preacher_form(browser, '//*[@id="id_password"]', PASSWORD)
sleep(1)
browser.find_element(By.XPATH, '/html/body/main/div/form/input[2]').click()

# registrar fluxo
for row in sheet.iter_rows(min_row=2, values_only=True):
    description, nature, type_cash, value = row
    preacher_form(browser, '//*[@id="id_description"]', description)
    preacher_form(browser, '//*[@id="id_nature"]', nature)
    if type_cash is not None:
        preacher_form(browser, '//*[@id="id_type_cash"]', type_cash)
    preacher_form(browser, '//*[@id="id_value"]', value)
    browser.find_element(By.XPATH, '/html/body/main/div[2]/form/input[2]').click()
    sleep(2)
    

sleep(5)
browser.quit()