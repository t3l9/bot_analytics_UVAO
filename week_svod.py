import os
import re
import time
import asyncio
from datetime import datetime, timedelta
from functools import reduce

import pandas as pd
from dotenv import load_dotenv

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# Telegram
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    InputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
    ConversationHandler,
    CallbackQueryHandler,
    CallbackContext,
)

# Excel / Windows
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule
import win32com.client
import pythoncom


load_dotenv()
login_MM = os.getenv("login_MM")
password_MM = os.getenv("password_MM")
# Получаем текущую директорию проекта
project_dir = os.path.dirname(os.path.abspath(__file__))
# Создаем папку data внутри проекта, если её нет
data_dir = os.path.join(project_dir, "data")
os.makedirs(data_dir, exist_ok=True)


def parcing_data_MM_sync(MM_start_date, MM_end_date):
    """Синхронная версия функции парсинга данных"""
    chrome_install = ChromeDriverManager().install()
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(chromedriver_path))
    driver.maximize_window()
    try:
        driver.get('https://arm-mmonitor.mos.ru')
        time.sleep(0.5)
        username = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[1]/div/input')
        password = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[2]/div/input')
        username.send_keys(login_MM)
        password.send_keys(password_MM)
        login_button = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[5]/div[1]/button')
        login_button.click()

        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div/section/section/main/div/div[1]/div[2]/span[1]')))
        time.sleep(0.3)
        button = driver.find_element(By.XPATH, "/html/body/div[1]/div/section/section/main/div/div[1]/div[2]/span[1]")
        button.click()

        button = driver.find_element(By.XPATH,
                                     "/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/span/div/div")
        button.click()
        time.sleep(0.5)
        button = driver.find_element(By.XPATH,
                                     "/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/span/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[3]/div")
        button.click()
        time.sleep(1)
        button1 = driver.find_element(By.XPATH,
                                      '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/div/div[1]/div/input')
        button1.click()
        button1.send_keys(Keys.CONTROL + 'a')
        button1.send_keys(Keys.BACKSPACE)
        time.sleep(0.3)
        button1.send_keys(MM_start_date)
        time.sleep(0.5)
        button2 = driver.find_element(By.XPATH,
                                      '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/div/div[2]/div/input')
        button2.click()
        button2.send_keys(Keys.CONTROL + 'a')
        button2.send_keys(Keys.BACKSPACE)
        time.sleep(0.3)
        button2.send_keys(MM_end_date)

        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[13]/div/div[1]/div')
        button.click()
        time.sleep(0.5)
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[13]/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[2]/div')
        button.click()
        time.sleep(0.5)
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[2]/button[1]')
        button.click()
        time.sleep(0.5)
        body = driver.find_element(By.TAG_NAME, 'body')
        body.click()
        time.sleep(0.5)
        button = driver.find_element(By.CSS_SELECTOR, "svg.icon.xls-icon")
        button.click()
        time.sleep(0.5)
        driver.get('https://arm-mmonitor.mos.ru/#/export-files')
        i = 0
        while i < 50:
            try:
                element = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH,
                                                                                           "/html/body/div/div/section/section/main/div/div/div[1]/div/div/div/div/div[2]/table/tbody/tr[3]/td[5]/div/button/span")))
                time.sleep(1)
                element.click()
                print("Элемент найден, прекращаем обновление страницы.")
                break
            except:
                print("Элемент не найден, обновляем страницу.")
                driver.refresh()
                i += 1
                print(i)
                time.sleep(3)
        time.sleep(6)
        return True
    except Exception as e:
        print(f"❌ Произошла ошибка при выгрузке ММ: {e}")
        return False
    finally:
        driver.quit()


async def parcing_data_MM_async(start_date, end_date):
    """Асинхронная версия функции парсинга данных"""
    try:
        loop = asyncio.get_event_loop()
        success = await loop.run_in_executor(
            None,
            lambda: parcing_data_MM_sync(start_date, end_date)
        )
        return success
    except Exception as e:
        print(f"Ошибка в асинхронной функции парсинга: {e}")
        return False


def process_file_MM_week(first_file, second_file):
    """Обработка файлов для еженедельного свода"""
    today = datetime.now()
    last_week_monday = today - timedelta(days=today.weekday() + 7)
    last_week_sunday = last_week_monday + timedelta(days=6)

    report_period = f"{last_week_monday.strftime('%d.%m.%Y')}-{last_week_sunday.strftime('%d.%m.%Y')}"
    report_start_date = last_week_monday.date()
    report_end_date = last_week_sunday.date()

    responsible_mapping = {
        'ГБУ «Автомобильные дороги ЮВАО»': 'АВД ЮВАО',
        'ГБУ «Жилищник Выхино района Выхино-Жулебино»': 'Выхино-Жулебино',
        'Управа района Выхино-Жулебино': 'Выхино-Жулебино',
        'ГБУ «Жилищник Нижегородского района»': 'Нижегородский',
        'Управа Нижегородского района': 'Нижегородский',
        'ГБУ «Жилищник района Капотня»': 'Капотня',
        'Управа района Капотня': 'Капотня',
        'ГБУ «Жилищник района Кузьминки»': 'Кузьминки',
        'Управа района Кузьминки': 'Кузьминки',
        'ГБУ «Жилищник района Лефортово»': 'Лефортово',
        'Управа района Лефортово': 'Лефортово',
        'ГБУ «Жилищник района Люблино»': 'Люблино',
        'Управа района Люблино': 'Люблино',
        'ГБУ «Жилищник района Марьино»': 'Марьино',
        'Управа района Марьино': 'Марьино',
        'ГБУ «Жилищник района Некрасовка»': 'Некрасовка',
        'Управа района Некрасовка': 'Некрасовка',
        'ГБУ «Жилищник района Печатники»': 'Печатники',
        'Управа района Печатники': 'Печатники',
        'ГБУ «Жилищник района Текстильщики»': 'Текстильщики',
        'Управа района Текстильщики': 'Текстильщики',
        'ГБУ «Жилищник Рязанского района»': 'Рязанский',
        'Управа Рязанского района': 'Рязанский',
        'ГБУ «Жилищник района Южнопортовый»': 'Южнопортовый',
        'Управа Южнопортового района': 'Южнопортовый'
    }

    output_file_path = os.path.join(data_dir, f'Все {report_period}.xlsx')

    # Обработка первого файла
    def process_first_file(filepath):
        df = pd.read_excel(filepath, sheet_name='КП_БП')
        df = df[df['Округ'] == 'ЮВАО']

        df1 = pd.read_excel(filepath, sheet_name='Первичные данные')
        df1 = df1[df1['Округ'] == 'ЮВАО']

        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            # Лист КП_БП
            df.to_excel(writer, sheet_name='КП_БП', index=False)

            # Лист Просроки
            df1['Район'] = df1['Ответственный исполнитель'].map(responsible_mapping)
            df_filtered = df1[df1['Район'].notnull()]
            df_filtered.to_excel(writer, sheet_name='Просроки', index=False)

            # Лист Новые просроки
            df_filtered_new = df_filtered[
                (df_filtered['Срок устранения до'].notnull()) &
                (df_filtered['Срок устранения до'].apply(
                    lambda x: pd.notna(x) and report_start_date <= pd.to_datetime(x).date() <= report_end_date))
                ]
            df_filtered_new = pd.concat([df_filtered_new, df_filtered[df_filtered['Срок устранения до'].isnull()]])

            df_filtered_new.to_excel(writer, sheet_name='Новые просроки', index=False)

            # Лист Снег
            df_filtered_snow = df_filtered[
                df_filtered['Срок устранения до'].isnull()]
            df_filtered_snow.to_excel(writer, sheet_name='Снег', index=False)

    # Обработка второго файла
    def process_second_file(filepath):
        df = pd.read_excel(filepath)

        df['Район'] = df['Ответственный исполнитель'].map(responsible_mapping)
        df_filtered = df[df['Район'].notnull()]

        with pd.ExcelWriter(output_file_path, mode='a', engine='openpyxl') as writer:
            # Лист Поступившие в отчетном
            df_in_report = df_filtered[
                (df_filtered['Дата фиксации нарушения'].notnull()) &
                (df_filtered['Дата фиксации нарушения'].apply(
                    lambda x: pd.notna(x) and report_start_date <= pd.to_datetime(x).date() <= report_end_date))
                ]
            df_in_report.to_excel(writer, sheet_name='Поступившие в отчетном', index=False)

            # Лист На исполнении в отчетном
            df_on_execution = df_filtered[
                (df_filtered['Срок устранения до'].notnull()) &
                (df_filtered['Срок устранения до'].apply(
                    lambda x: pd.notna(x) and report_start_date <= pd.to_datetime(x).date() <= report_end_date))
                ]
            df_on_execution.to_excel(writer, sheet_name='На исполнении в отчетном', index=False)

    # Выполняем обработку
    process_first_file(first_file)
    process_second_file(second_file)

    return output_file_path