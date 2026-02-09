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


home_dir = os.path.expanduser("~")
directory = os.path.join(home_dir, "Downloads")
# Загружаем переменные из .env
load_dotenv()
login_MM = os.getenv("login_MM")
password_MM = os.getenv("password_MM")


# Монитор ММ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
async def parcing_data_MM(context, chat_id, MM_start_date, MM_end_date):
    chrome_install = ChromeDriverManager().install()
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(chromedriver_path))
    driver.maximize_window()
    try:
        # Откройте страницу логина
        driver.get('https://arm-mmonitor.mos.ru')
        time.sleep(0.5)
        # Найдите поля для ввода логина и пароля и заполните их
        username = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[1]/div/input')
        password = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[2]/div/input')
        username.send_keys(login_MM)
        password.send_keys(password_MM)
        # Найдите и нажмите кнопку логина
        login_button = driver.find_element(By.XPATH, '/html/body/main/div/div[2]/div/form[1]/div[5]/div[1]/button')
        login_button.click()

        # Подождите, пока страница загрузится

        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div/section/section/main/div/div[1]/div[2]/span[1]')))
        time.sleep(0.3)
        button = driver.find_element(By.XPATH, "/html/body/div[1]/div/section/section/main/div/div[1]/div[2]/span[1]")
        button.click()

        # выпадающая дата
        button = driver.find_element(By.XPATH,
                                     "/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/span/div/div")
        button.click()
        time.sleep(0.5)
        # ставим дата отчета
        button = driver.find_element(By.XPATH,
                                     "/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/span/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[3]/div")
        button.click()
        time.sleep(1)
        # enter start date
        button1 = driver.find_element(By.XPATH,
                                      '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/div/div[1]/div/input')
        button1.click()
        button1.send_keys(Keys.CONTROL + 'a')  # Выделить весь текст
        button1.send_keys(Keys.BACKSPACE)  # Удалить выделенный текст
        time.sleep(0.3)
        button1.send_keys(MM_start_date)  # дата начала вводится
        time.sleep(0.5)
        # enter end date
        button2 = driver.find_element(By.XPATH,
                                      '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[4]/div/div[2]/div/input')
        button2.click()
        button2.send_keys(Keys.CONTROL + 'a')  # Выделить весь текст
        button2.send_keys(Keys.BACKSPACE)  # Удалить выделенный текст
        time.sleep(0.3)
        button2.send_keys(MM_end_date)  # дата конца вводится

        # доходим до ответсвенных
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[13]/div/div[1]/div')
        button.click()
        time.sleep(0.5)
        # выбираем территориальные
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[1]/label[13]/div/div[2]/div/div/div/div[2]/div[1]/div/div/div[2]/div')
        button.click()
        time.sleep(0.5)
        # нажимаем показать
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[2]/div/div[2]/div/div/div[2]/div[2]/button[1]')
        button.click()
        # time.sleep(1000)
        time.sleep(0.5)
        # нажимаем на кнопку что бы закрыть фильтры
        body = driver.find_element(By.TAG_NAME, 'body')
        body.click()
        time.sleep(0.5)
        # добавляем в очередь скачивания
        button = driver.find_element(By.CSS_SELECTOR, "svg.icon.xls-icon")
        button.click()
        time.sleep(0.5)
        # переходим в загрузки
        driver.get('https://arm-mmonitor.mos.ru/#/export-files')
        # обновляем страницу пока не появится нужный элемент, затем скачиваем
        i = 0
        while i < 50:
            try:
                # Ожидание элемента в течение 5 секунд (без обновления страницы)
                element = WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH,
                                                                                           "/html/body/div/div/section/section/main/div/div/div[1]/div/div/div/div/div[2]/table/tbody/tr[3]/td[5]/div/button/span")))
                time.sleep(1)
                element.click()
                print("Элемент найден, прекращаем обновление страницы.")
                break  # Выход из цикла, если элемент найден
            except:
                print("Элемент не найден, обновляем страницу.")
                driver.refresh()  # Обновление страницы
                i += 1
                print(i)
                time.sleep(3)  # Ожидание 3 секунд перед следующей проверкой
        time.sleep(6)
        return True
    except Exception as e:
        error_message = f"❌Произошла ошибка при выгрузке ММ. Пожалуйста, попробуйте еще раз."
        print(error_message)  # Выводим ошибку в консоль
        await context.bot.send_message(chat_id=chat_id, text=error_message)  # Отправляем сообщение в Telegram
        return False
    finally:
        driver.quit()
def choosing_time_MM():
    today = datetime.now()
    current_date = pd.Timestamp(datetime.now().date())

    eight_am_today = current_date + pd.Timedelta(hours=0)
    ten_am_today = current_date + pd.Timedelta(hours=10, minutes=59, seconds=59)

    twelf_am_today = current_date + pd.Timedelta(hours=11)
    therteen_am_today = current_date + pd.Timedelta(hours=14, minutes=59, seconds=59)

    three_pm_today = current_date + pd.Timedelta(hours=15)
    five_am_today = current_date + pd.Timedelta(hours=19, minutes=59, seconds=59)

    eight_pm_today = current_date + pd.Timedelta(hours=20)
    eleven_pm_today = current_date + pd.Timedelta(hours=23, minutes=59, seconds=59)

    if (today > eight_am_today) & (today < ten_am_today):
        timenow = "УТРО"
    elif (today > twelf_am_today) & (today < therteen_am_today):
        timenow = "ДЕНЬ"
    elif (today > three_pm_today) & (today < five_am_today):
        timenow = "ВЕЧЕР"
    elif (today > eight_pm_today) & (today < eleven_pm_today):
        timenow = "НОЧЬ"
    return timenow
def first_attribute(df):
    today = datetime.now()
    weekday = today.weekday()
    # Определяем начало и конец текущей недели
    start_of_week = today - timedelta(days=weekday)
    end_of_week = start_of_week + timedelta(days=6)
    # Фильтруем DataFrame в соответствии с требуемой логикой
    if weekday == 0:
        df.loc[(df['Просрок'] == 'Да') & (df['Статус в системе'] == 'Устранено')
           & (df[
                  'Срок устранения до'].dt.date == today.date()), "ТипСПросроком"] = "Устранено с нарушением срока " + today.strftime(
        "%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 1:
        start_day = start_of_week + timedelta(days=(weekday - 1))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'Устранено'), "ТипСПросроком"] = "Устранено с нарушением срока " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 2:
        start_day = start_of_week + timedelta(days=(weekday - 2))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'Устранено'), "ТипСПросроком"] = "Устранено с нарушением срока " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 3:
        start_day = start_of_week + timedelta(days=(weekday - 3))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'Устранено'), "ТипСПросроком"] = "Устранено с нарушением срока " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 4:
        start_day = start_of_week + timedelta(days=(weekday - 4))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'Устранено'), "ТипСПросроком"] = "Устранено с нарушением срока " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 5:
        start_day = start_of_week + timedelta(days=(weekday - 5))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'Устранено'), "ТипСПросроком"] = "Устранено с нарушением срока " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (На текущей уб. неделе)"
    elif weekday == 6:
        start_day = today - timedelta(days=6)
        end_day = today
        df.loc[(df['Просрок'] == 'Да') & (df['Статус в системе'] == 'Устранено') & (df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()), "ТипСПросроком"] = ("Устранено с нарушением срока " + start_day.strftime("%d.%m.%y") + " по " + today.strftime("%d.%m.%y")) + " (На текущей уб. неделе)"
def second_attribute(df):
    today = datetime.now()
    weekday = today.weekday()
    # Определяем начало и конец текущей недели
    start_of_week = today - timedelta(days=weekday)
    if weekday == 6:
        start_day = today - timedelta(days=6)
        end_day = today
        df.loc[(df['Просрок'] == 'Да') & (df['Статус в системе'] == 'В работе') & (
                df['Срок устранения до'].dt.date >= start_day.date())
               & (df['Срок устранения до'].dt.date <= end_day.date()), "ТипСПросроком"] = (
                                                                                                  "В работе с просроком " + start_day.strftime(
                                                                                              "%d.%m.%y") + " по " + today.strftime(
                                                                                              "%d.%m.%y")) + " (Текущая уб. неделя)"

    elif weekday == 0:
        df.loc[(df['Просрок'] == 'Да') & (df['Статус в системе'] == 'В работе')
               & (df[
                      'Срок устранения до'].dt.date == today.date()), "ТипСПросроком"] = "В работе с просроком " + today.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 1:
        start_day = start_of_week + timedelta(days=(weekday - 1))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 2:
        start_day = start_of_week + timedelta(days=(weekday - 2))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 3:
        start_day = start_of_week + timedelta(days=(weekday - 3))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 4:
        start_day = start_of_week + timedelta(days=(weekday - 4))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 5:
        start_day = start_of_week + timedelta(days=(weekday - 5))
        end_day = today
        df.loc[(df['Срок устранения до'].dt.date >= start_day.date()) &
               (df['Срок устранения до'].dt.date <= end_day.date()) & (df['Просрок'] == 'Да') &
               (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком " + start_day.strftime(
            "%d.%m.%y") + " по " + today.strftime("%d.%m.%y") + " (Текущая уб. неделя)"
def third_attribute(df):
    today = datetime.now()
    weekday = today.weekday()
    # Определяем начало и конец текущей недели
    if weekday == 0:
        end_of_last_week = today - timedelta(days=1)
        start_of_last_week = end_of_last_week - timedelta(days=6)
        df.loc[(df['Срок устранения до'].dt.date >= start_of_last_week.date()) &
               (df['Срок устранения до'].dt.date <= end_of_last_week.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком с " + start_of_last_week.strftime(
            "%d.%m.%y") + " по " + end_of_last_week.strftime("%d.%m.%y") + " (Прошедшая уб. неделя)"
    else:
        end_of_last_week = today - timedelta(days=(weekday+1))
        start_of_last_week = end_of_last_week - timedelta(days=6)
        df.loc[(df['Срок устранения до'].dt.date >= start_of_last_week.date()) &
               (df['Срок устранения до'].dt.date <= end_of_last_week.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком с " + start_of_last_week.strftime(
            "%d.%m.%y") + " по " + end_of_last_week.strftime("%d.%m.%y") + " (Прошедшая уб. неделя)"
def fourth_attribute(df):
    today = datetime.now()
    weekday = today.weekday()
    # Определяем начало и конец текущей недели
    earliest_date = df['Срок устранения до'].min()
    # if weekday == 0:
    #     end_of_last_week = today - timedelta(days=7)
    #     end_of_last_week_mon = end_of_last_week - timedelta(days=7)
    #     df.loc[(df['Срок устранения до'].dt.date >= earliest_date.date()) &
    #            (df['Срок устранения до'].dt.date <= end_of_last_week_mon.date()) & (df['Просрок'] == 'Да') &
    #            (df['Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком с " + earliest_date.strftime(
    #         "%d.%m.%y") + " по " + end_of_last_week_mon.strftime("%d.%m.%y") + " (Старые)"
    if weekday == 0:
        end_of_last_week = today - timedelta(days=1)
        end_of_last_week_mon = end_of_last_week - timedelta(days=7)
        df.loc[(df['Срок устранения до'].dt.date >= earliest_date.date()) &
               (df['Срок устранения до'].dt.date <= end_of_last_week_mon.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком с " + earliest_date.strftime(
            "%d.%m.%y") + " по " + end_of_last_week_mon.strftime("%d.%m.%y") + " (Старые)"
    else:
        end_of_last_week = today - timedelta(days=(weekday+1))
        end_of_last_week_mon = end_of_last_week - timedelta(days=7)
        df.loc[(df['Срок устранения до'].dt.date >= earliest_date.date()) &
               (df['Срок устранения до'].dt.date <= end_of_last_week_mon.date()) & (df['Просрок'] == 'Да') &
               (df[
                    'Статус в системе'] == 'В работе'), "ТипСПросроком"] = "В работе с просроком с " + earliest_date.strftime(
            "%d.%m.%y") + " по " + end_of_last_week_mon.strftime("%d.%m.%y") + " (Старые)"
def fifth_attribute(df):
    today = datetime.now()
    if today:
        df.loc[(df['Срок устранения до'].dt.date == today.date()) & (df['Просрок'] == 'Нет') &
               (df['Статус в системе'] == 'В работе'), "ТипБезПросрока"] = "Срок с " + pd.Timestamp(
            datetime.now()).strftime('%H:%M') + " " + today.strftime("%d.%m.%y") + " (Сегодня)"
def sixth_attribute(df):
    today = datetime.now()
    tommorow = today + timedelta(days=1)
    max_date = df[(df['Просрок'] == 'Нет') &
                  (df['Статус в системе'] == 'В работе')]['Срок устранения до'].max()
    if today:
        df.loc[((df['Срок устранения до'].dt.date >= tommorow.date()) & (
                df['Срок устранения до'].dt.date <= max_date.date()) & (df['Просрок'] == 'Нет') &
                (df['Статус в системе'] == 'В работе')) |
               ((df['Обещание устранения'].dt.date >= tommorow.date()) & (
                       df['Обещание устранения'].dt.date <= max_date.date()) & (df['Просрок'] == 'Нет') &
                (df['Статус в системе'] == 'В работе')), "ТипБезПросрока"] = "Срок с " + tommorow.strftime(
            "%d.%m.%y") + " по " + max_date.strftime("%d.%m.%y")
def snow_today(df):
    today = datetime.now()
    if today:
        df.loc[(df['Дата фиксации нарушения'].dt.date == today.date()) &
               ((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')), "ТипСнег"] = "Снег " + today.strftime("%d.%m.%y") + " (Сегодня)"
def snow_all_expect_today(df):
    today = datetime.now()
    tomorrow = today - timedelta(days=1)
    weekday = today.weekday()
    # Определяем начало и конец текущей недели
    start_of_week = today - timedelta(days=weekday)
    if weekday == 6:
        start_day = today - timedelta(days=6)
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег с " + start_day.strftime(
            "%d.%m.%y") + " по " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    # во monday в этом столбце ничего не будет, т.к. данный снег будет находиться в другом столбце (снег сегодня)
    elif weekday == 1:
        start_day = start_of_week + timedelta(days=(weekday - 1))
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 2:
        start_day = start_of_week + timedelta(days=(weekday - 2))
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег с " + start_day.strftime(
            "%d.%m.%y") + " по " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 3:
        start_day = start_of_week + timedelta(days=(weekday - 3))
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег с " + start_day.strftime(
            "%d.%m.%y") + " по " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 4:
        start_day = start_of_week + timedelta(days=(weekday - 4))
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег с " + start_day.strftime(
            "%d.%m.%y") + " по " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"

    elif weekday == 5:
        start_day = start_of_week + timedelta(days=(weekday - 5))
        end_day = tomorrow
        df.loc[((df['Проблема'] == 'Наличие снега, наледи') | (df['Проблема'] == 'Неочищенная кровля')) & (
                df['Дата фиксации нарушения'].dt.date >= start_day.date())
               & (df['Дата фиксации нарушения'].dt.date <= end_day.date()), "ТипСнег"] = "Снег с " + start_day.strftime(
            "%d.%m.%y") + " по " + tomorrow.strftime(
            "%d.%m.%y") + " (Текущая уб. неделя)"
def process_file_MM(filepath, timenow):
    df = pd.read_excel(filepath)
    # Список значений, которые должны присутствовать в столбце "Балансодержатель"
    wanted_values = [
        'ГБУ «Автомобильные дороги ЮВАО»',
        'ГБУ «Жилищник Выхино района Выхино-Жулебино»',
        'ГБУ «Жилищник Нижегородского района»',
        'ГБУ «Жилищник района Капотня»',
        'ГБУ «Жилищник района Кузьминки»',
        'ГБУ «Жилищник района Лефортово»',
        'ГБУ «Жилищник района Люблино»',
        'ГБУ «Жилищник района Марьино»',
        'ГБУ «Жилищник района Некрасовка»',
        'ГБУ «Жилищник района Печатники»',
        'ГБУ «Жилищник района Текстильщики»',
        'ГБУ «Жилищник района Южнопортовый»',
        'ГБУ «Жилищник Рязанского района»'
    ]
    df = df[df['Балансодержатель'].isin(wanted_values)]

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
    df['Район'] = df['Ответственный исполнитель'].map(responsible_mapping)

    df['Срок устранения до'] = pd.to_datetime(df['Срок устранения до'])
    df['Обещание устранения'] = pd.to_datetime(df['Обещание устранения'])
    df['ТипБезПросрока'] = ''
    df['ТипСПросроком'] = ''
    df['ТипСнег'] = ''
    first_attribute(df)
    second_attribute(df)
    third_attribute(df)
    fourth_attribute(df)
    fifth_attribute(df)
    sixth_attribute(df)
    print(df[df["Проблема"] == "Наличие снега, наледи"])
    if not df[df["Проблема"].isin(["Наличие снега, наледи", "Неочищенная кровля"])].empty:
        print("Есть снег")
        snow_today(df)
        snow_all_expect_today(df)
    processed_file_path = os.path.join(directory,
                                       f"Монитор в работе_{timenow}_{datetime.now().strftime('%d.%m.%y')}.xlsx")
    df.to_excel(processed_file_path, index=False)
    with pd.ExcelWriter(processed_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='СВОД', index=False, startrow=0)
    excel_file = processed_file_path
    # VBA код макроса, который будет добавлен в Excel
    vba_macro = """  
Sub CreatePivotTable1()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim foundTodayColumn As Boolean
    Dim cell As Range

    ' Укажите лист с данными
    Set wsData = ThisWorkbook.Sheets("СВОД") ' Замените на имя вашего листа с данными

    ' Создаем новый лист для сводной таблицы
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Сводная таблица").Delete ' Удаляем лист, если уже существует
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "Сводная таблица"

    ' Находим последний заполненный ряд и столбец на листе с данными
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    ' Создаем кэш для сводной таблицы
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Cells(1, 1).Resize(lastRow, lastCol))

    ' Создаем сводную таблицу
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Cells(3, 1), _
        TableName:="MyPivotTable")

    ' Настройка сводной таблицы: строки - "Район", столбцы - "ТипБезПросрока", значения - количество строк
    With pivotTable
        .PivotFields("Район").Orientation = xlRowField
        .PivotFields("ТипБезПросрока").Orientation = xlColumnField
        .AddDataField .PivotFields("ID нарушения"), "Количество", xlCount
    End With
    With pivotTable
        .GrandTotalName = "На устранении без просрока" ' Замените на нужное название для общего итога
    End With
    wsPivot.Range("A4").Value = "Район"
    ' Скрываем первую строку
    wsPivot.Rows(3).Hidden = True

    ' Убираем столбец "Пусто" (где ТипБезПросрока неопределен)
    Dim typePivotField As PivotField
    Set typePivotField = pivotTable.PivotFields("ТипБезПросрока")
    For Each item In typePivotField.PivotItems
        If item.Name = "(blank)" Then
            item.Visible = False
        End If
    Next item

    ' Обновляем сводную таблицу
    pivotTable.RefreshTable

    ' Форматирование сводной таблицы
    Dim rng As Range
    Set rng = wsPivot.Range("A4").CurrentRegion
    With rng
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .WrapText = True ' Перенос текста
        .HorizontalAlignment = xlCenter ' Выравнивание по центру
        .VerticalAlignment = xlCenter
    End With
    wsPivot.Columns("A").ColumnWidth = 24 ' Установите желаемую ширину столбца
    With rng
        .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
        .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
    End With

    ' Настройка высоты строк и ширины столбцов
    wsPivot.Range("6:16").RowHeight = 19
    wsPivot.Columns("B").ColumnWidth = 39
    wsPivot.Columns("C").ColumnWidth = 34
    wsPivot.Columns("D").ColumnWidth = 33
    wsPivot.Columns("E").ColumnWidth = 39

    ' Проверяем наличие столбца с названием, содержащим "Сегодня"
    foundTodayColumn = False
    For Each cell In wsPivot.Range("B4:E4")
        If InStr(1, cell.Value, "Сегодня", vbTextCompare) > 0 Then
            foundTodayColumn = True 
            cell.Font.Color = RGB(255, 0, 0) ' Красный цвет текста заголовка
            Dim dataRange As Range
            Dim lastDataRow As Long
            lastDataRow = wsPivot.Cells(wsPivot.Rows.Count, cell.Column).End(xlUp).Row - 1 ' Уменьшаем на одну строку для исключения итогов
            Set dataRange = wsPivot.Range(cell.Offset(1, 0), wsPivot.Cells(lastDataRow, cell.Column))
            ' Применяем заливку к значениям > 0, исключая итоги
            For Each dataCell In dataRange
                If IsNumeric(dataCell.Value) And dataCell.Value > 0 Then
                    dataCell.Interior.Color = RGB(247, 134, 126) ' Красная заливка для положительных значений
                End If
            Next dataCell
        End If
    Next cell
End Sub
"""
    vba_macro2 = """Sub CreatePivotTable2()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pivotCache As PivotCache
    Dim pivotTable As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pivotStartRow As Long

    ' Укажите лист с данными
    Set wsData = ThisWorkbook.Sheets("СВОД") ' Замените на имя вашего листа с данными

    ' Укажите существующий лист для сводной таблицы
    Set wsPivot = ThisWorkbook.Sheets("Сводная таблица") ' Замените на имя вашего листа с существующей сводной таблицей

    ' Находим последний заполненный ряд и столбец на листе с данными
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    ' Находим строку, где уже существует сводная таблица, и добавляем новую через 3 строки
    pivotStartRow = wsPivot.Cells(wsPivot.Rows.Count, 1).End(xlUp).Row + 3

    ' Создаем кэш для сводной таблицы
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Cells(1, 1).Resize(lastRow, lastCol))

    ' Создаем сводную таблицу
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=wsPivot.Cells(pivotStartRow, 1), _
        TableName:="MyPivotTableWithExpiration")

    ' Настройка сводной таблицы: строки - "Район", столбцы - "ТипСПросроком", значения - количество строк
    With pivotTable
        .PivotFields("Район").Orientation = xlRowField
        .PivotFields("ТипСПросроком").Orientation = xlColumnField
        .AddDataField .PivotFields("ID нарушения"), "Количество", xlCount
    End With
        With pivotTable
        .GrandTotalName = "Сумма по просрочкам" ' Замените на нужное название для общего итога
    End With
    wsPivot.Range(wsPivot.Cells(pivotStartRow + 1, 1), wsPivot.Cells(pivotStartRow + 1, 1)).Value = "Район" ' Устанавливаем заголовок
    wsPivot.Rows(pivotStartRow).Hidden = True ' Скрываем строку с заголовками сводной таблицы

    ' Убираем столбец "Пусто" (где ТипСПросроком неопределен)
    Dim typePivotField As PivotField
    Set typePivotField = pivotTable.PivotFields("ТипСПросроком")
    For Each item In typePivotField.PivotItems
        If item.Name = "(blank)" Then
            item.Visible = False
        End If
    Next item
    ' Обновляем сводную таблицу
    pivotTable.RefreshTable

    ' Форматирование сводной таблицы
    Dim rng As Range
    Set rng = wsPivot.Range(wsPivot.Cells(pivotStartRow + 1, 1), wsPivot.Cells(pivotStartRow + 1, 1)).CurrentRegion
    With rng
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .WrapText = True ' Перенос текста
        .HorizontalAlignment = xlCenter ' Выравнивание по центру
        .VerticalAlignment = xlCenter
    End With
    wsPivot.Columns("A").ColumnWidth = 24 ' Установите желаемую ширину столбца
    wsPivot.Rows(pivotStartRow + 1).RowHeight = 53
    wsPivot.Rows(pivotStartRow + 3).RowHeight = 19 ' Установите высоту строки
    wsPivot.Columns("B").ColumnWidth = 39
    wsPivot.Columns("C").ColumnWidth = 34
    wsPivot.Columns("D").ColumnWidth = 33 
    wsPivot.Columns("E").ColumnWidth = 39 

    ' Изменение цвета текста в столбцах, содержащих заданные словосочетания
    Dim col As Integer
    Dim cell As Range
    Dim found As Boolean
    Dim searchStrings As Variant
    searchStrings = Array("В работе с просроком") ' Массив искомых словосочетаний

    For col = 1 To rng.Columns.Count
        found = False
        For Each cell In rng.Columns(col).Cells
            ' Проверяем только строки со 2-й по последнюю (исключая заголовок и итоговые строки)
            If cell.Row > pivotStartRow And cell.Row < rng.Rows.Count + pivotStartRow Then
                If Not IsEmpty(cell.Value) Then
                    For Each searchString In searchStrings
                        If InStr(1, cell.Value, searchString, vbTextCompare) > 0 Then
                            found = True
                            Exit For
                        End If
                    Next searchString
                End If
            End If
            If found Then Exit For
        Next cell

        ' Если найдено, изменить цвет текста в столбце
        If found Then
            ' Изменяем цвет текста только для значений, начиная со 2-й строки
            For Each cell In rng.Columns(col).Cells
                If cell.Row > pivotStartRow + 1 And cell.Row < rng.Rows.Count + pivotStartRow -1 Then
                    cell.Font.Color = RGB(255, 0, 0) ' Красный цвет
                End If
            Next cell
        End If
    Next col
End Sub
"""
    vba_macro_snow = """Sub CreatePivotTableSnow()
        Dim wsData As Worksheet
        Dim wsPivot As Worksheet
        Dim pivotCache As PivotCache
        Dim pivotTable As PivotTable
        Dim lastRow As Long
        Dim lastCol As Long
        Dim pivotStartRow As Long

        ' Укажите лист с данными
        Set wsData = ThisWorkbook.Sheets("СВОД") ' Замените на имя вашего листа с данными

        ' Укажите существующий лист для сводной таблицы
        Set wsPivot = ThisWorkbook.Sheets("Сводная таблица") ' Замените на имя вашего листа с существующей сводной таблицей

        ' Находим последний заполненный ряд и столбец на листе с данными
        lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
        lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

        ' Находим строку, где уже существует сводная таблица, и добавляем новую через 3 строки
        pivotStartRow = wsPivot.Cells(wsPivot.Rows.Count, 1).End(xlUp).Row + 3

        ' Создаем кэш для сводной таблицы
        Set pivotCache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=wsData.Cells(1, 1).Resize(lastRow, lastCol))

        ' Создаем сводную таблицу
        Set pivotTable = pivotCache.CreatePivotTable( _
            TableDestination:=wsPivot.Cells(pivotStartRow, 1), _
            TableName:="Pivotsnow")

        ' Настройка сводной таблицы: строки - "Район", столбцы - "ТипСнег", значения - количество строк
        With pivotTable
            .PivotFields("Район").Orientation = xlRowField
            .PivotFields("ТипСнег").Orientation = xlColumnField
            .AddDataField .PivotFields("ID нарушения"), "Количество", xlCount
        End With
            With pivotTable
            .GrandTotalName = "Сумма по снегу" ' Замените на нужное название для общего итога
        End With
        wsPivot.Range(wsPivot.Cells(pivotStartRow + 1, 1), wsPivot.Cells(pivotStartRow + 1, 1)).Value = "Район" ' Устанавливаем заголовок
        wsPivot.Rows(pivotStartRow).Hidden = True ' Скрываем строку с заголовками сводной таблицы

        ' Убираем столбец "Пусто" (где ТипСнег неопределен)
        Dim typePivotField As PivotField
        Set typePivotField = pivotTable.PivotFields("ТипСнег")
        For Each item In typePivotField.PivotItems
            If item.Name = "(blank)" Then
                item.Visible = False
            End If
        Next item
        ' Обновляем сводную таблицу
        pivotTable.RefreshTable

        ' Форматирование сводной таблицы
        Dim rng As Range
        Set rng = wsPivot.Range("A39").CurrentRegion
        With rng
            .Font.Name = "Times New Roman"
            .Font.Size = 14
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .WrapText = True ' Перенос текста
            .HorizontalAlignment = xlCenter ' Выравнивание по центру
            .VerticalAlignment = xlCenter
        End With
        wsPivot.Columns("A").ColumnWidth = 24 ' Установите желаемую ширину столбца
        With rng
            .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
            .VerticalAlignment = xlCenter ' Вертикальное выравнивание по центру
        End With

        ' Настройка высоты строк и ширины столбцов
        wsPivot.Range("40:52").RowHeight = 19
        wsPivot.Columns("B").ColumnWidth = 39
        wsPivot.Columns("C").ColumnWidth = 34
        wsPivot.Columns("D").ColumnWidth = 33
        wsPivot.Columns("E").ColumnWidth = 39

        ' Проверяем наличие столбца с названием, содержащим "снег"
        foundTodayColumn = False
        For Each cell In wsPivot.Range("B37:C39")
            If InStr(1, cell.Value, "Сегодня", vbTextCompare) > 0 Then
                foundTodayColumn = True 
                cell.Font.Color = RGB(255, 0, 0) ' Красный цвет текста заголовка
                Dim dataRange As Range
                Dim lastDataRow As Long
                lastDataRow = wsPivot.Cells(wsPivot.Rows.Count, cell.Column).End(xlUp).Row - 1 ' Уменьшаем на одну строку для исключения итогов
                Set dataRange = wsPivot.Range(cell.Offset(1, 0), wsPivot.Cells(lastDataRow, cell.Column))
                ' Применяем заливку к значениям > 0, исключая итоги
                For Each dataCell In dataRange
                    If IsNumeric(dataCell.Value) And dataCell.Value > 0 Then
                        dataCell.Interior.Color = RGB(247, 134, 126) ' Красная заливка для положительных значений
                    End If
                Next dataCell
            End If
        Next cell
    End Sub
    """

    # Запускаем Excel
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True  # Если нужно, чтобы Excel не отображался, оставьте False

    # Открываем Excel-файл
    workbook = excel.Workbooks.Open(excel_file)

    # Добавляем новый модуль VBA и вставляем макрос
    vb_module = workbook.VBProject.VBComponents.Add(1)  # 1 = стандартный модуль
    vb_module.CodeModule.AddFromString(vba_macro)
    # Выполняем макрос
    excel.Application.Run("CreatePivotTable1")
    print("Pivot1 created")

    vb_module1 = workbook.VBProject.VBComponents.Add(1)  # 1 = стандартный модуль
    vb_module1.CodeModule.AddFromString(vba_macro2)
    excel.Application.Run("CreatePivotTable2")
    print("Pivot2 created")

    if not df[df["Проблема"].isin(["Наличие снега, наледи", "Неочищенная кровля"])].empty:
        vb_module2 = workbook.VBProject.VBComponents.Add(1)  # 1 = стандартный модуль
        vb_module2.CodeModule.AddFromString(vba_macro_snow)
        excel.Application.Run("CreatePivotTableSnow")
        print("CreatePivotTableSnow")

    pdf_file_name = f"Монитор_в_работе_{timenow}_{datetime.now().strftime('%d.%m.%y')}.pdf"
    pdf_path = os.path.join(os.path.dirname(processed_file_path), pdf_file_name)  # Формируем путь к PDF
    wsFirst = workbook.Worksheets(1)  # Ссылка на первый лист

    # Настройки страницы для печати
    wsFirst.PageSetup.FitToPagesWide = 1  # Устанавливаем количество страниц по ширине
    wsFirst.PageSetup.FitToPagesTall = 1  # Устанавливаем количество страниц по высоте на 1
    wsFirst.PageSetup.Zoom = False  # Отключаем масштабирование

    # Обновляем отступы страницы для уменьшения размера PDF
    wsFirst.PageSetup.LeftMargin = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.RightMargin = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.TopMargin = excel.Application.CentimetersToPoints(0.5)
    wsFirst.PageSetup.BottomMargin = excel.Application.CentimetersToPoints(0.5)
    workbook.Save()
    try:
        # Убираем ошибку, если файл уже существует
        if os.path.exists(pdf_path):
            print(f"Файл {pdf_path} существует. Удаление...")
            os.remove(pdf_path)  # Удаляем файл, если он существует
            print("Файл успешно удален.")

        print(f"Сохранение файла в {pdf_path}...")
        wsFirst.ExportAsFixedFormat(0, pdf_path)  # 0 - это xlTypePDF
        print(f"PDF успешно создан: {pdf_path}")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")
    sheet = workbook.Worksheets(2)
    sheet.Cells.EntireColumn.AutoFit()
    # Сохраняем и закрываем файл
    workbook.Save()
    workbook.Close()

    # Закрываем Excel
    excel.Quit()
    return processed_file_path, pdf_path
