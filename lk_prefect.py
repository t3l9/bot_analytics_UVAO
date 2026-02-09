from selenium.webdriver import Keys
from telegram import Update, ReplyKeyboardMarkup, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes, \
    ConversationHandler, CallbackQueryHandler, CallbackContext
import time
import pandas as pd
import os
from datetime import datetime, timedelta
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.formatting.rule import CellIsRule
import win32com.client
from functools import reduce
import pythoncom
from dotenv import load_dotenv

home_dir = os.path.expanduser("~")
directory = os.path.join(home_dir, "Downloads")
# Загружаем переменные из .env
load_dotenv()
login_NG = os.getenv("login_NG")
password_NG = os.getenv("password_NG")


# ЛК Префекта -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
async def call_lk_prefekta(update: Update, chat_id, context: CallbackContext, district: str) -> None:
    success = await parcing_data_lk_prefekta(context, chat_id)  # Передаем контекст и ID чата
    if not success:
        return  # Если произошла ошибка, выходим из обработчика
    files = os.listdir(directory)
    files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
    latest_downloaded_file = files[-1]
    filepath = os.path.join(directory, latest_downloaded_file)
    processed_file_path = process_lk_prefekta_file(directory, district, filepath)
    if not processed_file_path:
        error_message = f"❌ Заявок ЛК Префекта по данному району нет!"
        print(error_message)  # Выводим ошибку в консоль
        await context.bot.send_message(chat_id=chat_id, text=error_message)  # Отправляем сообщение в Telegram
        return  # выходим из обработки
    else:
        with open(processed_file_path, 'rb') as f:
            await update.callback_query.message.reply_document(InputFile(f))


async def parcing_data_lk_prefekta(context, chat_id):
    chrome_install = ChromeDriverManager().install()
    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")
    driver = webdriver.Chrome(service=ChromeService(chromedriver_path))
    try:
        # Откройте страницу логина
        driver.get('https://gorod.mos.ru/api/service/auth/auth')

        # Найдите поля для ввода логина и пароля и заполните их
        username = driver.find_element(By.XPATH, '//input[@placeholder="Логин *"]')
        password = driver.find_element(By.XPATH, '//input[@placeholder="Пароль*"]')
        username.send_keys(login_NG)
        password.send_keys(password_NG)

        # Найдите и нажмите кнопку логина
        login_button = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/main/div/div/div/div[2]/form[1]/button')
        login_button.click()
        # Подождите, пока страница загрузится
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,
                                                                        '//div[@class="dashboard__block-link"]//div[@class="button-big link"]//div[@class="dashboard-container__links-title" and contains(text(), "Аналитика")]')))
        # переход в ответы в работе
        driver.get('https://gorod.mos.ru/admin/ker/olap/report/155')
        time.sleep(10)
        # # прыжок в меню
        # button = driver.find_element(By.XPATH,
        #                              "/html/body/div[3]/div/div[2]/div/div/div/div/form/header/div[1]/button[1]/span[2]/i")
        # button.click()
        # time.sleep(4)
        # # выбор фильтра
        # WebDriverWait(driver, 20).until(EC.presence_of_element_located(
        #     (By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div/div/form/div[1]/aside/div/div[2]/div/div[1]/div/a')))
        # button = driver.find_element(By.XPATH,
        #                              "/html/body/div[3]/div/div[2]/div/div/div/div/form/div[1]/aside/div/div[2]/div/div[1]/div/a")
        # button.click()

        # экспорт
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[3]/div/div[2]/div/div/div/div/form/footer/button[3]/span[2]/span')))
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[3]/div/div[2]/div/div/div/div/form/footer/button[3]/span[2]/span')
        button.click()
        time.sleep(1)

        # # ок- выгркзка с экселя
        # button = driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/button[2]/span[2]/span')
        # button.click()
        # time.sleep(1)

        # one more time click to export
        button = driver.find_element(By.XPATH, "//button[contains(@class, 'bg-primary')]//span[text()='Экспорт']")
        button.click()
        time.sleep(1)

        # переход в загрузки
        driver.get('https://gorod.mos.ru/admin/ker/olap/downloads')
        # Подождите, пока страница загрузится)
        WebDriverWait(driver, 1500).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]/div[1]/table/tbody/tr[1]/td[5]/div/i')))
        # скачивание файла
        button = driver.find_element(By.XPATH,
                                     '/html/body/div[1]/div/div[2]/main/div/div[1]/div/div[2]/div[1]/table/tbody/tr[1]/td[5]/div/i')
        button.click()
        time.sleep(20)
        return True
    except Exception as e:
        error_message = f"❌Произошла ошибка при выгрузке ЛК префекта. Пожалуйста, попробуйте еще раз."
        print(error_message)  # Выводим ошибку в консоль
        await context.bot.send_message(chat_id=chat_id, text=error_message)  # Отправляем сообщение в Telegram
        return False
    finally:
        driver.quit()


def process_lk_prefekta_file(directory: str, selected_district: str, filepath: str) -> str:
    df = pd.read_excel(filepath)

    responsible_mapping = {
        'ГБУ «Автомобильные дороги ЮВАО»': 'АВД ЮВАО',
        'ГБУ Жилищник Выхино района Выхино-Жулебино города Москвы': 'Выхино-Жулебино',
        'Управа Выхино-Жулебино': 'Выхино-Жулебино',
        'ГБУ Жилищник Нижегородского района города Москвы': 'Нижегородский',
        'Управа Нижегородский': 'Нижегородский',
        'ГБУ Жилищник района Капотня города Москвы': 'Капотня',
        'Управа Капотня': 'Капотня',
        'ГБУ Жилищник района Кузьминки города Москвы': 'Кузьминки',
        'Управа Кузьминки': 'Кузьминки',
        'ГБУ Жилищник района Лефортово города Москвы': 'Лефортово',
        'Управа Лефортово': 'Лефортово',
        'ГБУ Жилищник района Люблино города Москвы': 'Люблино',
        'Управа Люблино': 'Люблино',
        'ГБУ Жилищник района Марьино города Москвы': 'Марьино',
        'Управа Марьино': 'Марьино',
        'ГБУ Жилищник района Некрасовка города Москвы': 'Некрасовка',
        'Управа Некрасовка': 'Некрасовка',
        'ГБУ Жилищник района Печатники города Москвы': 'Печатники',
        'Управа Печатники': 'Печатники',
        'ГБУ Жилищник района Текстильщики города Москвы': 'Текстильщики',
        'Управа Текстильщики': 'Текстильщики',
        'ГБУ Жилищник Рязанского района города Москвы': 'Рязанский',
        'Управа Рязанский': 'Рязанский',
        'ГБУ Жилищник Южнопортового района города Москвы': 'Южнопортовый',
        'Управа Южнопортовый': 'Южнопортовый'
    }

    # Функция для обновления значений в столбце 'Район'
    def update_region(row):
        if row['Ответственный ОИВ первого уровня'] == 'Префектура Юго-Восточного округа':
            return row['Район']  # Ничего не меняем
        else:
            return responsible_mapping.get(row['Ответственный ОИВ первого уровня'], row['Район'])

    # Применение функции к каждому ряду
    df['Район'] = df.apply(update_region, axis=1)

    df_filtered = df[df['Ответственный за подготовку ответа'] == 'Префектура Юго-Восточного округа']

    columns_to_keep = [
        "Номер заявки",
        "Регламентный срок у сообщения (Портал)",
        "Дата публикации сообщения",
        "Район",
        "Проблемная тема",
        "Адрес",
        "Категория объекта",
        "Категория/действие последнего ответа",
        "Ответственный за подготовку ответа",
        "Ответственный ОИВ первого уровня",
        "Статус подготовки ответа на сообщение"
    ]
    df_filtered = df_filtered[columns_to_keep]
    if selected_district != "Все районы":
        df_filtered = df_filtered[df_filtered['Район'] == selected_district]

    # Удаляем полностью пустые строки и проверяем количество оставшихся строк
    df_filtered = df_filtered.dropna(how='all')
    if df_filtered.empty:
        return False

    now = pd.Timestamp.now()
    processed_file_path = os.path.join(directory,
                                       f"{selected_district}_ЛК_Префекта_{datetime.now().strftime('%d.%m')}_на_{now.strftime('%H-%M')}.xlsx")
    print(f"Saving processed file to: {processed_file_path}")
    df_filtered.to_excel(processed_file_path, index=False)
    excel_file = processed_file_path
    vba_macro = """  
            Sub CreatePivotTable()  
                Dim wsData As Worksheet  
                Dim wsPivot As Worksheet  
                Dim pivotCache As PivotCache  
                Dim pivotTable As PivotTable  
                Dim lastRow As Long  
                Dim lastCol As Long  

                ' Укажите лист с данными  
                Set wsData = ThisWorkbook.Sheets("Sheet1") ' Замените на имя вашего листа с данными  

                With wsData.Columns("B")
                    .NumberFormat = "DD.MM.YYYY"
                End With

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

                With pivotTable  
                    .PivotFields("Район").Orientation = xlRowField  
                    .PivotFields("Регламентный срок у сообщения (Портал)").Orientation = xlColumnField  
                    .AddDataField .PivotFields("Номер заявки"), "Количество", xlCount  
                End With  

                wsPivot.Range("A4").Value = "Район" 
                ' Скрываем первую строку  
                wsPivot.Rows(3).Hidden = True  

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
                End With  
                wsPivot.Columns("A").ColumnWidth = 24 ' Установите желаемую ширину столбца  
                wsPivot.Rows(6).RowHeight = 19 ' Установите высоту 6-й строки

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
    excel.Application.Run("CreatePivotTable")
    print("Pivot created")
    pdf_file_name = f"{selected_district}_ЛК_Префекта_{datetime.now().strftime('%d.%m')}_на_{now.strftime('%H-%M')}.xlsx"
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
    return processed_file_path
