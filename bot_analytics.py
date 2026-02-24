import os
import asyncio
import logging
import pandas as pd
from datetime import datetime, timedelta
from functools import reduce
import re
import tempfile
import shutil

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes, MessageHandler, filters
from telegram.constants import ParseMode
from dotenv import load_dotenv

import logging

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
    print("Модуль win32com успешно импортирован")
except ImportError as e:
    WIN32COM_AVAILABLE = False
    print(f"Модуль win32com не установлен: {e}. Создание сводной таблицы и PDF будет пропущено.")

from oati import (
    get_week_dates_OATI, create_ppt_OATI, process_file_OATI
)

from week_svod import (
    parcing_data_MM_async, process_file_MM_week
)

from mji_svod import (
    parcing_MWI, MWI_choosing_files, MWI_process_file, create_pivot_and_pdf
)

from mmonitor import (
    parcing_data_MM, choosing_time_MM, process_file_MM
)

from ng_otvety import (
    choosing_time_NG, process_ng_prosroki_file, parcing_data,
    personalizating_table_osn, personalizating_table_prosrok,
    personalizating_table_eight_day, personalizating_table_seven_day,
    personalizating_table_six_day, personalizating_table_five_day,
    add_run_delete_and_save_files
)

from lk_prefect import (
    call_lk_prefekta, process_lk_prefekta_file, parcing_data_lk_prefekta
)

# Загружаем переменные из .env
load_dotenv()
TOKEN = os.getenv("TOKEN")
# Получаем домашнюю директорию пользователя
home_dir = os.path.expanduser("~")
# Путь к папке загрузок
directory = os.path.join(home_dir, "Downloads")
excluded_dates = [
    # Ваши исходные даты 2025 года
    "20.12.2025", "21.12.2025", "27.12.2025", "28.12.2025", "31.12.2025",

    # Новогодние каникулы и Рождество (2026)
    "01.01.2026", "02.01.2026", "03.01.2026", "04.01.2026", "05.01.2026",
    "06.01.2026", "07.01.2026", "08.01.2026", "09.01.2026", "10.01.2026",
    "11.01.2026",

    # Февральские праздники (23 февраля выпадает на понедельник, добавляем только его)
    "23.02.2026",
    # Предшествующие выходные
    "21.02.2026", "22.02.2026",

    # Мартовские праздники (8 марта - воскресенье, выходной переносится на 9 марта)
    "07.03.2026", "08.03.2026", "09.03.2026",

    # Майские праздники (1-3 мая и 9-11 мая)
    "01.05.2026", "02.05.2026", "03.05.2026",
    "09.05.2026", "10.05.2026", "11.05.2026",

    # День России (12 июня - пятница, длинные выходные 13-14 июня)
    "12.06.2026", "13.06.2026", "14.06.2026",

    # Ноябрьские праздники (4 ноября - среда, отдельный выходной)
    "04.11.2026",
    # Ближайшие выходные
    "31.10.2026", "01.11.2026", "07.11.2026", "08.11.2026",

    # Стандартные выходные 2026 года (субботы и воскресенья, не попавшие в периоды выше)
    # Январь
    "17.01.2026", "18.01.2026", "24.01.2026", "25.01.2026", "31.01.2026",
    # Февраль
    "01.02.2026", "07.02.2026", "08.02.2026", "14.02.2026", "15.02.2026", "28.02.2026",
    # Март
    "01.03.2026", "14.03.2026", "15.03.2026", "21.03.2026", "22.03.2026", "28.03.2026", "29.03.2026",
    # Апрель
    "04.04.2026", "05.04.2026", "11.04.2026", "12.04.2026", "18.04.2026", "19.04.2026", "25.04.2026", "26.04.2026",
    # Май (добавлены только неохваченные)
    "16.05.2026", "17.05.2026", "23.05.2026", "24.05.2026", "30.05.2026", "31.05.2026",
    # Июнь (добавлены только неохваченные)
    "06.06.2026", "07.06.2026", "20.06.2026", "21.06.2026", "27.06.2026", "28.06.2026",
    # Июль
    "04.07.2026", "05.07.2026", "11.07.2026", "12.07.2026", "18.07.2026", "19.07.2026", "25.07.2026", "26.07.2026",
    # Август
    "01.08.2026", "02.08.2026", "08.08.2026", "09.08.2026", "15.08.2026", "16.08.2026", "22.08.2026", "23.08.2026",
    "29.08.2026", "30.08.2026",
    # Сентябрь
    "05.09.2026", "06.09.2026", "12.09.2026", "13.09.2026", "19.09.2026", "20.09.2026", "26.09.2026", "27.09.2026",
    # Октябрь (добавлены только неохваченные)
    "10.10.2026", "11.10.2026", "17.10.2026", "18.10.2026", "24.10.2026", "25.10.2026",
    # Ноябрь (добавлены только неохваченные)
    "14.11.2026", "15.11.2026", "21.11.2026", "22.11.2026", "28.11.2026", "29.11.2026",
    # Декабрь
    "05.12.2026", "06.12.2026", "12.12.2026", "13.12.2026", "19.12.2026", "20.12.2026", "26.12.2026", "27.12.2026",
    # Канун Нового 2027 года
    "31.12.2026"
]

# Логирование
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Текстовые константы
EXPLANATION_TEXT = """
📋 *Объяснение команд:*

*🏢 ЛК префекта (НГ)*
Отчет по всем заявкам ЛК Префекта (все районы)

*📊 Монитор в Работе (ММ)*
Сводный отчет по Монитору ММ за выбранный период

*📈 Ответы в работе (НГ)*
Отчет "Ответы в работе" с просрочками по дням

*📄 Cвод МЖИ (НГ)*
Отчет по МЖИ с актуальными данными по заявкам МЖИ, которые сейчас в работе

*📎 Еженедельный свод*
Еженедельный свод с Монитора Мэра для презентаций

*🅾️ Слайд ОАТИ*
Создание слайда ОАТИ

*❓ Объяснение команд*
Это сообщение с объяснением всех команд
"""

# Создаем клавиатуру команд
MAIN_KEYBOARD = InlineKeyboardMarkup([
    [InlineKeyboardButton("🏢 ЛК префекта (НГ)", callback_data='lk_prefekt')],
    [InlineKeyboardButton("📊 Монитор в Работе (ММ)", callback_data='mm_monitor')],
    [InlineKeyboardButton("📈 Ответы в работе (НГ)", callback_data='ng_answers')],
    [InlineKeyboardButton("📄 Cвод МЖИ (НГ)", callback_data='mji_svod')],
    [InlineKeyboardButton("📎 Еженедельный свод", callback_data='week_svod')],
    [InlineKeyboardButton("🅾️ Слайд ОАТИ", callback_data='oati')],
    [InlineKeyboardButton("❓ Объяснение команд", callback_data='explain')],
])


def get_user_name(user):
    """Получает имя пользователя для отображения"""
    if user.username:
        return f"@{user.username}"
    elif user.first_name:
        return user.first_name
    else:
        return f"пользователь {user.id}"


async def delete_message_and_show_loading(query, context, loading_text="🔄 Загрузка данных..."):
    """Удаляет сообщение с кнопками и показывает анимацию загрузки с именем пользователя"""
    user_name = get_user_name(query.from_user)
    
    # Сначала удаляем сообщение с кнопками
    try:
        await query.message.delete()
    except Exception as e:
        logger.error(f"Ошибка при удалении сообщения: {e}")

    # Затем показываем анимацию загрузки
    loading_msg_id = await show_loading_animation(
        query.message.chat_id,
        context,
        f"{loading_text}\n👤 Запрос от {user_name}"
    )
    return loading_msg_id


# Функция для отправки анимированного сообщения "в процессе"
async def show_loading_animation(chat_id: int, context: ContextTypes.DEFAULT_TYPE,
                                 text: str = "🔄 Загрузка данных...") -> int:
    """Отправляет анимированное сообщение и возвращает его message_id"""
    message = await context.bot.send_message(
        chat_id=chat_id,
        text=text,
        parse_mode=ParseMode.HTML
    )
    return message.message_id


# Функция для обновления анимированного сообщения
async def update_loading_message(chat_id: int, message_id: int, context: ContextTypes.DEFAULT_TYPE,
                                 text: str):
    """Обновляет текст сообщения с анимацией"""
    try:
        await context.bot.edit_message_text(
            chat_id=chat_id,
            message_id=message_id,
            text=text,
            parse_mode=ParseMode.HTML
        )
    except Exception as e:
        logger.error(f"Ошибка при обновлении сообщения: {e}")


# Обработчик команды /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик команды /start"""
    user = update.effective_user
    user_name = get_user_name(user)
    
    welcome_text = f"""
👋 Привет, {user_name}!

Выберите команду из меню ниже:
    """

    # Если это ответ на сообщение, удаляем команду /start
    if update.message:
        await update.message.reply_text(
            welcome_text,
            reply_markup=MAIN_KEYBOARD,
            parse_mode=ParseMode.HTML
        )


# Обработчик кнопки "Объяснение команд"
async def explain_commands(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Объяснение всех команд"""
    query = update.callback_query
    await query.answer()

    # Удаляем старое сообщение
    try:
        await query.message.delete()
    except Exception as e:
        logger.error(f"Ошибка при удалении сообщения: {e}")

    # Отправляем новое сообщение с клавиатурой
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=EXPLANATION_TEXT,
        reply_markup=MAIN_KEYBOARD,
        parse_mode=ParseMode.MARKDOWN
    )


# Обработчик кнопки "ЛК префекта(НГ)"
async def lk_prefekt_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на отчет ЛК Префекта (все районы)"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    # Удаляем сообщение с кнопками и показываем загрузку
    loading_msg_id = await delete_message_and_show_loading(
        query,
        context,
        f"🏢 Загружаю отчет ЛК Префекта (все районы)..."
    )

    try:
        # Шаг 1: Выгрузка данных
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"🏢 Загружаю отчет ЛК Префекта (все районы)...\n👤 Запрос от {user_name}\n\n📥 Выполняется выгрузка данных..."
        )

        # Выполняем парсинг данных
        success = await parcing_data_lk_prefekta(context, query.message.chat_id)
        if not success:
            await context.bot.delete_message(
                chat_id=query.message.chat_id,
                message_id=loading_msg_id
            )
            # Показываем новую клавиатуру
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите команду:",
                reply_markup=MAIN_KEYBOARD
            )
            return

        # Шаг 2: Обработка файла
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"🏢 Загружаю отчет ЛК Префекта (все районы)...\n👤 Запрос от {user_name}\n\n⚙️ Обрабатываю данные..."
        )

        # Находим последний загруженный файл - С ИСПРАВЛЕНИЕМ!
        files = os.listdir(directory)

        # Отладочная информация
        print(f"Файлы в директории: {files}")

        if not files:
            raise Exception("Не найдены файлы в директории. Проверьте путь и права доступа.")

        # Фильтруем только .xlsx файлы
        excel_files = [f for f in files if f.endswith('.xlsx')]

        if not excel_files:
            raise Exception("Не найдены Excel файлы (.xlsx)")

        # Сортируем по времени изменения
        excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
        latest_downloaded_file = excel_files[-1]
        print(f"Последний файл: {latest_downloaded_file}")

        filepath = os.path.join(directory, latest_downloaded_file)

        # Обрабатываем файл для всех районов
        district = "Все районы"

        # Пробуем обработать файл
        try:
            processed_file_path = process_lk_prefekta_file(directory, district, filepath)
        except Exception as e:
            print(f"Ошибка при обработке файла: {e}")
            processed_file_path = None

        if not processed_file_path:
            await update_loading_message(
                query.message.chat_id,
                loading_msg_id,
                context,
                f"❌ Ошибка при обработке файла ЛК Префекта!\n👤 Запрос от {user_name}\n\nВозможно, файл пуст или поврежден."
            )
            await asyncio.sleep(3)
            await context.bot.delete_message(
                chat_id=query.message.chat_id,
                message_id=loading_msg_id
            )
            # Показываем новую клавиатуру
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите команду:",
                reply_markup=MAIN_KEYBOARD
            )
            return

        # Шаг 3: Отправка файла
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"🏢 Загружаю отчет ЛК Префекта (все районы)...\n👤 Запрос от {user_name}\n\n📤 Отправляю файл..."
        )

        # Проверяем существование файла перед отправкой
        if not os.path.exists(processed_file_path):
            raise Exception(f"Файл не найден: {processed_file_path}")

        # Отправляем файл
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M')
        with open(processed_file_path, 'rb') as f:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(f, filename=f"ЛК_Префекта_все_районы_{datetime.now().strftime('%d.%m_%H-%M')}.xlsx"),
                caption=f"🏢 Отчет ЛК Префекта (все районы) на {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )

        # Удаляем сообщение о загрузке
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем новую клавиатуру с именем пользователя
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"✅ Отчет успешно сформирован и отправлен для {user_name}!\n(время выгрузки: {current_time})\n\nВыберите следующую команду:",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception as e:
        logger.error(f"Ошибка при обработке ЛК Префекта: {e}")

        # Более подробное сообщение об ошибке
        error_details = f"""
❌ <b>Ошибка при обработке ЛК Префекта:</b>
👤 Запрос от {user_name}
<code>{str(e)}</code>

<b>Возможные причины:</b>
• Файл не был скачан
• Проблемы с доступом к порталу
• Неверный формат файла
• Пустой файл
        """

        # Обновляем сообщение об ошибке
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            error_details
        )

        # Ждем и удаляем сообщение об ошибке
        await asyncio.sleep(5)
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем новую клавиатуру
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="Выберите команду:",
            reply_markup=MAIN_KEYBOARD
        )


# Обработчик кнопки "Ответы в работе (НГ)"
async def ng_answers_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на отчет 'Ответы в работе'"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    # Удаляем сообщение с кнопками и показываем загрузку
    loading_msg_id = await delete_message_and_show_loading(
        query,
        context,
        f"📈 Загружаю отчет 'Ответы в работе'..."
    )

    try:
        # Шаг 1: Выгрузка данных
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📈 Загружаю отчет 'Ответы в работе'...\n👤 Запрос от {user_name}\n\n📥 Выполняется выгрузка данных с портала..."
        )

        success = await parcing_data(context, query.message.chat_id)
        if not success:
            await context.bot.delete_message(
                chat_id=query.message.chat_id,
                message_id=loading_msg_id
            )
            # Показываем новую клавиатуру
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите команду:",
                reply_markup=MAIN_KEYBOARD
            )
            return

        # Шаг 2: Обработка файла
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📈 Загружаю отчет 'Ответы в работе'...\n👤 Запрос от {user_name}\n\n⚙️ Обрабатываю данные..."
        )

        # Находим последний загруженный файл - С ИСПРАВЛЕНИЕМ!
        files = os.listdir(directory)

        # Отладочная информация
        print(f"Файлы в директории: {files}")

        if not files:
            raise Exception("Не найдены файлы в директории. Проверьте путь и права доступа.")

        # Фильтруем только .xlsx файлы
        excel_files = [f for f in files if f.endswith('.xlsx')]

        if not excel_files:
            raise Exception("Не найдены Excel файлы (.xlsx)")

        # Сортируем по времени изменения
        excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
        latest_downloaded_file = excel_files[-1]
        print(f"Последний файл: {latest_downloaded_file}")

        filepath = os.path.join(directory, latest_downloaded_file)

        # Получаем время
        timenow = choosing_time_NG()

        # Обрабатываем файл
        processed_file_path = process_ng_prosroki_file(timenow, filepath, excluded_dates)

        # Шаг 3: Форматирование таблиц
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📈 Загружаю отчет 'Ответы в работе'...\n👤 Запрос от {user_name}\n\n🎨 Применяю форматирование..."
        )

        # Применяем форматирование ко всем таблицам
        personalizating_table_osn(timenow)
        personalizating_table_prosrok(timenow)
        personalizating_table_eight_day(timenow)
        personalizating_table_seven_day(timenow)
        personalizating_table_six_day(timenow)
        personalizating_table_five_day(timenow)

        # Шаг 4: Создание PDF и финальных файлов
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📈 Загружаю отчет 'Ответы в работе'...\n👤 Запрос от {user_name}\n\n📄 Создаю финальные документы..."
        )

        pdf_path, first_sheet_file_path, full_file_path = add_run_delete_and_save_files(timenow)

        # Шаг 5: Отправка файлов
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M')
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📈 Загружаю отчет 'Ответы в работе'...\n👤 Запрос от {user_name}\n\n📤 Отправляю файлы..."
        )

        # 1. Отправляем PDF
        with open(pdf_path, 'rb') as pdf_file:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(pdf_file, filename=f"Ответы_в_работе_{datetime.now().strftime('%d.%m_%H-%M')}.pdf"),
                caption=f"📊 Отчет 'Ответы в работе' на {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )

        # 2. Отправляем Excel с одним листом (СВОД)
        with open(first_sheet_file_path, 'rb') as excel_file:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(excel_file,
                                   filename=f"СВОД_Ответы_в_работе_{datetime.now().strftime('%d.%m_%H-%M')}.xlsx"),
                caption=f"📋 Сводная таблица в Excel (выгрузка: {current_time})"
            )

        # 3. Отправляем ПОЛНЫЙ Excel со всеми листами
        if os.path.exists(full_file_path):
            with open(full_file_path, 'rb') as full_excel_file:
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=InputFile(full_excel_file,
                                       filename=f"Ответы_в_работе_{datetime.now().strftime('%d.%m_%H-%M')}.xlsx"),
                    caption=f"📁 Детальные данные по дням (выгрузка: {current_time})"
                )
        else:
            # Если полный файл не найден, ищем последний созданный Excel файл
            excel_files = [f for f in os.listdir(directory) if f.startswith('Ответы в работе_') and f.endswith('.xlsx')]
            if excel_files:
                excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
                latest_full_file = excel_files[-1]
                latest_full_path = os.path.join(directory, latest_full_file)

                with open(latest_full_path, 'rb') as full_excel_file:
                    await context.bot.send_document(
                        chat_id=query.message.chat_id,
                        document=InputFile(full_excel_file,
                                           filename=f"Ответы_в_работе_{datetime.now().strftime('%d.%m_%H-%M')}.xlsx"),
                        caption=f"📁 Детальные данные по дням (выгрузка: {current_time})"
                    )
            else:
                await context.bot.send_message(
                    chat_id=query.message.chat_id,
                    text="⚠️ Не удалось найти полный Excel файл со всеми листами."
                )

        # Удаляем сообщение о загрузке
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем новую клавиатуру с именем пользователя
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"✅ Отчет успешно сформирован и отправлен для {user_name}!\n(время выгрузки: {current_time})\n\nВыберите следующую команду:",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception as e:
        logger.error(f"Ошибка при обработке 'Ответы в работе': {e}")

        # Обновляем сообщение об ошибке
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"❌ <b>Ошибка при обработке 'Ответы в работе':</b>\n👤 Запрос от {user_name}\n<code>{str(e)}</code>"
        )

        # Ждем и удаляем сообщение об ошибке
        await asyncio.sleep(5)
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем новую клавиатуру
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="Выберите команду:",
            reply_markup=MAIN_KEYBOARD
        )


# Обработчик кнопки "Монитор в Работе (ММ)"
def choosing_time_frame_MM():
    today = datetime.now()
    weekday = today.weekday()
    start_of_week = today - timedelta(days=weekday)
    end_of_week = start_of_week + timedelta(days=6)
    if weekday == 0:
        start_day = today - timedelta(days=1)  # на один день назад
        end_day = today
    elif weekday == 1:
        start_day = start_of_week + timedelta(days=(weekday - 2))
        end_day = today
    elif weekday == 2:
        start_day = start_of_week + timedelta(days=(weekday - 3))
        end_day = today
    elif weekday == 3:
        start_day = start_of_week + timedelta(days=(weekday - 4))
        end_day = today
    elif weekday == 4:
        start_day = start_of_week + timedelta(days=(weekday - 5))
        end_day = today
    elif weekday == 5:
        start_day = start_of_week + timedelta(days=(weekday - 6))
        end_day = today
    elif weekday == 6:
        start_day = start_of_week + timedelta(days=(weekday - 7))
        end_day = today
    start_date = start_day.strftime("%d%m%Y")
    start_date = start_date + "2100"
    end_date = end_day.strftime("%d%m%Y")
    end_date = end_date + "2100"
    return start_date, end_date


async def mm_monitor_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на отчет 'Монитор в работе'"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    # Сохраняем chat_id до удаления сообщения
    chat_id = query.message.chat_id

    # Удаляем сообщение с кнопками и показываем загрузку
    loading_msg_id = await delete_message_and_show_loading(
        query,
        context,
        f"📊 Загружаю отчет 'Монитор в работе'..."
    )

    try:
        # Определяем даты
        MM_start_date, MM_end_date = choosing_time_frame_MM()

        # Шаг 1: Выгрузка данных
        await update_loading_message(
            chat_id,  # Используем сохраненный chat_id
            loading_msg_id,
            context,
            f"📊 Загружаю отчет 'Монитор в работе'...\n👤 Запрос от {user_name}\n\n📥 Выполняется выгрузка данных..."
        )

        success = await parcing_data_MM(context, chat_id, MM_start_date, MM_end_date)
        if not success:
            await context.bot.delete_message(
                chat_id=chat_id,
                message_id=loading_msg_id
            )
            # Показываем новую клавиатуру (используем chat_id)
            await context.bot.send_message(
                chat_id=chat_id,
                text="Выберите команду:",
                reply_markup=MAIN_KEYBOARD
            )
            return

        # Шаг 2: Обработка файла
        await update_loading_message(
            chat_id,
            loading_msg_id,
            context,
            f"📊 Загружаю отчет 'Монитор в работе'...\n👤 Запрос от {user_name}\n\n⚙️ Обрабатываю данные..."
        )

        # Находим последний загруженный файл
        files = os.listdir(directory)
        files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
        latest_downloaded_file = files[-1]
        filepath = os.path.join(directory, latest_downloaded_file)

        # Получаем время
        timenow = choosing_time_MM()

        # Обрабатываем файл
        processed_file_path, pdf_path = process_file_MM(filepath, timenow)

        # Шаг 3: Отправка файлов
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M')
        await update_loading_message(
            chat_id,
            loading_msg_id,
            context,
            f"📊 Загружаю отчет 'Монитор в работе'...\n👤 Запрос от {user_name}\n\n📤 Отправляю файлы..."
        )

        # Отправляем PDF
        with open(pdf_path, 'rb') as pdf_file:
            await context.bot.send_document(
                chat_id=chat_id,
                document=InputFile(pdf_file,
                                 filename=f"Монитор_в_работе_{timenow}_{datetime.now().strftime('%d.%m.%y_%H-%M')}.pdf"),
                caption=f"📊 Отчет 'Монитор в работе' (выгрузка: {current_time})"
            )

        # Отправляем Excel
        with open(processed_file_path, 'rb') as excel_file:
            await context.bot.send_document(
                chat_id=chat_id,
                document=InputFile(excel_file,
                                 filename=f"Монитор_в_работе_{timenow}_{datetime.now().strftime('%d.%m.%y_%H-%M')}.xlsx"),
                caption=f"📋 Полный отчет в Excel (выгрузка: {current_time})"
            )

        # Удаляем сообщение о загрузке
        await context.bot.delete_message(
            chat_id=chat_id,
            message_id=loading_msg_id
        )

        # Показываем меню (используем bot.send_message вместо query.message.reply_text)
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"✅ Отчет успешно сформирован и отправлен для {user_name}!\n\nВыберите следующую команду:",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception as e:
        logger.error(f"Ошибка при обработке 'Монитор в работе': {e}")

        # Обновляем сообщение об ошибке
        await update_loading_message(
            chat_id,
            loading_msg_id,
            context,
            f"❌ <b>Ошибка при обработке 'Монитор в работе':</b>\n👤 Запрос от {user_name}\n<code>{str(e)}</code>"
        )

        # Ждем и показываем меню
        await asyncio.sleep(5)
        await context.bot.delete_message(
            chat_id=chat_id,
            message_id=loading_msg_id
        )

        # Используем bot.send_message вместо query.message.reply_text
        await context.bot.send_message(
            chat_id=chat_id,
            text="Выберите команду:",
            reply_markup=MAIN_KEYBOARD
        )


# Обработчик кнопки "Свод МЖИ (НГ)"
async def mji_svod_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на отчет 'Свод МЖИ (НГ)'"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    # Удаляем сообщение с кнопками и показываем загрузку
    loading_msg_id = await delete_message_and_show_loading(
        query,
        context,
        f"📄 Загружаю отчет 'Свод МЖИ'..."
    )

    try:
        # Шаг 1: Выгрузка данных
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📄 Загружаю отчет 'Свод МЖИ'...\n👤 Запрос от {user_name}\n\n📥 Выполняется выгрузка данных..."
        )

        processed_count = await parcing_MWI(context, query.message.chat_id)

        # Шаг 2: Обработка файла
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📄 Загружаю отчет 'Свод МЖИ'...\n👤 Запрос от {user_name}\n\n⚙️ Обрабатываю данные..."
        )

        # Получаем DataFrame
        df = MWI_process_file(MWI_choosing_files(directory, processed_count))
        today = datetime.now()
        timenow = today.strftime("%H-%M")

        # Сохраняем в Excel
        excel_file = os.path.join(directory, f"СВОД МЖИ {datetime.now().strftime('%d.%m.%y')} на {timenow}.xlsx")
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='МЖИ', index=False, startrow=0)

        # Шаг 3: Создание сводной таблицы и PDF
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📄 Загружаю отчет 'Свод МЖИ'...\n👤 Запрос от {user_name}\n\n📊 Создаю сводную таблицу и PDF..."
        )

        # Используем функцию из модуля
        pdf_path, success, message = create_pivot_and_pdf(excel_file, directory)

        if not success:
            logger.warning(f"PDF не создан: {message}")

        # Шаг 4: Отправка файлов
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M')
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"📄 Загружаю отчет 'Свод МЖИ'...\n👤 Запрос от {user_name}\n\n📤 Отправляю файлы..."
        )

        # Отправляем Excel файл
        with open(excel_file, 'rb') as f:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=InputFile(f, filename=f"СВОД МЖИ {datetime.now().strftime('%d.%m.%y')} на {timenow}.xlsx"),
                caption=f"📊 Отчет 'Свод МЖИ' (Excel) на {current_time}"
            )

        # Отправляем PDF, если он был создан
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as pdf_f:
                await context.bot.send_document(
                    chat_id=query.message.chat_id,
                    document=InputFile(pdf_f,
                                       filename=f"СВОД МЖИ {datetime.now().strftime('%d.%m.%y')} на {timenow}.pdf"),
                    caption=f"📄 Отчет 'Свод МЖИ' (PDF) на {current_time}"
                )
        else:
            # Информируем о проблеме с PDF
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=f"ℹ️ PDF не был создан: {message}\n\nУстановите pywin32 для полной функциональности."
            )

        # Удаляем сообщение о загрузке
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем меню с именем пользователя
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"✅ Отчет успешно сформирован и отправлен для {user_name}!\n(время выгрузки: {current_time})\n\nВыберите следующую команду:",
            reply_markup=MAIN_KEYBOARD
        )

    except Exception as e:
        logger.error(f"Ошибка при обработке 'Свод МЖИ': {e}")

        # Обновляем сообщение об ошибке
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"❌ <b>Ошибка при обработке 'Свод МЖИ':</b>\n👤 Запрос от {user_name}\n<code>{str(e)}</code>"
        )

        # Ждем и показываем меню
        await asyncio.sleep(5)
        await context.bot.delete_message(
            chat_id=query.message.chat_id,
            message_id=loading_msg_id
        )

        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="Выберите команду:",
            reply_markup=MAIN_KEYBOARD
        )


# Обработчик кнопки по еженедельному своду
async def week_svod_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на еженедельный свод"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    request_text = (
        f"👤 Запрос от {user_name}\n\n"
        "*Введите две даты через пробел в формате дд.мм.гггг:*\n"
        "(например, *01.01.2022* *31.01.2022*)\n\n"
        "_Пожалуйста, убедитесь, что даты введены корректно, чтобы избежать ошибок в обработке._"
    )

    try:
        await query.message.delete()
    except Exception as e:
        logger.error(f"Ошибка при удалении сообщения: {e}")

    context.user_data['waiting_for_dates'] = True
    context.user_data['callback_query'] = query

    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=request_text,
        parse_mode=ParseMode.MARKDOWN
    )


# Обработчик текста от пользователя (ввод дат для еженедельного свода)
async def handle_dates_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка ввода дат от пользователя"""
    if not context.user_data.get('waiting_for_dates', False):
        return

    user = update.effective_user
    user_name = get_user_name(user)
    user_message = update.message.text.strip()
    parts = user_message.split()

    DATE_REGEX = r'\d{2}\.\d{2}\.\d{4}'

    if len(parts) != 2:
        await update.message.reply_text(
            f'❌ {user_name}, пожалуйста, введите ровно две даты через пробел.',
            parse_mode=ParseMode.MARKDOWN
        )
        return

    date1, date2 = parts

    if not re.match(DATE_REGEX, date1) or not re.match(DATE_REGEX, date2):
        await update.message.reply_text(
            f'❌ {user_name}, неверный формат даты. Пожалуйста, используйте формат дд.мм.гггг.',
            parse_mode=ParseMode.MARKDOWN
        )
        return

    try:
        datetime.strptime(date1, '%d.%m.%Y')
        datetime.strptime(date2, '%d.%m.%Y')
    except ValueError:
        await update.message.reply_text(
            f'❌ {user_name}, одна или обе даты некорректны. Проверьте правильность ввода.',
            parse_mode=ParseMode.MARKDOWN
        )
        return

    # Сохраняем даты в контексте
    context.user_data['dates'] = (date1, date2)

    # Показываем сообщение о начале выгрузки
    loading_msg = await update.message.reply_text(
        f"⏳ {user_name}, выгружаю данные с портала...\nПримерное время ожидания 1-2 минуты"
    )

    try:
        # Преобразуем даты в формат для парсинга
        start_date = date1 + "2100"
        end_date = date2 + "2059"

        # Выгружаем первый файл с портала
        await context.bot.edit_message_text(
            chat_id=update.message.chat_id,
            message_id=loading_msg.message_id,
            text=f"⏳ {user_name}, выгружаю данные с портала...\n\n📥 Соединение с порталом..."
        )

        success = await parcing_data_MM_async(start_date, end_date)

        if not success:
            await context.bot.edit_message_text(
                chat_id=update.message.chat_id,
                message_id=loading_msg.message_id,
                text=f"❌ {user_name}, ошибка при выгрузке данных с портала"
            )
            context.user_data['waiting_for_dates'] = False
            return

        # Обновляем сообщение, НЕ УДАЛЯЕМ ЕГО
        await context.bot.edit_message_text(
            chat_id=update.message.chat_id,
            message_id=loading_msg.message_id,
            text=f"✅ {user_name}, данные успешно выгружены с портала!\n\n📤Отправьте городскую выгрузку (Excel файл)\n\nФайл будет обработан автоматически"
        )

        # Устанавливаем состояние ожидания файла от пользователя
        context.user_data['waiting_for_dates'] = False
        context.user_data['waiting_for_file'] = True
        context.user_data['processing_step'] = 'first_file'
        # Сохраняем ID сообщения с инструкцией, чтобы удалить его позже
        context.user_data['instruction_message_id'] = loading_msg.message_id

    except Exception as e:
        logger.error(f"Ошибка при выгрузке данных с портала: {e}")
        await context.bot.edit_message_text(
            chat_id=update.message.chat_id,
            message_id=loading_msg.message_id,
            text=f"❌ {user_name}, ошибка при выгрузке: {str(e)[:100]}..."
        )
        context.user_data['waiting_for_dates'] = False


# Обработчик файлов от пользователя
async def handle_file_upload(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка загрузки файлов от пользователя"""
    
    user = update.effective_user
    user_name = get_user_name(user)

    # Проверяем, если это файл для ОАТИ
    if context.user_data.get('waiting_for_oati_file', False):
        await handle_oati_file(update, context)
        return

    if not context.user_data.get('waiting_for_file', False):
        return

    if update.message.document:
        file = await context.bot.get_file(update.message.document.file_id)

        # Проверяем, что это Excel файл
        file_name = update.message.document.file_name.lower()
        if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
            await update.message.reply_text(
                f"❌ {user_name}, пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)",
                parse_mode=ParseMode.MARKDOWN
            )
            return

        # Определяем директорию для загрузок
        home_dir = os.path.expanduser("~")
        directory = os.path.join(home_dir, "Downloads")
        temp_dir = os.path.join(directory, 'temp')
        os.makedirs(temp_dir, exist_ok=True)

        if context.user_data.get('processing_step') == 'first_file':
            # Сохраняем файл от пользователя
            user_file_path = os.path.join(temp_dir, 'user_file.xlsx')

            # Отправляем подтверждение получения файла
            file_received_msg = await update.message.reply_text(
                f"✅ {user_name}, файл получен. Обрабатываю данные...",
                parse_mode=ParseMode.MARKDOWN
            )

            # Скачиваем файл от пользователя
            await file.download_to_drive(user_file_path)

            # Показываем анимацию загрузки
            loading_msg = await update.message.reply_text(
                f"🔄 {user_name}, обрабатываю данные..."
            )

            try:
                # Получаем сохраненные даты
                date1, date2 = context.user_data.get('dates', ('', ''))

                await context.bot.edit_message_text(
                    chat_id=update.message.chat_id,
                    message_id=loading_msg.message_id,
                    text=f"🔄 {user_name}, нахожу файлы для обработки..."
                )

                # Находим последний скачанный файл с портала
                files = os.listdir(directory)
                excel_files = [f for f in files if f.endswith('.xlsx') or f.endswith('.xls')]
                if not excel_files:
                    raise Exception("Не найдены Excel файлы в папке загрузок")

                # Ищем самый новый файл
                excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
                downloaded_file = excel_files[-1]
                downloaded_file_path = os.path.join(directory, downloaded_file)

                await context.bot.edit_message_text(
                    chat_id=update.message.chat_id,
                    message_id=loading_msg.message_id,
                    text=f"⚙️ {user_name}, обрабатываю файлы..."
                )

                # Обрабатываем оба файла
                output_file_path = process_file_MM_week(user_file_path, downloaded_file_path)

                await context.bot.edit_message_text(
                    chat_id=update.message.chat_id,
                    message_id=loading_msg.message_id,
                    text=f"📤 {user_name}, отправляю файл (это может занять некоторое время)..."
                )

                current_time = datetime.now().strftime('%d.%m.%Y %H:%M')

                # Отправляем результат с увеличенным таймаутом
                try:
                    with open(output_file_path, 'rb') as f:
                        # Используем более длинный таймаут для отправки файла
                        await asyncio.wait_for(
                            context.bot.send_document(
                                chat_id=update.message.chat_id,
                                document=InputFile(f, filename=f"Все_{date1}_{date2}.xlsx"),
                                caption=f"📎 Еженедельный свод за период {date1}-{date2}\n(выгрузка: {current_time})"
                            ),
                            timeout=120.0  # 2 минуты на отправку файла
                        )
                except asyncio.TimeoutError:
                    # Если отправка заняла слишком много времени, но файл все равно мог отправиться
                    logger.warning("Таймаут при отправке файла, но операция могла завершиться успешно")
                    # Проверяем, отправлен ли файл
                    await update.message.reply_text(
                        f"⏳ {user_name}, файл обрабатывается... Проверяю статус отправки..."
                    )

                # Удаляем сообщение о загрузке и инструкцию
                try:
                    # Удаляем сообщение с инструкцией отправки файла
                    instruction_msg_id = context.user_data.get('instruction_message_id')
                    if instruction_msg_id:
                        await context.bot.delete_message(
                            chat_id=update.message.chat_id,
                            message_id=instruction_msg_id
                        )
                except Exception as e:
                    logger.warning(f"Не удалось удалить сообщение с инструкцией: {e}")

                try:
                    # Удаляем сообщение о получении файла
                    await context.bot.delete_message(
                        chat_id=update.message.chat_id,
                        message_id=file_received_msg.message_id
                    )
                except Exception as e:
                    logger.warning(f"Не удалось удалить сообщение о получении файла: {e}")

                # Удаляем сообщение о обработке
                try:
                    await context.bot.delete_message(
                        chat_id=update.message.chat_id,
                        message_id=loading_msg.message_id
                    )
                except Exception as e:
                    logger.warning(f"Не удалось удалить сообщение о загрузке: {e}")

                # Показываем меню с именем пользователя
                await update.message.reply_text(
                    f"✅ Отчет успешно сформирован и отправлен для {user_name}!\n\nВыберите следующую команду:",
                    reply_markup=MAIN_KEYBOARD
                )

                # Очищаем состояние
                context.user_data['waiting_for_file'] = False
                context.user_data['processing_step'] = None
                context.user_data['dates'] = None
                context.user_data['instruction_message_id'] = None

                # Удаляем временные файлы
                try:
                    os.remove(user_file_path)
                except:
                    pass

            except asyncio.TimeoutError:
                # Обработка таймаута отдельно
                logger.error("Таймаут при обработке файла")
                await context.bot.edit_message_text(
                    chat_id=update.message.chat_id,
                    message_id=loading_msg.message_id,
                    text=f"⏳ {user_name}, операция заняла слишком много времени, но файл мог быть отправлен. Проверьте чат."
                )

                await update.message.reply_text(
                    "Выберите команду:",
                    reply_markup=MAIN_KEYBOARD
                )

                # Очищаем состояние
                context.user_data['waiting_for_file'] = False
                context.user_data['processing_step'] = None

            except Exception as e:
                logger.error(f"Ошибка при обработке еженедельного свода: {e}")
                await context.bot.edit_message_text(
                    chat_id=update.message.chat_id,
                    message_id=loading_msg.message_id,
                    text=f"❌ {user_name}, ошибка при обработке: {str(e)[:100]}..."
                )

                await update.message.reply_text(
                    "Выберите команду:",
                    reply_markup=MAIN_KEYBOARD
                )

                # Очищаем состояние
                context.user_data['waiting_for_file'] = False
                context.user_data['processing_step'] = None

    else:
        await update.message.reply_text(
            f'❌ {user_name}, пожалуйста, отправьте документ.',
            parse_mode=ParseMode.MARKDOWN
        )


# Обработчик создания слайда ОАТИ
async def oati_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка запроса на создание слайда ОАТИ"""
    query = update.callback_query
    user_name = get_user_name(query.from_user)
    await query.answer()

    # Удаляем сообщение с кнопками и показываем загрузку
    loading_msg_id = await delete_message_and_show_loading(
        query,
        context,
        f"🅾️ Создаю слайд ОАТИ..."
    )

    try:
        # Запрашиваем файл у пользователя
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"🅾️ Создаю слайд ОАТИ...\n👤 Запрос от {user_name}\n\n📤 Пришлите выгрузку для создания слайда ОАТИ"
        )

        # Устанавливаем состояние ожидания файла ОАТИ
        context.user_data['waiting_for_oati_file'] = True
        context.user_data['loading_msg_id'] = loading_msg_id

    except Exception as e:
        logger.error(f"Ошибка в обработчике ОАТИ: {e}")
        await update_loading_message(
            query.message.chat_id,
            loading_msg_id,
            context,
            f"❌ Ошибка: {str(e)}"
        )


async def handle_oati_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработка файлов ОАТИ"""
    user = update.effective_user
    user_name = get_user_name(user)
    
    if not update.message.document:
        await update.message.reply_text(
            f'❌ {user_name}, пожалуйста, отправьте документ.',
            parse_mode=ParseMode.MARKDOWN
        )
        return

    file = await context.bot.get_file(update.message.document.file_id)

    # Проверяем, что это Excel файл
    file_name = update.message.document.file_name.lower()
    if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
        await update.message.reply_text(
            f"❌ {user_name}, пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)",
            parse_mode=ParseMode.MARKDOWN
        )
        return

    # Получаем ID сообщения загрузки
    loading_msg_id = context.user_data.get('loading_msg_id')
    if not loading_msg_id:
        loading_msg_id = await show_loading_animation(
            update.message.chat_id,
            context,
            f"🅾️ {user_name}, обрабатываю файл ОАТИ..."
        )

    try:
        await update_loading_message(
            update.message.chat_id,
            loading_msg_id,
            context,
            f"🅾️ {user_name}, обрабатываю файл ОАТИ...\n\n📥 Скачиваю файл..."
        )

        # Создаем временную директорию
        temp_dir = os.path.join(directory, 'temp')
        os.makedirs(temp_dir, exist_ok=True)
        temp_file_path = os.path.join(temp_dir, f"oati_file_{datetime.now().strftime('%H%M%S')}.xlsx")

        # Скачиваем файл
        await file.download_to_drive(temp_file_path)

        await update_loading_message(
            update.message.chat_id,
            loading_msg_id,
            context,
            f"🅾️ {user_name}, обрабатываю файл ОАТИ...\n\n⚙️ Анализирую данные..."
        )

        # Обрабатываем файл (теперь функция возвращает 3 значения)
        ppt_path, message = process_file_OATI(temp_file_path)

        await update_loading_message(
            update.message.chat_id,
            loading_msg_id,
            context,
            f"🅾️ {user_name}, обрабатываю файл ОАТИ...\n\n📤 Отправляю файлы..."
        )

        # Отправляем PPT файл
        with open(ppt_path, 'rb') as ppt_file:
            await context.bot.send_document(
                chat_id=update.message.chat_id,
                document=InputFile(ppt_file, filename=os.path.basename(ppt_path)),
                caption=f"🅾️ Слайд ОАТИ для {user_name}"
            )

        # ОТПРАВЛЯЕМ СТАТИСТИЧЕСКОЕ СООБЩЕНИЕ
        await context.bot.send_message(
            chat_id=update.message.chat_id,
            text=message,
            parse_mode=ParseMode.MARKDOWN
        )

        # Удаляем сообщение о загрузке
        await context.bot.delete_message(
            chat_id=update.message.chat_id,
            message_id=loading_msg_id
        )

        # Показываем меню с именем пользователя
        await update.message.reply_text(
            f"✅ Слайд ОАТИ успешно создан и отправлен для {user_name}!\n\nВыберите следующую команду:",
            reply_markup=MAIN_KEYBOARD
        )

        # Очищаем состояние
        context.user_data['waiting_for_oati_file'] = False
        context.user_data['loading_msg_id'] = None

        # Удаляем временные файлы
        try:
            os.remove(temp_file_path)
        except:
            pass

    except Exception as e:
        logger.error(f"Ошибка при обработке файла ОАТИ: {e}")

        if loading_msg_id:
            await update_loading_message(
                update.message.chat_id,
                loading_msg_id,
                context,
                f"❌ {user_name}, ошибка при обработке файла ОАТИ: {str(e)[:100]}..."
            )

        await update.message.reply_text(
            "Выберите команду:",
            reply_markup=MAIN_KEYBOARD
        )

        # Очищаем состояние
        context.user_data['waiting_for_oati_file'] = False
        context.user_data['loading_msg_id'] = None


# Основная функция
def main() -> None:
    """Запуск бота"""
    application = Application.builder() \
        .token(TOKEN) \
        .connect_timeout(60.0) \
        .read_timeout(60.0) \
        .write_timeout(60.0) \
        .pool_timeout(60.0) \
        .build()

    # Регистрируем обработчики
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CallbackQueryHandler(explain_commands, pattern='^explain$'))
    application.add_handler(CallbackQueryHandler(lk_prefekt_handler, pattern='^lk_prefekt$'))
    application.add_handler(CallbackQueryHandler(ng_answers_handler, pattern='^ng_answers$'))
    application.add_handler(CallbackQueryHandler(mm_monitor_handler, pattern='^mm_monitor$'))
    application.add_handler(CallbackQueryHandler(week_svod_handler, pattern='^week_svod$'))
    application.add_handler(CallbackQueryHandler(oati_handler, pattern='^oati$'))
    application.add_handler(CallbackQueryHandler(mji_svod_handler, pattern='^mji_svod$'))

    # Обработчики для еженедельного свода
    application.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND,
        handle_dates_input
    ))

    application.add_handler(MessageHandler(
        filters.Document.ALL,
        handle_file_upload
    ))

    print("🤖 Бот запущен...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == '__main__':
    main()
