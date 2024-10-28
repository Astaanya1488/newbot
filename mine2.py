import logging
import datetime
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    ConversationHandler,
    filters,
)
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Константы состояний для ConversationHandler
REGISTER, DATE, INTERVAL, VIEW, CALC_HOURS, TRANSFER_ACTIVITY, INTERVAL_INPUT, SICK_LEAVE_OPEN_DATE, COLOR_ROWS, NOTIFY_MESSAGE,BAN_USER, BAN_DATE, BAN_REASON = range(13)

# Путь к Excel-файлу
EXCEL_FILE = 'data.xlsx'

# Список специальных user IDs для доступа к новой функции
SPECIAL_USER_IDS = [461549398, 402468895,1352307342]


def init_excel():
    """Инициализирует Excel-файл с двумя листами, если он не существует."""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()

        # Создание первого листа "Users"
        ws_users = wb.active
        ws_users.title = "Users"
        ws_users.append(["UserID", "ФИО"])

        # Создание второго листа "Activities"
        ws_activities = wb.create_sheet(title="Activities")
        ws_activities.append(["UserID", "ФИО", "Дата", "Интервал"])

        # Создание листа "Transfers" для переноса активности
        ws_transfers = wb.create_sheet(title="Transfers")
        ws_transfers.append(["UserID", "ФИО", "Дата", "Интервал"])
        #Создание третьего листа "Banned"
        ws_banned = wb.create_sheet("Banned")
        ws_banned.append(["UserID", "ФИО", "Дата окончания бана", "Причина"])

        wb.save(EXCEL_FILE)

def get_users():
    """
    Считывает список зарегистрированных пользователей из листа "Users" Excel-файла.
    Возвращает список UserID (int).
    """
    if not os.path.exists(EXCEL_FILE):
        init_excel()

    wb = load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]

    users = []
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        user_id = row[0]
        if isinstance(user_id, int):
            users.append(user_id)
    return users

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start. Инициирует регистрацию или показывает главное меню."""
    user_id = update.effective_user.id
    username = update.effective_user.username or update.effective_user.first_name

    # Инициализируем Excel, если нужно
    init_excel()

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]

    # Проверим, зарегистрирован ли пользователь (по UserID)
    registered = False
    fio = None
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            registered = True
            fio = row[1]
            break

    if not registered:
        await update.message.reply_text(
            "Привет! Для начала регистрации, пожалуйста, введи своё ФИО:",
            reply_markup=ReplyKeyboardRemove()
        )
        return REGISTER
    else:
        await update.message.reply_text(
            f"С возвращением, {fio}!",
            reply_markup=main_menu(user_id)
        )
        return ConversationHandler.END

def main_menu(user_id):
    """Создаёт главное меню с кнопками. Добавляет 'Скачать таблицу' для специальных пользователей."""
    keyboard = [
        ['Активности', 'Больничный'],
        ['Рассчитать оплату за час']
    ]

    # Добавляем кнопку "Особых действий" только для специальных пользователей
    if user_id in SPECIAL_USER_IDS:
        keyboard.append(['Особые действия'])

    # Добавляем кнопку отмены
    keyboard.append(['Отмена'])

    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def activities_menu():
    """Создаёт подменю для 'Активности'."""
    keyboard = [
        ['Добавить активность', 'Перенос активности'],
        ['Внесенные активности', 'Проставленные активности'],
        ['Назад']
    ]

    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def sick_menu():
    """Создаёт подменю для 'Больничный'."""
    keyboard = [
        ['Я открыл больничный', 'Я закрыл больничный', 'Больничный продлен'],
        ['Назад']
    ]

    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

def special_menu():
    """Создаёт подменю для 'Специальных пользователей'."""
    keyboard = [
        ['Скачать таблицу', 'Очистить таблицу', ],
        ['Закрасить строки', 'Оповестить всех', 'Внести в бан'],
        ['Назад']
    ]
    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

async def ban_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Введите ID пользователя для бана:",
        reply_markup=ReplyKeyboardRemove()
    )
    return BAN_USER

async def ban_user_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.text.strip()
    if not user_id.isdigit():
        await update.message.reply_text("ID должен быть числом. Попробуйте снова:")
        return BAN_USER
    context.user_data['ban_user_id'] = int(user_id)
    await update.message.reply_text("Введите дату окончания блокировки (ДД.ММ.ГГГГ):")
    return BAN_DATE

def validate_date(date_text):
    try:
        datetime.strptime(date_text, '%d.%m.%Y')
        return True
    except ValueError:
        return False

async def ban_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    date_text = update.message.text.strip()
    if not validate_date(date_text):
        await update.message.reply_text(
            "Неверный формат даты. Пожалуйста, введите дату в формате ДД.ММ.ГГГГ:"
        )
        return BAN_DATE
    context.user_data['ban_end_date'] = date_text
    await update.message.reply_text("Укажите причину блокировки:")
    return BAN_REASON

async def ban_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    reason = update.message.text.strip()
    if not reason:
        await update.message.reply_text("Причина не может быть пустой. Укажите причину блокировки:")
        return BAN_REASON
    context.user_data['ban_reason'] = reason

    # Сохранение в Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_banned = wb["Banned"]
    ban_user_id = context.user_data['ban_user_id']

    # Получение ФИО пользователя из листа "Users"
    ws_users = wb["Users"]
    fio = None
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == ban_user_id:
            fio = row[1]
            break
    if not fio:
        await update.message.reply_text("Пользователь с таким ID не найден.")
        return ConversationHandler.END

    ws_banned.append([ban_user_id, fio, context.user_data['ban_end_date'], reason])
    wb.save(EXCEL_FILE)

    await update.message.reply_text("Пользователь успешно забанен!", reply_markup=main_menu(ban_user_id))
    return ConversationHandler.END

async def ban_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Процесс бана отменён.", reply_markup=main_menu(update.effective_user.id))
    return ConversationHandler.END

async def back_to_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возвращает пользователя в главное меню."""
    user_id = update.effective_user.id
    await update.message.reply_text(
        "Главное меню:",
        reply_markup=main_menu(user_id)
    )

async def sick_leave_open(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие на кнопку 'Я открыл больничный'."""
    await update.message.reply_text("Введите дату следующего приёма (пример 21.10.2024):")
    return SICK_LEAVE_OPEN_DATE

async def sick_leave_return(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие на кнопку 'Больничный продлен'."""
    await update.message.reply_text("Введите дату нового приёма (пример 17.12.2024):")
    return SICK_LEAVE_OPEN_DATE

async def sick_leave_close(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие на кнопку 'Я закрыл больничный'."""
    user_id = update.effective_user.id
    # Получаем ФИО пользователя из листа "Users"
    fio = None

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            fio = row[1]
            break
    wb.close()

    if not fio:
        await update.message.reply_text("Не удалось найти ваше ФИО в базе данных.")
        return

    # Формируем сообщение для отправки специальным пользователям
    message = f"{fio} закрыл больничный"

    # Отправляем сообщение специальным пользователям
    for special_user_id in SPECIAL_USER_IDS:
        try:
            await context.bot.send_message(chat_id=special_user_id, text=message)
        except Exception as e:
            logger.error(f"Не удалось отправить сообщение пользователю {special_user_id}: {e}")


    await update.message.reply_text("Информация отправлена ВС.", reply_markup=main_menu(user_id))

async def sick_leave_open_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает ввод даты следующего приёма."""
    next_appointment_date = update.message.text.strip()

    user_id = update.effective_user.id
    # Получаем ФИО пользователя из листа "Users"
    fio = None

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            fio = row[1]
            break
    wb.close()

    if not fio:
        await update.message.reply_text("Не удалось найти ваше ФИО в базе данных.")
        return ConversationHandler.END

    # Формируем сообщение для отправки специальным пользователям
    message = f"{fio} открыл больничный, дата приема - {next_appointment_date}"

    # Отправляем сообщение специальным пользователям
    for special_user_id in SPECIAL_USER_IDS:
        try:
            await context.bot.send_message(chat_id=special_user_id, text=message)
        except Exception as e:
            logger.error(f"Не удалось отправить сообщение пользователю {special_user_id}: {e}")

    await update.message.reply_text("Информация отправлена.", reply_markup=main_menu(user_id))
    return ConversationHandler.END

async def download_table(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отправляет файл data.xlsx специальному пользователю."""
    user_id = update.effective_user.id
    if user_id not in SPECIAL_USER_IDS:
        await update.message.reply_text(
            "У вас нет доступа к этой функции.",
            reply_markup=main_menu(user_id)
        )
        return

    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text(
            "Файл не найден.",
            reply_markup=main_menu(user_id)
        )
        return

    with open(EXCEL_FILE, 'rb') as file:
        await update.message.reply_document(document=file, filename=EXCEL_FILE)

    await update.message.reply_text(
        "Файл отправлен.",
        reply_markup=main_menu(user_id)
    )

async def clear_table(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие на кнопку 'Очистить таблицу'."""
    user_id = update.effective_user.id

    if user_id not in SPECIAL_USER_IDS:
        await update.message.reply_text("У вас нет доступа к этой функции.")
        return

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet_names = wb.sheetnames

        for sheet_name in sheet_names:
            if sheet_name != 'Users':
                ws = wb[sheet_name]
                # Очищаем строки, начиная со второй, сохраняя заголовки
                ws.delete_rows(2, ws.max_row)

        wb.save(EXCEL_FILE)
        wb.close()

        await update.message.reply_text("Таблица успешно очищена, кроме данных на листе 'Users'.")
    except Exception as e:
        logger.error(f"Ошибка при очистке таблицы: {e}")
        await update.message.reply_text(f"Произошла ошибка при очистке таблицы: {e}")

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена текущей операции."""
    await update.message.reply_text(
        'Операция отменена.',
        reply_markup=main_menu(update.effective_user.id)
    )
    return ConversationHandler.END

async def register_fio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Сохраняет ФИО пользователя и завершает регистрацию."""
    fio = update.message.text.strip()
    if not fio:
        await update.message.reply_text(
            "ФИО не может быть пустым. Пожалуйста, введите своё ФИО:"
        )
        return REGISTER

    user_id = update.effective_user.id

    # Записываем ФИО в Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]

    # Проверим ещё раз, зарегистрирован ли пользователь
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            await update.message.reply_text(
                "Вы уже зарегистрированы.",
                reply_markup=main_menu(user_id)
            )
            return ConversationHandler.END

    ws_users.append([user_id, fio])

    wb.save(EXCEL_FILE)

    await update.message.reply_text(
        f"Спасибо за регистрацию, {fio}!",
        reply_markup=main_menu(user_id)
    )
    return ConversationHandler.END

def get_user_fio(user_id):
    """Получает ФИО пользователя по его user_id."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb['Users']
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            return row[1]
    return None

async def add_activity_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс добавления активности."""
    user_id = update.effective_user.id
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]
    ws_banned = wb["Banned"]

    fio = None
    registered = False
    for row in ws_users.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            registered = True
            fio = row[1]
            break

    if not registered:
        await update.message.reply_text(
            "Пожалуйста, зарегистрируйтесь сначала, отправив /start."
        )
        return ConversationHandler.END

    # Проверка бана
    today = datetime.today().date()
    is_banned = False
    ban_end_date = None
    for row in ws_banned.iter_rows(min_row=2, values_only=True):
        banned_user_id, _, end_date_str, _ = row
        if banned_user_id == user_id:
            try:
                end_date = datetime.strptime(end_date_str, '%d.%m.%Y').date()
                if today <= end_date:
                    is_banned = True
                    ban_end_date = end_date
                    break
            except ValueError:
                continue

    if is_banned:
        await update.message.reply_text(
            f"Вы забанены до {ban_end_date.strftime('%d.%m.%Y')}. "
            "Вы не можете добавлять активности до окончания бана."
        )
        return ConversationHandler.END

    context.user_data['fio'] = fio  # Сохраняем ФИО пользователя для дальнейшего использования

    await update.message.reply_text(
        "Пожалуйста, введите дату активности в формате ДД.ММ.ГГГГ:",
        reply_markup=ReplyKeyboardRemove()
    )
    return DATE

async def add_activity_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Сохраняет дату и запрашивает интервал."""
    date_text = update.message.text.strip()
    if not validate_date(date_text):
        await update.message.reply_text(
            "Неверный формат даты. Пожалуйста, введите дату в формате ДД.ММ.ГГГГ:"
        )
        return DATE

    context.user_data['date'] = date_text
    await update.message.reply_text(
        "Пожалуйста, введите интервал активности (пример: СВУ с 12 до 13:45):"
    )
    return INTERVAL

async def add_activity_interval(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Сохраняет интервал и записывает данные в Excel."""
    interval = update.message.text.strip()
    date = context.user_data.get('date')
    fio = context.user_data.get('fio')

    if not interval:
        await update.message.reply_text(
            "Интервал не может быть пустым. Пожалуйста, введите интервал активности:"
        )
        return INTERVAL

    # Записываем данные в Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_activities = wb["Activities"]

    user_id = update.effective_user.id
    ws_activities.append([user_id, fio, date, interval])
    wb.save(EXCEL_FILE)

    await update.message.reply_text(
        "Активность добавлена успешно!",
        reply_markup=main_menu(user_id)
    )

    return ConversationHandler.END

async def show_recorded_activities(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_activities = wb["Activities"]

    recorded_activities = []

    # Определите, какая колонка отвечает за закрашенные ячейки.
    # Предположим, что это колонка "Interval" (например, 4-я колонка)
    INTERVAL_COLUMN_INDEX = 1

    for row in ws_activities.iter_rows(min_row=2):  # Пропустить заголовок
        row_user_id = row[0].value
        if row_user_id != user_id:
            continue

        interval_cell = row[INTERVAL_COLUMN_INDEX - 1]  # Индексация с 0
        fill = interval_cell.fill

        # Проверим, есть ли цвет заливки (исключим пустую заливку)
        if fill and fill.fgColor and fill.fgColor.type != 'none' and fill.fgColor.rgb != '00000000':
            # Получаем необходимые данные, например: ФИО, дату и интервал
            fio = row[1].value
            date = row[2].value
            interval = row[3].value
            recorded_activities.append(f"Дата: {date}, Время: {interval}")

    if not recorded_activities:
        await update.message.reply_text("Пока ничего не проставили.")
    else:
        activities_text = "\n".join(recorded_activities)
        await update.message.reply_text(f"Проставлено:\n{activities_text}", reply_markup=main_menu(user_id))

    wb.close()

async def color_rows_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс закрашивания строк."""
    await update.message.reply_text(
        "Введите номер строки, до которой необходимо закрасить данные (целое число):",
        reply_markup=ReplyKeyboardRemove()
    )
    return COLOR_ROWS

async def color_rows_process(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Закрашивает строки до указанного номера включительно в желтый цвет."""
    row_input = update.message.text.strip()

    if not row_input.isdigit():
        await update.message.reply_text(
            "Пожалуйста, введите корректное целое число для номера строки:"
        )
        return COLOR_ROWS

    row_number = int(row_input)

    if row_number < 2:
        await update.message.reply_text(
            "Минимально допустимый номер строки - 2."
        )
        return COLOR_ROWS

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_activities = wb["Activities"]

        max_row = ws_activities.max_row
        if row_number > max_row:
            await update.message.reply_text(
                f"В таблице только {max_row} строк. Будут закрашены все строки до {max_row}."
            )
            row_number = max_row

        # Определяем желтый цвет
        yellow_fill = PatternFill(start_color="FFFF00",
                                  end_color="FFFF00",
                                  fill_type="solid")

        # Закрашиваем строки от 2 до row_number включительно (предполагая, что 1-я строка - заголовок)
        for row in ws_activities.iter_rows(min_row=2, max_row=row_number, max_col=ws_activities.max_column):
            for cell in row:
                cell.fill = yellow_fill

        wb.save(EXCEL_FILE)

        await update.message.reply_text(
            f"Строки с 2 по {row_number} успешно закрашены в желтый цвет.",
            reply_markup=main_menu(update.effective_user.id)
        )

        return ConversationHandler.END

    except Exception as e:
        logger.error(f"Ошибка при закрашивании строк: {e}")
        await update.message.reply_text(
            "Произошла ошибка при обработке запроса. Пожалуйста, попробуйте позже."
        )
        return ConversationHandler.END

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие кнопок главного меню."""
    text = update.message.text
    user_id = update.effective_user.id

    if text == 'Активности':
        await update.message.reply_text(
            "Выберите действие в разделе 'Активности':",
            reply_markup=activities_menu()
        )
    elif text == 'Больничный':
        await update.message.reply_text(
            "Выберите действие в разделе 'Больничный':",
            reply_markup=sick_menu()
        )
    elif text == 'Особые действия':
        await update.message.reply_text(
            "Выберите действие в разделе для специальных пользователей:",
            reply_markup=special_menu()
        )
    elif text == 'Оповестить всех':
        await update.message.reply_text(
            "Введите текст для отправки всем пользователям:",
            reply_markup=ReplyKeyboardRemove()
        )
        return NOTIFY_MESSAGE
    elif text == 'Скачать таблицу':
        return await download_table(update, context)
    elif text == 'Назад':
        await back_to_main_menu(update, context)
    elif text == 'Отмена':
        await cancel(update, context)
    else:
        await update.message.reply_text(
            "Пожалуйста, используйте кнопки меню.",
            reply_markup=main_menu(user_id)
        )

async def transfer_activity(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Инициирует процесс переноса активности, запрашивая интервал."""
    await update.message.reply_text(
        "Введите интервал переноса (пример: с 7:00 на 7:30):",
        reply_markup=ReplyKeyboardRemove()
    )
    return INTERVAL_INPUT

async def notify_all_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс отправки уведомления всем пользователям."""
    user_id = update.effective_user.id
    if user_id not in SPECIAL_USER_IDS:
        await update.message.reply_text("У вас нет прав для выполнения этого действия.")
        return ConversationHandler.END

    await update.message.reply_text(
        "Введите текст для отправки всем пользователям:",
        reply_markup=ReplyKeyboardRemove()
    )
    return NOTIFY_MESSAGE

async def notify_all_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отправляет введённое сообщение всем зарегистрированным пользователям."""
    user_id = update.effective_user.id
    text = update.message.text

    users = get_users()

    if not users:
        await update.message.reply_text("Список пользователей пуст.")
        return ConversationHandler.END

    sent_count = 0
    failed_users = []

    for uid in users:
        try:
            await context.bot.send_message(chat_id=uid, text=text)
            sent_count += 1
        except Exception as e:
            failed_users.append(uid)
            print(f"Не удалось отправить сообщение пользователю {uid}: {e}")

    response = f"Сообщение отправлено {sent_count} из {len(users)} пользователей."
    if failed_users:
        response += f"\nНе удалось отправить сообщение следующим пользователям: {failed_users}"
    await update.message.reply_text(response)

    # Возвращаемся в главное меню
    await back_to_main_menu(update, context)
    return ConversationHandler.END

async def interval_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Получает введенный интервал и сохраняет его."""
    interval = update.message.text.strip()
    user_id = update.effective_user.id
    fio = get_user_fio(user_id)
    date = datetime.now()

    if not interval:
        await update.message.reply_text(
            "Интервал не может быть пустым. Пожалуйста, введите интервал переноса (пример: с 7:00 на 7:30):"
        )
        return INTERVAL_INPUT

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_transfers = wb['Transfers']

    ws_transfers.append([user_id, fio, date.strftime('%Y-%m-%d %H:%M:%S'), interval])
    wb.save(EXCEL_FILE)

    await update.message.reply_text("Информация передана ВС.", reply_markup=main_menu(user_id))
    return ConversationHandler.END

async def view_activities_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс просмотра активностей."""
    user_id = update.effective_user.id

    wb = openpyxl.load_workbook(EXCEL_FILE)

    ws_activities = wb["Activities"]

    user_activities = []
    for row in ws_activities.iter_rows(min_row=2, values_only=True):
        if row[0] == user_id:
            user_activities.append(row)

    if not user_activities:
        await update.message.reply_text(
            "У вас пока нет добавленных активностей.",
            reply_markup=main_menu(user_id)
        )
        return ConversationHandler.END

    message = "Ваши активности:\n\n"
    for idx, activity in enumerate(user_activities, start=1):
        date = activity[2]
        interval = activity[3]
        message += f"{idx}. Дата: {date}\n   Интервал: {interval}\n\n"

    await update.message.reply_text(message, reply_markup=main_menu(user_id))
    return ConversationHandler.END

async def calc_salary_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс расчета оплаты за час, устанавливая фиксированный оклад."""
    user_id = update.effective_user.id

    # Устанавливаем оклад в зависимости от того, является ли пользователь специальным
    if user_id in SPECIAL_USER_IDS:
        fixed_salary = 51000  # Фиксированный оклад для специальных пользователей
    else:
        fixed_salary = 40329.6  # Стандартный фиксированный оклад

    context.user_data['salary'] = fixed_salary

    await update.message.reply_text(
        f"Фиксированный оклад установлен: {fixed_salary} ₽\n"
        "Введите норму часов в месяце: (узнать по ссылке https://www.consultant.ru/law/ref/calendar/proizvodstvennye/)",
        reply_markup=ReplyKeyboardRemove()
    )
    return CALC_HOURS

async def calc_hours(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Сохраняет норму часов и выполняет расчет."""
    hours_text = update.message.text.strip()
    try:
        hours = float(hours_text.replace(',', '.'))
        if hours <= 0:
            raise ValueError
        context.user_data['hours'] = hours
    except ValueError:
        user_id = update.effective_user.id
        await update.message.reply_text(
            "Пожалуйста, введите правильное числовое значение нормы часов :"
        )
        return CALC_HOURS

    salary = context.user_data.get('salary')
    hours = context.user_data.get('hours')

    svu = ((salary * 0.87) / hours) * 1.5
    rvd = ((salary * 0.87) / hours) * 2

    # Форматирование чисел до 2 знаков после запятой
    svu = round(svu, 2)
    rvd = round(rvd, 2)

    message = (
        f"1 час СВУ стоит = {svu} ₽\n"
        f"1 час РВД стоит = {rvd} ₽"
    )

    user_id = update.effective_user.id
    await update.message.reply_text(message, reply_markup=main_menu(user_id))
    return ConversationHandler.END

#CАМОЕ ГЛАВНОЕ В РАБОТЕ БОТА
def main():
    """Основная функция для запуска бота."""
    # Получите ваш токен от @BotFather и вставьте ниже
    TOKEN = '7380924967:AAGbTShrh6-X59LY_sHX2NUCFvdOaNbzCwQ'  # Замените на ваш токен

    application = ApplicationBuilder().token(TOKEN).build()

    application.add_handler(MessageHandler(filters.Regex('^Проставленные активности$'), show_recorded_activities))
    # Обработчик Conversation для "Внести в бан"
    ban_conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^Внести в бан$"), ban_start)],
        states={
            BAN_USER: [MessageHandler(filters.TEXT & ~filters.COMMAND, ban_user_id)],
            BAN_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, ban_date)],
            BAN_REASON: [MessageHandler(filters.TEXT & ~filters.COMMAND, ban_reason)],
        },
        fallbacks=[CommandHandler('cancel', ban_cancel)],
    )

    # Обработчик Conversation для "Оповестить всех"
    notify_conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^Оповестить всех$"), notify_all_start)],
        states={
            NOTIFY_MESSAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, notify_all_message)]
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        allow_reentry=True
    )
    # ConversationHandler для регистрации
    register_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={

            REGISTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, register_fio)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )
    # Conversation handler for transfer activity
    transfer_activity_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^Перенос активности$'), transfer_activity)],
        states={
            INTERVAL_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, interval_input)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        per_user=True,
        per_chat=True,
    )
    # ConversationHandler для добавления активности и закрашивания строк
    conv_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex("^Закрасить строки$"), color_rows_start)],
        states={
            COLOR_ROWS: [MessageHandler(filters.TEXT & ~filters.COMMAND, color_rows_process)],
        },
        fallbacks=[CommandHandler('cancel', lambda update, context: ConversationHandler.END)],
    )

    # ConversationHandler для добавления активности
    add_activity_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Добавить активность)$'), add_activity_start)],
        states={
            DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_activity_date)],
            INTERVAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_activity_interval)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # ConversationHandler для просмотра активностей
    view_activities_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Внесенные активности)$'), view_activities_start)],
        states={
            VIEW: [MessageHandler(filters.ALL, handle_message)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # ConversationHandler для расчета оплаты за час
    calc_salary_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Рассчитать оплату за час)$'), calc_salary_start)],
        states={
            CALC_HOURS: [MessageHandler(filters.TEXT & ~filters.COMMAND, calc_hours)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )
    # Обработчик для кнопки "Очистить таблицу"
    clear_table_handler = MessageHandler(
        filters.TEXT & filters.Regex('^Очистить таблицу$'),
        clear_table
    )
    # Обработчик для кнопки "Я открыл больничный"
    sick_leave_open_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^Я открыл больничный$'), sick_leave_open)],
        states={
            SICK_LEAVE_OPEN_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sick_leave_open_date)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )
    #Обработчик для кнопки "Больничный продлен"
    sick_leave_return_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^Больничный продлен$'), sick_leave_open)],
        states={
            SICK_LEAVE_OPEN_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sick_leave_open_date)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # Обработчик для кнопки "Я закрыл больничный"
    sick_leave_close_handler = MessageHandler(
        filters.Regex('^Я закрыл больничный$'), sick_leave_close
    )
    # Обработчик сообщений для главного меню
    application.add_handler(ban_conv_handler)
    application.add_handler(notify_conv_handler)
    application.add_handler(sick_leave_return_handler)
    application.add_handler(sick_leave_open_handler)
    application.add_handler(conv_handler)
    application.add_handler(sick_leave_close_handler)
    application.add_handler(clear_table_handler)
    application.add_handler(transfer_activity_handler)
    application.add_handler(register_handler)
    application.add_handler(add_activity_handler)
    application.add_handler(view_activities_handler)
    application.add_handler(calc_salary_handler)

    # Добавляем обработчик для кнопок меню
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Добавляем обработчик для команды /cancel
    application.add_handler(CommandHandler('cancel', cancel))
    # Запуск бота
    application.run_polling()


if __name__ == '__main__':
    main()
