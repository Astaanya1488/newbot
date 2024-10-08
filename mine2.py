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
from openpyxl import Workbook
import os

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Константы состояний для ConversationHandler
REGISTER, DATE, INTERVAL, VIEW, CALC_SALARY, CALC_HOURS, TRANSFER_ACTIVITY, INTERVAL_INPUT, SICK_LEAVE_OPEN_DATE = range(9)

# Путь к Excel-файлу
EXCEL_FILE = 'data.xlsx'

# Список специальных user IDs для доступа к новой функции
SPECIAL_USER_IDS = [461549398, 402468895]


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

        wb.save(EXCEL_FILE)


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
            f"Добро пожаловать обратно, {fio}!",
            reply_markup=main_menu(user_id)
        )
        return ConversationHandler.END


def main_menu(user_id):
    """Создаёт главное меню с кнопками. Добавляет 'Скачать таблицу' для специальных пользователей."""
    keyboard = [
        ['Добавить активность', 'Перенос активности'],
        ['Просмотреть активности'],
        ['Я открыл больничный', 'Я закрыл больничный'],
        ['Рассчитать оплату за час']
    ]

    # Добавляем кнопку "Скачать таблицу" только для специальных пользователей
    if user_id in SPECIAL_USER_IDS:
        keyboard.append(['Скачать таблицу', 'Очистить таблицу'])

    # Добавляем кнопку отмены
    keyboard.append(['/cancel'])

    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

async def sick_leave_open(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие на кнопку 'Я открыл больничный'."""
    await update.message.reply_text("Введите дату следующего приёма (пример 21.10.2024):")
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
    # Проверим, зарегистрирован ли пользователь
    user_id = update.effective_user.id
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws_users = wb["Users"]
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


def validate_date(date_text):
    """Проверяет, соответствует ли дата формату ДД.ММ.ГГГГ."""
    import datetime
    try:
        datetime.datetime.strptime(date_text, '%d.%m.%Y')
        return True
    except ValueError:
        return False


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена процесса добавления активности или регистрации."""
    user_id = update.effective_user.id
    await update.message.reply_text(
        "Операция отменена.",
        reply_markup=main_menu(user_id)
    )
    return ConversationHandler.END


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие кнопок главного меню."""
    text = update.message.text
    user_id = update.effective_user.id

    if text == 'Добавить активность':
        return await add_activity_start(update, context)
    elif text == 'Просмотреть активности':
        return await view_activities_start(update, context)
    elif text == 'Рассчитать оплату за час':
        return await calc_salary_start(update, context)
    elif text == 'Скачать таблицу':
        return await download_table(update, context)
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


async def interval_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Получает введенный интервал и сохраняет его."""
    interval = update.message.text.strip()
    user_id = update.effective_user.id
    fio = get_user_fio(user_id)
    date = datetime.datetime.now()

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
    """Начинает процесс расчета оплаты за час."""
    user_id = update.effective_user.id
    await update.message.reply_text(
        "Введите ваш оклад с учетом северных и районных (голый оклад*1.6):",
        reply_markup=ReplyKeyboardRemove()
    )
    return CALC_SALARY

async def calc_salary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Сохраняет оклад и запрашивает норму часов."""
    salary_text = update.message.text.strip()
    try:
        salary = float(salary_text.replace(',', '.'))
        if salary <= 0:
            raise ValueError
        context.user_data['salary'] = salary
    except ValueError:
        user_id = update.effective_user.id
        await update.message.reply_text(
            "Пожалуйста, введите правильное числовое значение оклада:"
        )
        return CALC_SALARY

    user_id = update.effective_user.id
    await update.message.reply_text(
        "Введите норму часов в месяце:",
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
            "Пожалуйста, введите правильное числовое значение нормы часов:"
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

def main():
    """Основная функция для запуска бота."""
    # Получите ваш токен от @BotFather и вставьте ниже
    TOKEN = '7380924967:AAGbTShrh6-X59LY_sHX2NUCFvdOaNbzCwQ'  # Замените на ваш токен

    application = ApplicationBuilder().token(TOKEN).build()

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
        entry_points=[MessageHandler(filters.Regex('^(Просмотреть активности)$'), view_activities_start)],
        states={
            VIEW: [MessageHandler(filters.ALL, handle_message)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # ConversationHandler для расчета оплаты за час
    calc_salary_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Рассчитать оплату за час)$'), calc_salary_start)],
        states={
            CALC_SALARY: [MessageHandler(filters.TEXT & ~filters.COMMAND, calc_salary)],
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
    application.add_handler(sick_leave_open_handler)

    # Обработчик для кнопки "Я закрыл больничный"
    sick_leave_close_handler = MessageHandler(
        filters.Regex('^Я закрыл больничный$'), sick_leave_close
    )
    # Обработчик сообщений для главного меню
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
