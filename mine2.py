import logging
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
REGISTER, DATE, INTERVAL, VIEW, CALC_SALARY, CALC_HOURS = range(6)

# Путь к Excel-файлу
EXCEL_FILE = 'data.xlsx'


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
            f"Зачем перезапускаешь, {fio}!",
            reply_markup=main_menu()
        )
        return ConversationHandler.END


def main_menu():
    """Создаёт главное меню с кнопками."""
    keyboard = [
        ['Добавить СВУ или РВД'],
        ['Просмотреть уже внесенное'],
        ['Узнать стоимость часа РВД и СВУ'],
        ['/cancel']
    ]
    return ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)


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
                reply_markup=main_menu()
            )
            return ConversationHandler.END

    ws_users.append([user_id, fio])
    wb.save(EXCEL_FILE)

    await update.message.reply_text(
        f"Спасибо за регистрацию, {fio}!",
        reply_markup=main_menu()
    )
    return ConversationHandler.END

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
        "Пожалуйста, введите интервал активности (например, РВД с 17 до 22):"
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
        reply_markup=main_menu()
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
    await update.message.reply_text(
        "Операция отменена.",
        reply_markup=main_menu()
    )
    return ConversationHandler.END

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатывает нажатие кнопок главного меню."""
    text = update.message.text
    if text == 'Добавить СВУ или РВД':
        return await add_activity_start(update, context)
    elif text == 'Просмотреть уже внесенное':
        return await view_activities_start(update, context)
    elif text == 'Узнать стоимость часа РВД и СВУ':
        return await calc_salary_start(update, context)
    else:
        await update.message.reply_text(
            "Пожалуйста, используйте кнопки меню.",
            reply_markup=main_menu()
        )

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
            reply_markup=main_menu()
        )
        return ConversationHandler.END

    message = "Ваши активности:\n\n"
    for idx, activity in enumerate(user_activities, start=1):
        date = activity[2]
        interval = activity[3]
        message += f"{idx}. Дата: {date}\n   Интервал: {interval}\n\n"

    await update.message.reply_text(message, reply_markup=main_menu())
    return ConversationHandler.END


async def calc_salary_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начинает процесс расчета оплаты за час."""
    await update.message.reply_text(
        "Введите ваш оклад с учетом северных и районных надбавок (умножь свой голый оклад на 1,6):",
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
        await update.message.reply_text(
            "Пожалуйста, введите правильное числовое значение оклада:"
        )
        return CALC_SALARY

    await update.message.reply_text(
        "Введите числом норму часов в месяце:",
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
        f"1 час СВУ стоит в указанном месяце - {svu} ₽\n"
        f"1 час РВД в указанном месяце - {rvd} ₽"
    )

    await update.message.reply_text(message, reply_markup=main_menu())
    return ConversationHandler.END


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

    # ConversationHandler для добавления активности
    add_activity_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Добавить СВУ или РВД)$'), add_activity_start)],
        states={
            DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_activity_date)],
            INTERVAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_activity_interval)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # ConversationHandler для просмотра активностей
    view_activities_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Просмотреть уже внесенное)$'), view_activities_start)],
        states={
            VIEW: [MessageHandler(filters.ALL, handle_message)],
        },

        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # ConversationHandler для расчета оплаты за час
    calc_salary_handler = ConversationHandler(
        entry_points=[MessageHandler(filters.Regex('^(Узнать стоимость часа РВД и СВУ)$'), calc_salary_start)],
        states={
            CALC_SALARY: [MessageHandler(filters.TEXT & ~filters.COMMAND, calc_salary)],
            CALC_HOURS: [MessageHandler(filters.TEXT & ~filters.COMMAND, calc_hours)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    # Обработчик сообщений для главного меню
    application.add_handler(register_handler)
    application.add_handler(add_activity_handler)
    application.add_handler(view_activities_handler)
    application.add_handler(calc_salary_handler)
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Запуск бота
    application.run_polling()

if __name__ == '__main__':
    main()
