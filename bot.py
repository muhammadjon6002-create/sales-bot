import logging
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE = "candidates.xlsx"
ADMIN_CHAT_ID = int(os.environ.get("ADMIN_CHAT_ID", "0"))

# Шаги
(S1, S2, S3, S4, S5, S6, S7, S8, S9, S10,
 S11, S12, S13, S14, S15, S16, S17, S18, S19, S20,
 S21, S22, S23, S24, S25, S26, S27, S28, S29, S30) = range(30)

QUESTIONS = [
    "👤 *1/30* Ваше полное ФИО?",
    "🎂 *2/30* Ваш возраст?",
    "🏙 *3/30* В каком городе вы находитесь?",
    "📞 *4/30* Ваш номер телефона?",
    "📧 *5/30* Ваш email?",
    "🎓 *6/30* Образование (ВУЗ, специальность, год окончания)?",
    "📚 *7/30* Дополнительные курсы или сертификаты в сфере продаж/маркетинга?",
    "📅 *8/30* Сколько лет общего опыта в продажах/маркетинге?",
    "🏢 *9/30* Перечислите последние 3 места работы (компания и период)?",
    "💼 *10/30* Какие должности вы занимали?",
    "🏆 *11/30* Ваши главные достижения в продажах (с конкретными цифрами)?",
    "🎯 *12/30* Какие техники продаж вы знаете и применяете?",
    "🛠 *13/30* Какие CRM-системы и маркетинговые инструменты вы используете?",
    "📊 *14/30* Какие KPI вы выполняли на последнем месте работы?",
    "💰 *15/30* Средний чек сделки, с которой вы работали?",
    "🤝 *16/30* Какие типы клиентов: B2B, B2C или оба?",
    "📦 *17/30* Что именно вы продавали (товары/услуги/digital)?",
    "💡 *18/30* Почему вас интересует именно наша вакансия?",
    "💵 *19/30* Ваши ожидания по зарплате (фикс + бонус)?",
    "🕐 *20/30* Предпочтительный формат работы?",
    "🌍 *21/30* Уровень английского языка?",
    "🚗 *22/30* Есть ли у вас водительские права и автомобиль?",
    "✈️ *23/30* Готовы ли вы к командировкам?",
    "🏠 *24/30* Рассматриваете ли вы удалённую работу?",
    "🔗 *25/30* Ссылка на ваш LinkedIn или HH.ru профиль (или напишите «нет»)?",
    "📁 *26/30* Есть ли портфолио или кейсы? Укажите ссылку или опишите.",
    "⭐️ *27/30* Ваши 3 главных профессиональных сильных стороны?",
    "🔍 *28/30* Над чем вы сейчас работаете / что хотите улучшить в себе?",
    "📝 *29/30* Расскажите коротко о себе (2–3 предложения)?",
    "❓ *30/30* Есть ли у вас вопросы к нашей компании или вакансии?",
]

LABELS = [
    "ФИО", "Возраст", "Город", "Телефон", "Email",
    "Образование", "Курсы/Сертификаты", "Опыт (лет)", "Места работы", "Должности",
    "Достижения", "Техники продаж", "CRM/Инструменты", "KPI", "Средний чек",
    "Тип клиентов", "Что продавал(а)", "Мотивация", "Ожидания ЗП", "График",
    "Английский", "Водительские права", "Командировки", "Удалёнка",
    "LinkedIn/HH", "Портфолио", "Сильные стороны", "Зоны роста",
    "О себе", "Вопросы к компании"
]

KB_SCHEDULE = ReplyKeyboardMarkup([["Офис", "Удалёнка", "Гибрид"]], resize_keyboard=True, one_time_keyboard=True)
KB_ENGLISH  = ReplyKeyboardMarkup([["Нет", "Базовый", "Intermediate"], ["Upper-Intermediate", "Fluent"]], resize_keyboard=True, one_time_keyboard=True)
KB_DRIVING  = ReplyKeyboardMarkup([["Да, есть права и авто", "Есть права, нет авто", "Нет"]], resize_keyboard=True, one_time_keyboard=True)
KB_YESNO    = ReplyKeyboardMarkup([["Да", "Нет"]], resize_keyboard=True, one_time_keyboard=True)
KB_CLIENTS  = ReplyKeyboardMarkup([["B2B", "B2C", "Оба"]], resize_keyboard=True, one_time_keyboard=True)
KB_REMOVE   = ReplyKeyboardRemove()

KEYBOARDS = {
    S16: KB_CLIENTS,
    S20: KB_SCHEDULE,
    S21: KB_ENGLISH,
    S22: KB_DRIVING,
    S23: KB_YESNO,
    S24: KB_YESNO,
}


def init_excel():
    if os.path.exists(EXCEL_FILE):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Кандидаты"
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(bold=True, color="FFFFFF", size=11)
    headers = ["№", "Дата"] + LABELS
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 18
    for i in range(3, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 30
    wb.save(EXCEL_FILE)


def save_excel(answers):
    init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    row_n = ws.max_row + 1
    alt = PatternFill("solid", fgColor="EBF0FA")
    bd = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    row = [row_n - 1, datetime.now().strftime("%d.%m.%Y %H:%M")] + answers
    for col, val in enumerate(row, 1):
        cell = ws.cell(row=row_n, column=col, value=val)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = bd
        if row_n % 2 == 0:
            cell.fill = alt
    ws.row_dimensions[row_n].height = 55
    wb.save(EXCEL_FILE)


def generate_pdf(answers, name):
    fname = f"resume_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    doc = SimpleDocTemplate(fname, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t", parent=styles["Title"], fontSize=16,
                             textColor=colors.HexColor("#1F3864"), spaceAfter=4)
    sub_s   = ParagraphStyle("s", parent=styles["Normal"], fontSize=9,
                             textColor=colors.grey, spaceAfter=14)
    ans_s   = ParagraphStyle("a", parent=styles["Normal"], fontSize=9, leading=13)

    story = [
        Paragraph("Анкета кандидата — Продажи / Маркетинг", title_s),
        Paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}", sub_s),
    ]
    data = []
    for label, val in zip(LABELS, answers):
        data.append([
            Paragraph(f"<b>{label}</b>", ans_s),
            Paragraph(str(val) if val else "—", ans_s),
        ])
    t = Table(data, colWidths=[5*cm, 12*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EBF0FA")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#F7F9FD")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(t)
    doc.build(story)
    return fname


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["answers"] = []
    await update.message.reply_text(
        "👋 Добро пожаловать!\n\n"
        "Это анкета для кандидатов на вакансию *Продажи / Маркетинг*.\n\n"
        "📋 Вас ждут *30 вопросов* — отвечайте честно и развёрнуто.\n"
        "⏱ Займёт около 10–15 минут.\n\n"
        "Нажмите кнопку чтобы начать 👇",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["✅ Начать анкету"]], resize_keyboard=True)
    )
    return S1


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get("step", 0)
    answers = context.user_data.get("answers", [])

    # Сохраняем ответ (кроме первого нажатия «Начать»)
    text = update.message.text
    if text != "✅ Начать анкету":
        answers.append(text)
        context.user_data["answers"] = answers

    next_step = len(answers)  # следующий вопрос = индекс

    if next_step >= 30:
        return await finish(update, context)

    kb = KEYBOARDS.get(next_step, KB_REMOVE)
    await update.message.reply_text(
        QUESTIONS[next_step],
        parse_mode="Markdown",
        reply_markup=kb
    )
    return next_step


async def finish(update: Update, context: ContextTypes.DEFAULT_TYPE):
    answers = context.user_data.get("answers", [])
    user = update.effective_user
    name = answers[0] if answers else (user.username or str(user.id))
    username = user.username or str(user.id)

    await update.message.reply_text("⏳ Сохраняю анкету...", reply_markup=KB_REMOVE)

    save_excel(answers)
    pdf = generate_pdf(answers, name)

    await update.message.reply_text(
        "✅ *Анкета принята!*\n\n"
        "Спасибо за уделённое время. Мы свяжемся с вами в течение 2–3 рабочих дней. Удачи! 🍀",
        parse_mode="Markdown"
    )

    if ADMIN_CHAT_ID:
        await context.bot.send_message(
            ADMIN_CHAT_ID,
            f"📥 *Новая анкета!*\n\n"
            f"👤 {name}\n"
            f"📱 @{username}\n"
            f"📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}",
            parse_mode="Markdown"
        )
        with open(pdf, "rb") as f:
            await context.bot.send_document(ADMIN_CHAT_ID, f, filename=f"Анкета_{name}.pdf")
        await context.bot.send_document(ADMIN_CHAT_ID, open(EXCEL_FILE, "rb"), filename="Все_кандидаты.xlsx")

    os.remove(pdf)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Анкета отменена. Напишите /start чтобы начать заново.", reply_markup=KB_REMOVE)
    return ConversationHandler.END


def main():
    token = os.environ.get("BOT_TOKEN")
    if not token:
        raise ValueError("BOT_TOKEN не задан!")

    app = ApplicationBuilder().token(token).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={i: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle)] for i in range(30)},
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv)
    init_excel()
    logger.info("✅ Бот запущен!")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
