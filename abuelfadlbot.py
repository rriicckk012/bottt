import telebot
import pandas as pd
import os

ADMIN_CODE = "01228314076"
EXCEL_FILE = "products.xlsx"
PROGRAMMER = "مبرمج البوت: احمد شريف"

admins = set()

TOKEN = os.getenv("7856214662:AAHII6D5KgRnermFK17lnChwOa8UMnz2DKE") or "7856214662:AAHII6D5KgRnermFK17lnChwOa8UMnz2DKE"
bot = telebot.TeleBot(TOKEN)

def search_product(query):
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna('')
    except Exception as e:
        return f"حدث خطأ في قراءة ملف الإكسل: {e}"

    # البحث باسم المنتج أو الباركود
    found = df[
        (df['اسم المنتج'].str.contains(query, case=False, na=False)) |
        (df['الباركود'].astype(str) == str(query))
    ]
    if found.empty:
        return None

    results = []
    for _, row in found.iterrows():
        try:
            quantity = float(str(row['الكمية']).replace(',', '').strip())
        except Exception:
            quantity = 0
        if quantity > 5:
            status = "متوفر"
        elif 0 < quantity <= 5:
            status = "الكمية قليلة جدا"
        else:
            status = "غير متوفر"
        results.append(
            f"اسم المنتج: {row['اسم المنتج']}\n"
            f"الباركود: {row['الباركود']}\n"
            f"وحدة التخزين: {row['وحدة التخزين']}\n"
            f"الكمية: {row['الكمية']}\n"
            f"الحالة: {status}"
        )
    return "\n\n".join(results)

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(
        message,
        "مرحباً! أرسل اسم المنتج أو الباركود للبحث.\n"
        "للتحديث أرسل رمز الأدمن."
        f"\n{PROGRAMMER}"
    )

@bot.message_handler(content_types=['text'])
def handle_message(message):
    text = message.text.strip()
    chat_id = message.chat.id

    if text == ADMIN_CODE:
        admins.add(chat_id)
        bot.reply_to(message, "أرسل ملف الإكسل الجديد لتحديث البيانات (XLSX فقط).")
        return

    result = search_product(text)
    if result:
        bot.reply_to(message, result)
    else:
        bot.reply_to(message, "لم يتم العثور على المنتج.\n" + PROGRAMMER)

@bot.message_handler(content_types=['document'])
def handle_file(message):
    chat_id = message.chat.id
    if chat_id not in admins:
        bot.reply_to(message, "أنت لست أدمن. أرسل رمز الأدمن أولاً.")
        return

    document = message.document
    if not document.file_name.endswith('.xlsx'):
        bot.reply_to(message, "الرجاء إرسال ملف إكسل بصيغة XLSX فقط.")
        return

    try:
        file_info = bot.get_file(document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        with open(EXCEL_FILE, 'wb') as new_file:
            new_file.write(downloaded_file)
        bot.reply_to(message, "تم تحديث ملف الإكسل بنجاح!\n" + PROGRAMMER)
    except Exception as e:
        bot.reply_to(message, f"حدث خطأ أثناء رفع الملف: {e}")

if __name__ == "__main__":
    print("Bot running...")
    bot.infinity_polling()