import telebot
import os
from docx2pdf import convert

# ←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←
TOKEN = '8606785863:AAGgtfEEn1i8Af3aD2UxpksiQ7Wb0VgivUs'  # ← Замени!
# ←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←←

bot = telebot.TeleBot(TOKEN)

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id,
                     "👋 Привет!\n\n"
                     "Отправь .docx файл — конвертирую в PDF через Microsoft Word (идеальное качество).")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    doc = message.document
    if not doc.file_name.lower().endswith('.docx'):
        bot.reply_to(message, "❌ Только .docx файлы.")
        return

    file_info = bot.get_file(doc.file_id)
    downloaded = bot.download_file(file_info.file_path)

    input_path = f"temp_{doc.file_id}.docx"
    pdf_path = f"temp_{doc.file_id}.pdf"

    with open(input_path, 'wb') as f:
        f.write(downloaded)

    bot.reply_to(message, "🔄 Конвертирую через Microsoft Word...")

    try:
        convert(input_path, pdf_path)   # ←←← настоящая конвертация через Word

        with open(pdf_path, 'rb') as pdf:
            bot.send_document(
                message.chat.id,
                pdf,
                caption="✅ Готово! Идеальный PDF (docx2pdf)."
            )

    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка:\n{str(e)}\n\nУбедись, что Microsoft Word установлен.")
    finally:
        for path in [input_path, pdf_path]:
            if os.path.exists(path):
                os.remove(path)

print("Бот запущен...")
bot.infinity_polling()
