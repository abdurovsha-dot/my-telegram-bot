import time
import schedule
import telebot

TOKEN = "1601071734:AAFG7H_l2vTbsi4SZ8TsnCI8Y_mKSibWdOc"
bot = telebot.TeleBot(TOKEN)

CHAT_ID = -1003719999075


def send_message():
  caption_text = (
      "#muhim\n\n"
      '❗️ <i>"<b>Mahalla yettiligi</b>" xodimlari "<b>Raqamli mahalla'
      ' xodim</b>" mobil ilovasi orqali davomatdan o\'tishni'
      " unutmang.</i>\n\n"
      '📱 <a href="https://t.me/c/3271562585/1290">Mahalla yettiligi uchun mobil'
      " ilova</a>\n\n"
      '⭐️ Raqamli "MAHALLA" platformasi bo\'yicha barcha yangiliklardan xabardor'
      " bo'lish uchun kanalga obuna bo'ling!"
  )

  photo_path = r"C:\Users\USER\OneDrive\Рабочий стол\image.jpg.jpg"

  try:
    with open(photo_path, "rb") as photo:
      bot.send_photo(CHAT_ID, photo, caption=caption_text, parse_mode="HTML")
    print("Xabar va rasm muvaffaqiyatli yuborildi!")
  except Exception as e:
    print(f"Xatolik yuz berdi: {e}")


schedule.every().day.at("13:49").do(send_message)

print("Bot ishga tushdi va belgilangan vaqtni kutyapti...")

while True:
  schedule.run_pending()
  time.sleep(1)