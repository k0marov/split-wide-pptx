import os

from dotenv import load_dotenv

load_dotenv()

# Переменные для Telethon (для скачивания больших файлов)
TELEGRAM_API_ID = os.getenv("TELEGRAM_API_ID")
TELEGRAM_API_HASH = os.getenv("TELEGRAM_API_HASH")
TELETHON_ADMIN_ID = os.getenv("TELETHON_ADMIN_ID")  # ID пользователя для пересылки больших файлов
TELEGRAM_BOT_USERNAME = os.getenv("TELEGRAM_BOT_USERNAME")  # Username бота без @
