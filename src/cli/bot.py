import os
import tempfile
from distutils.command.config import config
from typing import Optional

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import Message
from aiogram.utils.keyboard import ReplyKeyboardBuilder
import asyncio

from dotenv import load_dotenv

from src.renote.processor import process_pptx, ProcessingOptions
from src.userbot import telethon_file_manager

from src.config import TELETHON_ADMIN_ID

class PPTXBot:
    def __init__(self, token: str):
        self.bot = Bot(token=token)
        self.dp = Dispatcher()
        self.setup_handlers()

        # Store user states: {user_id: 'waiting_for_file'}
        self.user_states = {}

    def setup_handlers(self):
        self.dp.message(Command("start"))(self.start_handler)
        self.dp.message(Command("help"))(self.help_handler)
        self.dp.message()(self.message_handler)

    async def start_handler(self, message: Message):
        """Handle /start command"""
        welcome_text = (
            "👋 Добро пожаловать в Renote PPTX transformer!\n\n"
            "Я помогу вам преобразовать вашу презентацию.\n"
            "Просто отправьте мне файл презентации (.pptx) и я обработаю его.\n\n"
            "Используйте /help для справки."
        )

        await message.answer(
            welcome_text,
        )

    async def help_handler(self, message: Message):
        """Handle /help command"""
        help_text = (
            "ℹ️ **Справка по использованию бота:**\n\n"
            "1. Отправьте мне файл презентации в формате .pptx\n"
            "2. Я обработаю его и верну преобразованную версию\n"
            "3. Файл будет сохранен с тем же именем, но с приставкой '_renote'\n\n"
        )
        await message.answer(help_text)

    async def message_handler(self, message: Message):
        if str(message.from_user.id) == TELETHON_ADMIN_ID:
            print("Ignoring message from telethon admin")
            return
        """Handle regular messages"""
        if message.document:
            await self.handle_document(message)
        else:
            await message.answer("Пожалуйста, отправьте файл презентации или используйте /help для справки")

    async def handle_document(self, message: Message):
        """Handle document upload"""
        if not message.document.file_name or not message.document.file_name.endswith('.pptx'):
            await message.answer("❌ Пожалуйста, отправьте файл в формате .pptx")
            return

        await message.answer("⏳ Обрабатываю презентацию...")

        try:
            with tempfile.NamedTemporaryFile(
                    prefix='splitted_' + ''.join(message.document.file_name.split('.')[:-1]),
                    suffix='.pptx',
                    delete=False) as output_file:
                output_path = output_file.name

            input_path = None

            try:
                # Download the file
                input_path = await telethon_file_manager.download_file_smart(self.bot, message.document, message)
                # Process the presentation
                opts = ProcessingOptions(
                    title_min_font_pt=36.0,  # Default values
                    title_min_width_ratio=1.2,  # Default values
                )

                process_pptx(
                    input_path,
                    output_path,
                    options=opts,
                    direct=True,
                )

                msg_file = await telethon_file_manager.upload_file_smart(self.bot, output_path)
                print('got msg_file =', msg_file)
                if msg_file is None:
                    # this means that telethon big file upload was not needed and we can use bot api
                    msg_file = types.InputFile(output_path)
                await message.answer_document(
                    msg_file,
                    caption="✅ Ваша презентация обработана!"
                )
            finally:
                # Clean up temporary files
                if input_path is not None and os.path.exists(input_path):
                    os.unlink(input_path)

                if os.path.exists(output_path):
                    os.unlink(output_path)

        except Exception as e:
            await message.answer(f"❌ Произошла ошибка при обработке: {str(e)}")

    async def run(self):
        """Start the bot"""
        await self.dp.start_polling(self.bot)


async def main():
    # Get bot token from environment variable
    load_dotenv()
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        print("Error: TELEGRAM_BOT_TOKEN environment variable is not set")
        return

    bot = PPTXBot(token)
    await bot.run()


if __name__ == "__main__":
    asyncio.run(main())