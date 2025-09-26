import os
import re
import tempfile

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import Message
import asyncio

from dotenv import load_dotenv

from src.datasources import admin_auth_db
from src.renote.processor import process_pptx, ProcessingOptions
from src.userbot import telethon_file_manager

from src import config

class PPTXBot:
    def __init__(self, token: str, db: admin_auth_db.AdminAuthDB):
        self.bot = Bot(token=token)
        self.db = db
        self.dp = Dispatcher()
        self.setup_handlers()

    def setup_handlers(self):
        self.dp.message(Command("start"))(self.start_handler)
        self.dp.message(Command("help"))(self.help_handler)
        self.dp.message(Command(re.compile("accept_[0-9]+")))(self.accept_handler)
        self.dp.message(Command(re.compile("reject_[0-9]+")))(self.reject_handler)
        self.dp.message()(self.message_handler)

    async def start_handler(self, message: Message):
        """Handle /start command"""
        user_id = message.from_user.id
        if str(user_id) == config.SUPERADMIN_TELEGRAM_ID:
            self.db.set_admin(str(user_id), str(message.chat.id))
            await message.answer("Вы суперадмин.")
            return

        welcome_text = [
            "👋 Добро пожаловать в Renote PPTX transformer!\n\n"
            "Я помогу вам преобразовать вашу презентацию.\n"
            "Просто отправьте мне файл презентации (.pptx) и я обработаю его.\n\n"
            "Используйте /help для справки."
        ]
        if not self.db.check_is_approved(str(user_id)):
            admin_info = self.db.get_is_admin_info(config.SUPERADMIN_TELEGRAM_ID)
            if admin_info is not None:
                await self.bot.send_message(admin_info.chat_id,
                                            f'Пользователь @{message.from_user.username or "<нет юзернейма>"} хочет присоединиться\n'
                                            f'/accept_{user_id}\n/reject_{user_id}'
                )
                welcome_text.append('\nВы пока не приняты админом, он должен принять вас, инструкции были посланы ему в чат.')
            else:
                welcome_text.append(f'\nВы пока не приняты админом, и админ ещё не заходил в бота, поэтому ему не было прислано сообщения о Вас.\nЧтобы принять вас, ему нужно будет ввести /accept_{user_id}')

        await message.answer(
            ''.join(welcome_text),
        )

    async def accept_handler(self, message: Message):
        user_id = message.from_user.id
        if self.db.get_is_admin_info(str(user_id)) is None:
            await message.answer("Вы не админ.")
            return
        accept_user_id = message.text.split('_')[-1]
        self.db.set_approved(accept_user_id)
        await message.answer(f'Пользователь {accept_user_id} принят.')

    async def reject_handler(self, message: Message):
        user_id = message.from_user.id
        if self.db.get_is_admin_info(str(user_id)) is None:
            await message.answer("Вы не админ.")
            return
        await message.answer("Пользователь был отвергнут.")

    async def help_handler(self, message: Message):
        """Handle /help command"""
        help_text = (
            "ℹ️ **Справка по использованию бота:**\n\n"
            f"0. Админ должен аппрувнуть вас, написав в своём чате /accept_{message.from_user.id} \n"
            "1. Отправьте мне файл презентации в формате .pptx\n"
            "2. Я обработаю его и верну преобразованную версию\n"
            "3. Файл будет сохранен с тем же именем, но с приставкой '_renote'\n\n"
        )
        await message.answer(help_text)

    async def message_handler(self, message: Message):
        if message.document and str(message.from_user.id) == config.TELETHON_ADMIN_ID:
            print("Ignoring document message from telethon admin")
            return

        if not self.db.check_is_approved(str(message.from_user.id)):
            await message.answer("Вы пока не приняты админом.")
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
    token = config.TELEGRAM_BOT_TOKEN
    if not token:
        print("Error: TELEGRAM_BOT_TOKEN environment variable is not set")
        return

    db = admin_auth_db.AdminAuthDB(config.SQLITE_URL, create_tables=True)
    db.set_approved(config.SUPERADMIN_TELEGRAM_ID)

    bot = PPTXBot(token, db)
    await telethon_file_manager.telethon_downloader.init_client()
    await bot.run()


if __name__ == "__main__":
    asyncio.run(main())