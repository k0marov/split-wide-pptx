import os
import re
import tempfile

import aiogram
from aiogram import F

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import Message, KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup, \
    CallbackQuery
import asyncio

from dotenv import load_dotenv

from src.create_triptych import create_triptych
from src.datasources import admin_auth_db
from src.renote.processor import process_pptx, ProcessingOptions
from src.userbot import telethon_file_manager

from src import config
from src.libre import convert_to_pdf

USER_MODE_CREATE = 'create'
USER_MODE_CUT = 'cut'

class PPTXBot:
    def __init__(self, token: str, db: admin_auth_db.AdminAuthDB):
        self.bot = Bot(token=token)
        self.db = db
        self.dp = Dispatcher()
        self.user_modes = {}  # user_id: mode
        self.setup_handlers()

    def setup_handlers(self):
        self.dp.message(Command("start"))(self.start_handler)
        self.dp.message(Command("help"))(self.help_handler)
        self.dp.message(Command(re.compile("accept_[0-9]+")))(self.accept_handler)
        self.dp.message(Command(re.compile("reject_[0-9]+")))(self.reject_handler)
        self.dp.callback_query.register(
            self.callback_handler,
            lambda c: c.data in [USER_MODE_CREATE, USER_MODE_CUT]
        )
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
        reply_markup = None
        approved = self.db.check_is_approved(str(user_id))
        if not approved:
            admins = self.db.get_admins_list()
            for admin in admins:
                msg = ('Новая заявка на доступ:\n'
                    f'ID: {message.from_user.id}\n'
                    f'Username: @{message.from_user.username}\n'
                    f'Имя: {message.from_user.full_name}\n'
                    f'Используйте /accept_{message.from_user.id} для принятия или /reject_{message.from_user.id} для отклонения\n')
                await self.bot.send_message(admin.chat_id, msg)
            welcome_text.append('\nВы пока не приняты админом, он должен принять вас, инструкции были посланы ему в чат.')
        await message.answer(
            ''.join(welcome_text),
            reply_markup=self.get_inline_kb() if approved else None,
        )

    def get_inline_kb(self):
        kb_list = [
            [InlineKeyboardButton(text="Разрезать триптих", callback_data=USER_MODE_CUT),
             InlineKeyboardButton(text="Создать триптих", callback_data=USER_MODE_CREATE)],
        ]
        return InlineKeyboardMarkup(inline_keyboard=kb_list)

    async def callback_handler(self, callback: CallbackQuery):
        if callback.data == USER_MODE_CUT:
            await callback.message.answer("Вы выбрали режим разрезания триптихов", reply_markup=self.get_inline_kb())
            self.user_modes[callback.from_user.id] = USER_MODE_CUT
        elif callback.data == USER_MODE_CREATE:
            await callback.message.answer("Вы выбрали режим создания триптихов", reply_markup=self.get_inline_kb())
            self.user_modes[callback.from_user.id] = USER_MODE_CREATE
        await callback.answer()

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


    async def _send_split_presentation(self, path: str, message: Message):
        msg_file = await telethon_file_manager.upload_file_smart(self.bot, path)
        if msg_file is None:
            # this means that telethon big file upload was not needed and we can use bot api
            msg_file = types.FSInputFile(path)
        await message.answer_document(
            msg_file,
            caption="✅ Ваша презентация обработана!"
        )

    async def _convert_to_pdf(self, path: str, message: Message):
        await message.answer("⏳ Конвертация в PDF займёт около 1 мин...")
        pdf_path = path.replace('.pptx', '.pdf')
        if not await convert_to_pdf.convert_pptx_to_pdf(path, pdf_path):
            await message.answer("Ошибка при конвертации в PDF")
        else:
            msg_file = await telethon_file_manager.upload_file_smart(self.bot, pdf_path)
            if msg_file is None:
                # this means that telethon big file upload was not needed and we can use bot api
                msg_file = types.FSInputFile(pdf_path)
            await message.answer_document(
                msg_file,
                caption="✅ Успешно сконвертировал разрезанную презентацию в PDF"
            )


    async def handle_document(self, message: Message):
        if self.user_modes.get(message.from_user.id, USER_MODE_CUT) == USER_MODE_CUT:
            return await self.handle_document_cut_pptx(message)
        else:
            return await self.handle_document_create(message)

    async def handle_document_create(self, message: Message):
        if not message.document.file_name or not message.document.file_name.endswith('.pdf'):
            await message.answer("❌ Пожалуйста, отправьте файл в формате .pdf")
            return
        await message.answer("⏳ Создаю презентацию-триптих...")
        try:
            input_path = await telethon_file_manager.download_file_smart(self.bot, message.document, message)
            try:
                with tempfile.NamedTemporaryFile(
                        prefix='triptych_' + ''.join(message.document.file_name.split('.')[:-1]),
                        suffix='.pdf',
                        delete=False) as output_file:
                    output_path = output_file.name
                    create_triptych(input_path, output_path)
                    msg_file = await telethon_file_manager.upload_file_smart(self.bot, output_path)
                    if msg_file is None:
                        # this means that telethon big file upload was not needed and we can use bot api
                        msg_file = types.FSInputFile(output_path)
                    await message.answer_document(
                        msg_file,
                        caption="✅ Успешно сконвертировал разрезанную презентацию в PDF"
                    )
            finally:
                if input_path is not None and os.path.exists(input_path):
                    os.unlink(input_path)
        except Exception as e:
            await message.answer(f"❌ Произошла ошибка при обработке: {str(e)}")

    async def handle_document_cut_pptx(self, message: Message):
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

                await asyncio.gather(
                    self._send_split_presentation(output_path, message),
                    self._convert_to_pdf(output_path, message)
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