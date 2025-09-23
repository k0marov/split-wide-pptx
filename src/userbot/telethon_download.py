import os
import tempfile
from telethon import TelegramClient
from dotenv import load_dotenv

from src.config import TELEGRAM_API_ID, TELEGRAM_API_HASH, TELETHON_ADMIN_ID, TELEGRAM_BOT_USERNAME

class TelethonDownloader:
    """Класс для скачивания больших файлов через Telethon"""
    
    def __init__(self):
        self.client = None
        self.session_file = "data/telethon_session"
        
    async def _get_client(self):
        """Получает авторизованный Telethon клиент"""
        if self.client is None:
            if not TELEGRAM_API_ID or not TELEGRAM_API_HASH:
                raise Exception("Не настроены TELEGRAM_API_ID и TELEGRAM_API_HASH")
            
            # Создаем директорию для сессии
            os.makedirs(os.path.dirname(self.session_file), exist_ok=True)
            
            self.client = TelegramClient(self.session_file, TELEGRAM_API_ID, TELEGRAM_API_HASH)
            await self.client.start()
            
            if not await self.client.is_user_authorized():
                print("Telethon клиент не авторизован! Требуется первоначальная настройка.")
                raise Exception("Telethon клиент не авторизован")
        
        return self.client
    
    async def download_large_file(self, bot_chat: str, expected_size: int, filename: str = None) -> str:
        """
        Скачивает большой файл через Telethon, находя его среди последних сообщений по размеру.
        
        :param bot_chat: Username бота (без @) для поиска чата
        :param expected_size: Ожидаемый размер файла в байтах (по нему ищем сообщение)
        :param filename: Имя файла (опционально)
        :return: Путь к скачанному файлу
        """
        try:
            print(f"Скачиваем файл {filename} из чата @{bot_chat} с размером {expected_size} байт")

            client = await self._get_client()
            
            # Готовим папку downloads в корне репозитория
            repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            downloads_dir = os.path.join(repo_root, 'downloads')
            os.makedirs(downloads_dir, exist_ok=True)

            found = False
            file_path = None

            # Проходим последние 50 сообщений и ищем пересланный документ нужного размера
            async for message in client.iter_messages(bot_chat, limit=50):
                if not message:
                    continue
                if not message.media or not hasattr(message.media, 'document'):
                    continue
                try:
                    file_size = message.media.document.size
                except Exception:
                    file_size = None

                if file_size == expected_size:
                    # Имя файла
                    detected_name = None
                    try:
                        for attr in getattr(message.media.document, 'attributes', []) or []:
                            if hasattr(attr, 'file_name'):
                                detected_name = attr.file_name
                                break
                    except Exception:
                        detected_name = None

                    if not filename:
                        filename = detected_name or f"file_{message.id}.mp4"

                    file_path = os.path.join(downloads_dir, filename)
                    print(f"Найден файл с размером {file_size} байт: {filename}")
                    # Скачиваем
                    file_path = await message.download_media(file=file_path)
                    print(f"Файл успешно сохранен: {file_path}")
                    found = True
                    break

            if not found:
                raise Exception("Не удалось найти пересланный медиафайл среди последних сообщений по размеру")

            # Отключаемся после скачивания
            await self._disconnect_after_download()

            return file_path
            
        except Exception as e:
            error_msg = str(e).lower()
            
            # Обрабатываем database is locked
            if "database is locked" in error_msg:
                print("Ошибка database is locked, переподключаюсь к сессии...")
                await self._reconnect()
                # Повторяем попытку скачивания
                return await self.download_large_file(bot_chat, expected_size, filename)
            
            print(f"Ошибка скачивания через Telethon: {e}")
            raise
    
    async def close(self):
        """Закрывает соединение"""
        if self.client:
            await self.client.disconnect()
            self.client = None
    
    async def _reconnect(self):
        """Переподключается к сессии при ошибках"""
        try:
            if self.client:
                await self.client.disconnect()
            self.client = None
            print("Переподключение к Telethon сессии...")
            await self._get_client()
            print("Переподключение успешно")
        except Exception as e:
            print(f"Ошибка переподключения: {e}")
            raise
    
    async def _disconnect_after_download(self):
        """Отключается от сессии после скачивания"""
        try:
            if self.client:
                await self.client.disconnect()
                self.client = None
                print("Отключился от Telethon сессии после скачивания")
        except Exception as e:
            print(f"Ошибка отключения: {e}")
            # Не выбрасываем ошибку, так как файл уже скачан


# Глобальный экземпляр
telethon_downloader = TelethonDownloader()


def is_file_too_large(file_size: int) -> bool:
    """
    Проверяет, больше ли файл лимита Bot API (20MB)
    
    :param file_size: Размер файла в байтах
    :return: True если файл больше 20MB
    """
    return file_size > 20 * 1024 * 1024  # 20MB в байтах


async def get_telethon_admin_id() -> int:
    """
    Получает ID администратора для пересылки больших файлов
    
    :return: ID администратора
    :raises: Exception если ID не настроен
    """
    if not TELETHON_ADMIN_ID:
        raise Exception("Не настроен TELETHON_ADMIN_ID для пересылки больших файлов")
    
    try:
        return int(TELETHON_ADMIN_ID)
    except ValueError:
        raise Exception(f"Некорректный TELETHON_ADMIN_ID: {TELETHON_ADMIN_ID}")


async def download_file_smart(bot, file_obj, message) -> str:
    """
    Умно скачивает файл: через Bot API или Telethon в зависимости от размера
    
    :param bot: Экземпляр telebot
    :param file_obj: Объект файла из сообщения
    :param message: Сообщение с файлом
    :return: Путь к скачанному файлу
    """
    
    # Проверяем размер файла
    file_size = file_obj.file_size
    
    print(f"Размер файла: {file_size / 1024 / 1024:.1f} MB")
    
    if is_file_too_large(file_size):
        print("Файл больше 20MB, использую Telethon")
        
        # Получаем ID администратора для пересылки
        admin_id = await get_telethon_admin_id()
        
        # Пересылаем файл администратору
        forwarded_message = await bot.forward_message(admin_id, message.chat.id, message.message_id)
        print(f"Файл переслан администратору {admin_id}, message_id: {forwarded_message.message_id}")
        
        # Скачиваем через Telethon: ищем файл в чате с ботом по размеру
        if not TELEGRAM_BOT_USERNAME:
            raise Exception("Не настроен TELEGRAM_BOT_USERNAME для поиска файлов")
        
        filename = getattr(file_obj, 'file_name', f"video_{file_obj.file_id}.mp4")
        expected_size = getattr(file_obj, 'file_size', None)
        return await telethon_downloader.download_large_file(
            TELEGRAM_BOT_USERNAME,
            expected_size,
            filename
        )
    else:
        print("Файл меньше 20MB, использую Bot API")
        
        # Скачиваем через обычный Bot API
        file_info = bot.get_file(file_obj.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Сохраняем во временный файл
        temp_dir = tempfile.gettempdir()
        video_filename = f"video_{file_obj.file_id}.mp4"
        video_path = os.path.join(temp_dir, video_filename)
        
        with open(video_path, 'wb') as video_file:
            video_file.write(downloaded_file)
        
        return video_path


def setup_telethon_instructions():
    """Возвращает инструкции по настройке Telethon"""
    return """
📋 Настройка Telethon для скачивания больших файлов:

1. Получите API_ID и API_HASH на https://my.telegram.org/auth
2. Узнайте ваш Telegram ID (напишите @userinfobot)
3. Узнайте username вашего бота (без @)
4. Добавьте в .env файл:
   TELEGRAM_API_ID=ваш_api_id
   TELEGRAM_API_HASH=ваш_api_hash
   TELETHON_ADMIN_ID=ваш_telegram_id
   TELEGRAM_BOT_USERNAME=username_бота_без_собаки

5. Запустите бота для авторизации Telethon клиента
6. Введите номер телефона и код подтверждения

После настройки бот будет пересылать большие файлы вам и искать их в чате с ботом!
"""
