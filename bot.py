import os
import logging
import telebot
from dotenv import load_dotenv


logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, filename="logs.log",
                    format="%(asctime)s::%(levelname)s::%(message)s")

load_dotenv()

if not os.getenv("BOT_TOKEN", None):
    print("Токен бота не указан")
    logger.error("Токен бота не указан")
    exit()
if not os.getenv("USER_IDS", None):
    print("Не указаны id пользователей")
    logger.error("Не указаны id пользователей")
    exit()


USER_IDS = os.getenv("USER_IDS").split(",")
BOT_TOKEN = os.getenv("BOT_TOKEN")


def remove_report_file(filename: str) -> None:
    try:
        os.remove(filename)
        logger.info(f"File {filename} deleted")
    except Exception as err:
        print("Не получилось удалить файл", err)
        logger.error(err)


def send_report_to_tg(filename: str) -> None:
    bot = telebot.TeleBot(token=BOT_TOKEN)

    document = open(filename, 'rb')

    for user_id in USER_IDS:
        try:
            bot.send_document(user_id, document)
            logger.info(f"File {filename} sent to user {user_id}")
        except telebot.apihelper.ApiTelegramException as err:
            print("Не получилось отправить файл", err)
            logger.error(err)

    remove_report_file(filename)

send_report_to_tg("students-predicts.xlsx")