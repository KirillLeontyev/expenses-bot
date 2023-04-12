from aiogram import Bot, Dispatcher, executor, types
# from aiogram.types import ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
import time as t
import os
import re
import sys
from functions import *
from settings import *

API_TOKEN = SEC_TOKEN
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)
bot_history = History()
VERSION = "0.1.3a" # parse_version()


@dp.message_handler(commands=["start"])
async def send_welcome(message: types.Message):
    # Обработка команды старт

    userid = message.from_user.id
    if userid not in AUTHORIZED:
        output = RESPONSE_DONTKNOWYOU
        await message.answer(output)
        return False

    kb_buttons = [
        [
            types.KeyboardButton(text=COMMAND_COUNT_EXPENCES.capitalize()),
            types.KeyboardButton(text=COMMAND_COUNT_EXPENCES_BY_USER.capitalize())
        ],
        [
            types.KeyboardButton(text=COMMAND_HISTORY.capitalize()),
            types.KeyboardButton(text=COMMAND_HISTORY_BY_DATE.capitalize())
        ],
        [

            types.KeyboardButton(text=COMMAND_EXPORT_TO_EXCEL),
            types.KeyboardButton(text=COMMAND_EXPORT_TO_EXCEL_ALL)
        ],
        [

            types.KeyboardButton(text=COMMAND_CLEAR_LAST.capitalize())
        ]
    ]
    bot_keyboard = types.ReplyKeyboardMarkup(keyboard=kb_buttons)
    await message.reply(RESPONSE_WELCOME + f" Версия {VERSION}", reply_markup=bot_keyboard)


@dp.message_handler()
async def echo(message: types.Message):
    # Обработка входящих сообщений

    msg = message.text
    date = message.forward_date if message.forward_date else message.date
    userid = message.from_user.id
    first_name = message.from_user.first_name if message.from_user.first_name else ""
    last_name = message.from_user.last_name if message.from_user.last_name else ""
    username = message.forward_sender_name if message.forward_sender_name else str(first_name + " " + last_name)
    output = False

    userid = message.from_user.id
    if userid not in AUTHORIZED:
        output = RESPONSE_DONTKNOWYOU
        await message.answer(output)
        return False


    if msg.lower() == COMMAND_HISTORY or msg.lower() == COMMAND_HISTORY_BY_DATE:
        # Вывод истории

        output = []
        history = bot_history.get()
        if history[0] == RESPONSE_HISTORY_EMPTY:
            output = RESPONSE_HISTORY_EMPTY
        else:
            if msg.lower() == COMMAND_HISTORY:
                for _ in history:
                    bold_in = ""
                    bold_out = ""
                    if _[1] > 0:
                        bold_in = "<b>"
                        bold_out = "</b>"
                    output.append(f"{bold_in}<i>{_[2][2]:0>2}.{_[2][1]:0>2}.{_[2][0]:0>4}</i>: {_[0]}: {_[1]:+,.2f}"
                                  f"{bold_out} р.")
            if msg.lower() == COMMAND_HISTORY_BY_DATE:
                history = bot_history.get_by_date()
                for date, value in sorted(history.items()):
                    bold_in = ""
                    bold_out = ""
                    if value > 0:
                        bold_in = "<b>"
                        bold_out = "</b>"
                    output.append(f"{bold_in}<i>{date}</i>: {value:+,.2f}{bold_out} р.")


    elif msg.lower() == COMMAND_CLEAR:
        # Очистка истории

        output = bot_history.clear()

    elif msg.lower() == COMMAND_CLEAR_LAST:
        # Очистка последней записи

        output = bot_history.clear_last()

    elif msg.lower() == COMMAND_COUNT_EXPENCES or msg.lower() == COMMAND_COUNT_EXPENCES_BY_USER:
        # Подсчет

        output = []
        output.append(f"<u><b>{months(current_month)}, {current_year}</b></u>:")
        expences_dict = bot_history.count_detailed( )

        if msg.lower() == COMMAND_COUNT_EXPENCES_BY_USER:
            expences_dict = bot_history.count_by_users()

        expences_total = bot_history.count_total()[0]
        income_total = bot_history.count_total()[1]
        for category, value in sorted(expences_dict.items()):
            output.append(f"{category}: {value:+,.2f} р.")
        # output.append("-"*40)
        output.append(f"\n<u><b>ИТОГО</b></u>: \nДоходы: {income_total:+,.2f}\nРасходы: {expences_total:+,.2f}\n"
                      f"Разница: {income_total + expences_total:+,.2f}")

    elif msg.lower() == COMMAND_EXPORT_TO_EXCEL.lower():
        output = False
        filename = export_to_excel(bot_history)
        await message.reply_document(open(filename, 'rb'))

    elif msg.lower() == COMMAND_EXPORT_TO_EXCEL_ALL.lower():
        output = False
        filename = export_to_excel(bot_history, 1)
        await message.reply_document(open(filename, 'rb'))

    else:
        # Обработка иных сообщений

        # log(f"[MESSAGE]: {msg} | {date.year}.{date.month} | {userid}: {username}")
        output = bot_history.add_entry(msg, date, [userid, username])

    if output:
        # Вывод ответа (если есть)

        if isinstance(output, list):
            output = f"\n".join(output)

        await message.answer(output, parse_mode=types.ParseMode.HTML)




if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)

