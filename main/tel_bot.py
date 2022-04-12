import os

import aiofiles
from aiocsv import AsyncWriter
from aiogram.dispatcher.filters import Text
from aiogram.utils import executor
from aiogram import Bot, Dispatcher, types

bot = Bot(token=os.environ.get('TG_BOT_TOKEN'))
dp = Dispatcher(bot)
executor.start_polling(dp)


@dp.message_handler(commands='start')
async def start(message: types.Message):
    await message.answer('Hello!')
    start_buttons = ['Start', 'Get CSV', 'TRU la la']
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(*start_buttons)
    await message.answer('Select button', reply_markup=keyboard)


@dp.message_handler(Text(equals='Get CSV'))
async def get_csv(message: types.Message):
    await message.answer('Wait a bit!')
    chat_id = message.chat.id
    await send_data(chat_id=chat_id)


async def write_csv():
    path = os.path.join(os.getcwd(), 'send_data.csv')
    async with aiofiles.open(path, 'w') as file:
        writer = AsyncWriter(file)
        await writer.writerow(('заголовок1', 'заголовок2', 'заголовок3',))
        await writer.writerow(('11', '564', '766',))
        await writer.writerow(('76', '98', '987655',))
    return path


async def send_data(chat_id):
    file_path = await write_csv()
    await bot.send_document(chat_id=chat_id, document=open(file_path, 'rb'))
