from datetime import datetime
from random import randint

from main.g_spreadsheets import add_spreadsheet_data


def pars_data():
    spreadsheet_id = '1kbZ3fwwm8Sv4QnEFcwJV7wp2AG3Nbq6tyQ6baGqyX60'
    header = ['Дата', 'Текст']
    date = datetime.now().strftime("%d.%m.%Y, %H:%M:%S")
    rows = [[date, f'{randint(0, 1000000)}']]
    add_spreadsheet_data(rows, spreadsheet_id, header)
