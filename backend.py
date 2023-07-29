import os

import openpyxl
from config import *
from datetime import datetime

xl_size_pos = {}
for size, letter in zip(sizes[1:], size_letters):
    xl_size_pos[size] = letter

FOLDER = str(datetime.today().year)
if FOLDER not in os.listdir(os.getcwd()):
    os.mkdir(FOLDER)


def get_path(name):
    return os.path.join(os.getcwd(), FOLDER, f"{name}.xlsx")


def copy_sheet_from_template(wb, template_name, sheet_name):
    source = wb[template_name]
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    else:
        ws = wb.copy_worksheet(source)
        ws.title = sheet_name
        return ws


def load_workbook(name: str):
    if not os.path.isfile(get_path(name)):
        wb = openpyxl.load_workbook('template_full.xlsx')
        wb.save(get_path(name))
        wb.close()
    return openpyxl.load_workbook(get_path(name))


def save_workbook(wb, name: str):
    wb.save(get_path(name))
    wb.close()


def read_workbook(name: str):
    if os.path.isfile(get_path(name)):
        with open(get_path(name), "rb") as template_file:
            return template_file.read()


def fill_xl_student(result: dict):
    wb = load_workbook(result['group'])
    ws = copy_sheet_from_template(wb, 'Шаблон', result['second_name'])
    for i, position in enumerate(result['positions'], start=4):
        ws[f"A{i}"] = position['title']
        ws[f"B{i}"] = position['price']

        if position['size'] == NOT_CHOSEN:
            pass
        elif position['size'] == 'NO SIZE':
            ws[f"{xl_size_pos[sizes[1]]}{i}"] = position['number']
        else:
            ws[f"{xl_size_pos[position['size']]}{i}"] = position['number']
        ws[f"N{i}"] = f"=SUM(C{i}:M{i})*B{i}"
    l = len(result['positions'])
    ws[f"N{l + 5}"] = f"=SUM(N{1}:N{l + 4})"
    ws[f"M{l + 5}"] = f"Итого: "

    save_workbook(wb, result['group'])
