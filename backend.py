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


def copy_sheet_from_template(wb, template_name, sheet_name):
    source = wb[template_name]
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    else:
        ws = wb.copy_worksheet(source)
        ws.title = sheet_name
        return ws


def load_workbook(name: str):
    path = os.path.join(os.getcwd(), FOLDER, f"{name}.xlsx")
    if not os.path.exists(path):
        wb = openpyxl.load_workbook('template_full.xlsx')
        wb.save(path)
        wb.close()
    return openpyxl.load_workbook(path)


def save_workbook(wb, name: str):
    wb.save(os.path.join(os.getcwd(), FOLDER, f"{name}.xlsx"))
    wb.close()


def read_workbook(name: str):
    with open(os.path.join(os.getcwd(), FOLDER, f"{name}.xlsx"), "rb") as template_file:
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
