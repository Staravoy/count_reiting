import openpyxl
import tkinter as tk

from tkinter import filedialog


#
def browse_file():
    filepath = filedialog.askopenfilename()
    return filepath


# завантажуємо файл
rozstanovka = openpyxl.load_workbook(browse_file())
# Отримуємо усі назви робочих листів
sheet_names = rozstanovka.sheetnames
letter_column = input('Введіть літеру стовбчика:')


def sos(sheet_erorr):
    print(f"{sheet_erorr} Помилка даних!")


x = 1
for i in range(7, 27):
    forces = f"A{i}"
    call = f"{letter_column}{i}"

    total = 0
    for one_sheet in sheet_names:
        sheet = rozstanovka[one_sheet]
        forces_names = sheet[forces].value
        data = sheet[call].value
        if data is None:
            data = 0
        elif type(data) == str:
            sos(sheet)
            data = 0
        total += data
    print(f"{x}) {forces_names}: {call} - {total}")
    x += 1

rozstanovka.close()
