import os
import openpyxl
from generate_name_colum import generate_name_column
import tkinter as tk
from tkinter import filedialog

def get_abs_path(file_name):
    # Отримати абсолютний шлях до файлу
    current_dir = os.getcwd()
    abs_path = os.path.join(current_dir, file_name)
    return abs_path

# Решта вашого коду без змін

# Замість викликів `openpyxl.load_workbook(file_path)` змініть на `openpyxl.load_workbook(get_abs_path(file_path))`

# Замість викликів `new_table.save('new_table.xlsx')` змініть на `new_table.save(get_abs_path('new_table.xlsx'))`

# Замість викликів `wb.save(name)` змініть на `wb.save(get_abs_path(name))`

# Інші місця, де зберігається або зчитується файл, також використовуйте `get_abs_path()` для правильного шляху.
