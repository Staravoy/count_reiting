import tkinter as tk
from tkinter import filedialog

# Створення головного вікна
root = tk.Tk()
root.withdraw()

# Відкриваємо діалогове вікно для вибору файлу
file_path = filedialog.askopenfilename(title="Оберіть файл для обробки", filetypes=[("All Files", "*.*")])

# Виведення шляху до вибраного файлу
print("Обраний файл:", file_path)
