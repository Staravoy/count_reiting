import tkinter as tk
from tkinter import filedialog

# Функція для обробки натискання кнопки "Вибрати документ Excel для аналізу"
def select_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    # Додайте ваш код для аналізу документу Excel

# Функція для обробки натискання кнопки "Закрити програму"
def close_program():
    root.destroy()

# Створення вікна
root = tk.Tk()
root.title("ПідраХуй")
root.geometry("400x400") # Розмір вікна

# Header
header_label = tk.Label(root, text="ПідраХуй", font=("Helvetica", 24, "bold"), fg="yellow")
header_label.pack(pady=20)

# Body
rules_label = tk.Label(root, text="Правила користування програмою:", font=("Helvetica", 14))
rules_label.pack()

rules_list = tk.Label(root, text="- Виберіть документ Excel для аналізу\n- Закрийте програму", font=("Helvetica", 12))
rules_list.pack()

select_button = tk.Button(root, text="Вибрати документ Excel для аналізу", command=select_excel)
select_button.pack(pady=10)

close_button = tk.Button(root, text="Закрити програму", command=close_program)
close_button.pack()

# Footer
footer_label = tk.Label(root, text="Company: Staravoy", font=("Helvetica", 10))
footer_label.pack(pady=20)

root.mainloop()
