import pandas as pd

# завантажуємо книгу
path_file = 'rozstanovka.xlsx'
xlsx = pd.ExcelFile(path_file)
sheet_name = '01.07.2023'
df = xlsx.parse(sheet_name)

# отримання імені перешого підрозділу
name_first_force = input("Enter name first force:")

# Пошук номера рядку за іменем підрозділу
for row in range(0, 10):
    cell_value = df.iloc[row, 0]
    if cell_value == name_first_force:
        first_row = row

# отримання останнього номеру рядку
last_row = df.index[-1]

# отримання усіх імен підрозділів
for name_force in range(first_row, last_row):
    print(f"{df.iloc[name_force, 0]}")

