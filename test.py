import openpyxl

file = 'rozstanovka.xlsx'
wb = openpyxl.load_workbook(file, data_only=True)
all_sheets = wb.sheetnames
name_sheets_for_table = []
for x in all_sheets:
    name_sheets_for_table.append(x[:-5])
name_sheets_for_table = [' '] + name_sheets_for_table

new_table = openpyxl.Workbook()
new_table.save('new_table.xlsx')

def fill_table(num_row, my_list):
    # Створюємо новий документ Excel
    wb = openpyxl.load_workbook('new_table.xlsx')
    # Вибираємо активний аркуш
    sheet = wb.active
    col = 1
    for x in my_list:
        sheet.cell(row=num_row, column=col, value=x)
        col += 1
    wb.save('new_table.xlsx')

fill_table(1, name_sheets_for_table)

second_row = ['value1', 'value2', 'value3', 'value4', 'value5', 'value6', 'value7']

therd_row = ['value1', 'value2', 'value3', 'value4', 'value5', 'value6', 'value7']

fill_table(2, second_row)
fill_table(3, therd_row)
