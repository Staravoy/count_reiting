import pandas as pd

df = pd.read_excel('rozstanovka.xlsx', sheet_name='01.07.2023')

valueB7 = df.iloc[2, 5]


print(valueB7)
