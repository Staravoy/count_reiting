import pandas as pd
from generate_name_colum import generate_name_column

file_path = 'rozstanovka.xlsx'
df = pd.ExcelFile(file_path)
data = df.parse('01.07.2023')
