import pandas as pd

# Чтение данных из Excel-файла и создание DataFrame

df = pd.read_excel('C:/Users/Misha/Desktop/Pandas/2_агрегированный.xlsx')

df['Доля'] = df['Unnamed: 5'] / df['Unnamed: 5'].sum()

df['Аккум.доля'] = df['Доля'].cumsum()

df["Категория"] = ' '

df.loc[df['Аккум.доля'] < 0.6, 'Категория'] = 'A'
df.loc[(df['Аккум.доля'] > 0.6) & (df['Аккум.доля'] < 0.75), 'Категория'] = 'B'
df.loc[df['Аккум.доля'] > 0.75, 'Категория'] = 'C'

# Сохранение результатов (необязательно)

df.to_excel('C:/Users/Misha/Desktop/Pandas/abc_xyz_analysis_sum_results.xlsx')
print(df)  # добавь в мой код