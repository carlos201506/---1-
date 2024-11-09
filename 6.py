import pandas as pd

# Чтение данных из Excel-файла и создание DataFrame
df = pd.read_excel('C:/Users/Misha/Desktop/Pandas/2.xls')

# Удаление столбца "Unnamed: 0" из DataFrame
df.drop("Unnamed: 0", axis=1, inplace=True)

# Номер колонки, для которой нужно присвоить числовой индекс (нумерация начинается с 0)
column_number = 1

# Присвоение числового индекса на основе текста в указанной колонке
df = df.set_index(df.columns[column_number])

# Сортировка DataFrame по индексу
df_sorted = df.sort_values(by=df.index.name)

# Агрегирование строк с одинаковым индексом
df_aggregated = df_sorted.groupby(level=0).sum()

# Вычисление результата деления значений столбца "Unnamed: 5" на значения столбца "Unnamed: 3"
df_aggregated["Unnamed: 4"] = df_aggregated["Unnamed: 5"] / df_aggregated["Unnamed: 3"]
df_aggregated["Unnamed: 4"] = df_aggregated["Unnamed: 4"].fillna(0)

# Удаление столбца "Unnamed: 0" из DataFrame `df_aggregated`
df_aggregated.drop("Unnamed: 0", axis=1, inplace=True)

# Сохранение DataFrame с передвинутыми строками и новой колонкой в файл Excel
with pd.ExcelWriter('C:/Users/Misha/Desktop/Pandas/2_сдвинутый.xlsx', mode='w') as writer:
    df_sorted.to_excel(writer, sheet_name='Sheet1', index=True)

# Сохранение агрегированного DataFrame с удаленным столбцом в файл Excel
with pd.ExcelWriter('C:/Users/Misha/Desktop/Pandas/2_агрегированный.xlsx', mode='w') as writer:
    df_aggregated.to_excel(writer, sheet_name='Sheet1', index=True)

# Вывод DataFrame с передвинутыми строками
print(df_sorted)

# Вывод агрегированного DataFrame с удаленным столбцом
print(df_aggregated)









