import pandas as pd

# Чтение данных из Excel-файла и создание DataFrame
df = pd.read_excel('C:/Users/Misha/Desktop/Pandas/2.xls')

# Номер колонки, для которой нужно присвоить числовой индекс (нумерация начинается с 0)
column_number = 1

# Присвоение числового индекса на основе текста в указанной колонке
df = df.set_index(df.columns[column_number])

# Сортировка DataFrame по индексу
df_sorted = df.sort_values(by=df.index.name)

# Агрегирование строк с одинаковым индексом
df_aggregated = df_sorted.groupby(level=0).sum()

# Сохранение DataFrame с передвинутыми строками в файл Excel
df_sorted.to_excel('C:/Users/Misha/Desktop/Pandas/2_сдвинутый.xlsx')

# Сохранение агрегированного DataFrame в файл Excel
df_aggregated.to_excel('C:/Users/Misha/Desktop/Pandas/2_агрегированный.xlsx')

# Вывод DataFrame с передвинутыми строками
print(df_sorted)

# Вывод агрегированного DataFrame
print(df_aggregated)