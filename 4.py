import pandas as pd

# Загрузка данных из Excel-файла
df = pd.read_excel('C:/Users/Misha/Desktop/Pandas/2_агрегированный.xlsx')

# Переименование столбцов
df.rename(columns={'Unnamed: 3': 'Количество', 'Unnamed: 4': 'Цена за шт', 'Unnamed: 5': 'Сумма'}, inplace=True)

# АБЦ-ХЮЗ анализ для столбца "Сумма"
df_sorted_sum = df.sort_values(by='Сумма')
total_sum = df_sorted_sum['Сумма'].sum()
df_sorted_sum['Кумулятивная сумма'] = df_sorted_sum['Сумма'].cumsum()
df_sorted_sum['Доля кумулятивной суммы'] = df_sorted_sum['Кумулятивная сумма'] / total_sum

# Сохранение результатов в новый Excel-файл
output_file = 'C:/Users/Misha/Desktop/Pandas/кумулятивная_сумма.xlsx'
df_sorted_sum.to_excel(output_file, index=False)

print("Кумулятивная сумма сохранена в файл:", output_file)