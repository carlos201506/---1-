import pandas as pd

# Чтение данных из Excel-файла, пропуская строки перед нужной
df = pd.read_excel('C:/Users/Misha/Desktop/Альтена/17052024 Альтена.xls', header=None)

# Извлечение названий столбцов из четвёртой строки (индекс 3)
df.columns = df.iloc[3]

# Удаление строк с заголовками
df = df[4:]

# Удаление лишних пробелов в названиях столбцов
df.columns = df.columns.str.strip()

# Выбор необходимых столбцов
columns_to_keep = ['ТМЦ', 'ЕдиницаИзмерения', 'Количество', 'ПерваяЦена', 'СуммаСНДС']
df_filtered = df[columns_to_keep]

# Сохранение отфильтрованного DataFrame в новый Excel-файл,
# начиная с 2-й строки (индекс 1) и 2-го столбца (индекс 1)
df_filtered.to_excel('C:/Users/Misha/Desktop/Альтена/2.xlsx', index=False, header=False, startrow=1, startcol=1)

print("Данные успешно сохранены, начиная с 2-й строки и 2-го столбца.")