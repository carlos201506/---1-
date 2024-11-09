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
aggregation_dict = {'Unnamed: 2': 'first', 'Unnamed: 3': 'sum', 'Unnamed: 4': 'sum', 'Unnamed: 5': 'sum'}
df_aggregated = df_sorted.groupby(level=0).agg(aggregation_dict)

# Вычисление результата деления значений столбца "Unnamed: 5" на значения столбца "Unnamed: 3"
df_aggregated["Unnamed: 4"] = df_aggregated["Unnamed: 5"] / df_aggregated["Unnamed: 3"]

# Расчет доли каждого товара
df_aggregated['Доля'] = df_aggregated['Unnamed: 5'] / df_aggregated['Unnamed: 5'].sum()

# Расчет кумулятивной доли
df_aggregated['Аккум.доля'] = df_aggregated['Доля'].cumsum()

# Классификация товаров по ABC-анализу
df_aggregated["Категория (ABC)"] = ' '
df_aggregated.loc[df_aggregated['Аккум.доля'] < 0.6, 'Категория (ABC)'] = 'A'
df_aggregated.loc[(df_aggregated['Аккум.доля'] > 0.6) & (df_aggregated['Аккум.доля'] < 0.75), 'Категория (ABC)'] = 'B'
df_aggregated.loc[df_aggregated['Аккум.доля'] > 0.75, 'Категория (ABC)'] = 'C'

# Выбор столбца с "Расходом годовым"
df_aggregated['Расход годовой'] = df_aggregated['Unnamed: 3']

# Проверка на нечисловые данные
if not pd.to_numeric(df_aggregated['Расход годовой'], errors='coerce').notnull().all():
    print("Столбец 'Расход годовой' содержит нечисловые данные. Обработайте их перед выполнением анализа.")
    exit()

# Проверка на нулевые значения
if df_aggregated['Расход годовой'].eq(0).sum() > 0:
    print("Столбец 'Расход годовой' содержит нулевые значения. Обработайте их перед выполнением анализа.")
    exit()

# Расчет стандартного отклонения
std = df_aggregated['Расход годовой'].std(ddof=0)

# Расчет среднего значения
mean = df_aggregated['Расход годовой'].mean()

# Проверка на деление на ноль
if mean == 0:
    print("Среднее значение равно нулю. Коэффициент вариации не может быть рассчитан.")
    exit()

# Расчет коэффициента вариации
cv = (std / mean) * 100

# Классификация товаров по XYZ-анализу
df_aggregated['Категория (XYZ)'] = ' '
df_aggregated.loc[df_aggregated['Расход годовой'] <= 0.25 * mean, 'Категория (XYZ)'] = 'Z'
df_aggregated.loc[(df_aggregated['Расход годовой'] > 0.25 * mean) & (df_aggregated['Расход годовой'] <= 0.5 * mean), 'Категория (XYZ)'] = 'Y'
df_aggregated.loc[df_aggregated['Расход годовой'] > 0.5 * mean, 'Категория (XYZ)'] = 'X'

# Изменение названий столбцов
df_aggregated.rename(columns={'Unnamed: 2': 'Единицы измерения', 'Unnamed: 3': 'Количество', 'Unnamed: 4': 'Цена за шт без НДС в среднем', 'Unnamed: 5': 'Сумма без НДС'}, inplace=True)

# Создание нового DataFrame с наименованиями
df_with_names = df_aggregated.copy()
df_with_names.insert(0, 'Наименование', df_aggregated.index)

# Определение порядка сортировки
order = ['AX', 'AY', 'AZ', 'BX', 'BY', 'BZ', 'CX', 'CY', 'CZ']

# Создание нового столбца для хранения порядка
df_with_names['Сортировка'] = df_with_names['Категория (ABC)'] + df_with_names['Категория (XYZ)']

# Применение пользовательского порядка сортировки
df_with_names['Сортировка'] = pd.Categorical(df_with_names['Сортировка'], categories=order, ordered=True)

# Сортировка DataFrame
df_sorted_final = df_with_names.sort_values(by='Сортировка')

# Удаление временного столбца 'Сортировка'
df_sorted_final.drop(columns=['Сортировка'], inplace=True)

# Сохранение нового DataFrame в Excel-файл
df_sorted_final.to_excel('C:/Users/Misha/Desktop/Pandas/main_AX_AY_AZ_BX_BY_BZ_CX_CY_CZ_analysis_results3.xlsx', index=False)

# Вывод отсортированного DataFrame
print(df_sorted_final)