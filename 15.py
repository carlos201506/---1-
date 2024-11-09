


# Размер заказа (РЗ 1)
df_with_names.loc[condition_fixed_interval, 'Размер заказа (РЗ 1)'] = df_with_names.loc[condition_fixed_interval, 'Максимальный желательный уровень запасов (МЖЗ)']-df_with_names['Остаток']+((df_with_names['Количество'] / working_days_in_year) * order_lead_time).astype('float64')


# Размер заказа (РЗ 1) 6
df_with_names.loc[condition_fixed_replenishment_periodicity, 'Размер заказа (РЗ 1)'] = df_with_names.loc[condition_fixed_interval, 'Максимальный желательный уровень запасов (МЖЗ)']-df_with_names['Остаток']+((df_with_names['Количество'] / working_days_in_year) * order_lead_time).astype('float64')