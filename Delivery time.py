import pandas as pd
import numpy as np

FORMAT = '%d.%m.%Y %H:%M:%S'


provider_orders_data = pd.read_excel(
    'data/Заказы поставщикам.xlsx',
    sheet_name='Поступление товаров',
    skiprows=0,
    usecols=['Период', 'Код "Инфор"', 'Заказ поставщику', 'Дата'],
    converters={'Код "Инфор"': str,
                'Период': lambda x: pd.to_datetime(x, format=FORMAT),
                'Дата': lambda x: pd.to_datetime(x, format=FORMAT)})
provider_orders_data['Код инфор+Заказ поставщику'] = (
    provider_orders_data['Код "Инфор"'] +
    provider_orders_data['Заказ поставщику']
)

place_orders_data = pd.read_excel(
    'data/Размещение заказов.xlsx',
    sheet_name='Заказ поставщику',
    skiprows=0,
    usecols=['Регистратор', 'Дата', 'Код "Инфор"'],
    converters={'Код "Инфор"':str,
                'Дата': lambda x: pd.to_datetime(x, format=FORMAT)})
place_orders_data['сцепить'] = (
    place_orders_data['Код "Инфор"'] +
    place_orders_data['Регистратор']
)

# -----------------ЭТАП 1-----------------
df = pd.merge(
    provider_orders_data,
    place_orders_data.drop_duplicates('сцепить'),
    how='left',
    left_on='Код инфор+Заказ поставщику',
    right_on='сцепить')

df['Срок доставки'] = np.where(
    df['Дата_y'].isnull(),
    df['Период'] - df['Дата_x'],
    df['Период'] - df['Дата_y'])

# -----------------ЭТАП 3 и 4-----------------
def median_by_id(row):
    id = row['Код "Инфор"_x']
    temp_df = df.loc[df['Код "Инфор"_x'] == id]
    return temp_df['Срок доставки'].median()

df['Медиана'] = df.apply(lambda row: median_by_id(row), axis=1)

# -----------------ЭТАП 5-----------------
df['Срок доставки очищенный'] = np.where(
    np.logical_or(
        df['Срок доставки'] < (df['Медиана'] / 2),
        df['Срок доставки'] > (df['Медиана'] * 2)
    ),
        False,
        df['Срок доставки']
)

# -----------------ЭТАП 6-----------------
def high_quantile_by_id(row):
    id = row['Код "Инфор"_x']
    temp_df = df.loc[np.logical_and(
        df['Код "Инфор"_x'] == id, df['Срок доставки очищенный'] != pd.Timedelta(0)
    )]
    return temp_df['Срок доставки очищенный'].quantile([0.75])

df['Верхний квартиль'] = df.apply(lambda row: high_quantile_by_id(row), axis=1)

# -----------------ЭТАП 7-----------------
df['Срок доставки из верхнего квартиля'] = np.where(
    df['Срок доставки очищенный'] >= df['Верхний квартиль'],
    df['Срок доставки очищенный'],
    False)

# -----------------ЭТАП 8-----------------
def median_of_high_quantile_by_id(row):
    id = row['Код "Инфор"_x']
    temp_df = df.loc[np.logical_and(
        df['Код "Инфор"_x'] == id,
        df['Срок доставки из верхнего квартиля'] != pd.Timedelta(0)
    )]
    return temp_df['Срок доставки из верхнего квартиля'].median()

df['Медиана верхнего квартиля'] = df.apply(
    lambda row: median_of_high_quantile_by_id(row), axis=1)
df = df.drop_duplicates('Код "Инфор"_x')

# -----------------ЭТАП 9-----------------
# df['Итоговое значение'] = df['Медиана верхнего квартиля'].dt.round("D") # Округление математически
df['Срок доставки'] = df['Медиана верхнего квартиля'].dt.ceil("D") # Округление вверх

# -----------------УДАЛЕНИЕ ЛИШНЕГО-----------------
df = df[['Код "Инфор"_x', 'Срок доставки']]
df['Срок производства'] = 0

# -----------------ВЫВОД-----------------
df.to_excel("Данные для загрузки.xlsx")