import pandas as pd
import numpy as np
import math
from openpyxl import Workbook


# Метод расчета RSI
def calculate_rsi(data):
    buy_threshold = 30
    sell_threshold = 70
    actions = []
    money = 10000  # Капитал
    stocks_amount = 0  # Количество акций
    period = 10  # Период для расчета
    start_index = 1  # Индекс, с которого начинаем расчет

    close = data['Close'].round(2)  # Цена закрытия
    delta = close.diff()  # Разница между закрытиями

    # Рост и падение цены
    gain = delta.where(delta > 0, 0).round(2)
    loss = -delta.where(delta < 0, 0).round(2)

    # Средние значения
    avg_gain = gain.rolling(window=period, min_periods=period).mean().round(2)
    avg_loss = loss.rolling(window=period, min_periods=period).mean().round(2)

    rs = np.where(avg_loss != 0, avg_gain / avg_loss, 0)
    rsi = 100 - (100 / (1 + rs)).round(2)

    # Логика торговли
    for i in range(len(rsi)):
        if rsi[i] < buy_threshold and money >= close[i]:
            stocks_to_buy = math.floor(money // close[i])  # Сколько акций можем купить
            stocks_amount += stocks_to_buy
            money -= stocks_to_buy * close[i]

        elif rsi[i] > sell_threshold and stocks_amount > 0:
            money += stocks_amount * close[i]
            stocks_amount = 0

    final_cash = round(money + stocks_amount * close.iloc[-1], 2)
    return final_cash, rsi


# Разделение данных по месяцам
def split_data_by_month(file_path):
    df = pd.read_csv(file_path)
    df['Date'] = pd.to_datetime(df['Date'])
    df['Month'] = df['Date'].dt.to_period('M')
    return df


# Расчет RSI для каждого месяца
def process_data(df):
    final_results = []
    monthly_data = {}

    for month, group in df.groupby('Month'):
        group = group.sort_values(by='Date').reset_index(drop=True)
        final_cash, rsi = calculate_rsi(group)
        group['RSI'] = rsi  # Добавляем RSI в данные
        monthly_data[str(month)] = group  # Сохраняем данные
        final_results.append((str(month), final_cash))  # Добавляем месяц и финальный кэш

    return monthly_data, final_results


# Сохранение данных в Excel
def write_to_excel(monthly_data, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for month, data in monthly_data.items():
            data.to_excel(writer, sheet_name=str(month), index=False)
    print(f"Данные сохранены в {output_file}")


# Основной код
file_path = "TSLA.csv"
df = split_data_by_month(file_path)  # Разделяем данные
monthly_data, final_results = process_data(df)  # Рассчитываем RSI

# Сохраняем в Excel
write_to_excel(monthly_data, "output_data.xlsx")

# Выводим результаты для приложения
print("Результаты торговли по месяцам:")
for month, final_cash in final_results:
    print(f"Месяц: {month}, Финальный капитал: {final_cash}")
