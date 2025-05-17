import cvxpy as cp
from pypfopt import expected_returns, risk_models, EfficientFrontier, DiscreteAllocation
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk
import os
import math

# Глобальные переменные для хранения данных
csv_data = {}
monthly_data_by_file = {}
results = {}

# Функции для загрузки и обработки данных
def load_csv():
    global csv_data, company_combobox

    file_paths = filedialog.askopenfilenames(
        title="Выберите CSV файлы",
        filetypes=(("CSV файлы", "*.csv"), ("Все файлы", "*.*"))
    )

    if not file_paths:
        return

    for file_path in file_paths:
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        data = pd.read_csv(file_path)
        csv_data[file_name] = data

    load_csv_label.config(text=f"Загружено файлов: {len(csv_data)}")
    company_combobox['values'] = list(csv_data.keys())
    company_combobox.set("")

    # Отладочный вывод загруженных данных
    print("Загруженные данные:")
    for file_name, data in csv_data.items():
        print(f"{file_name}:\n{data.head()}")

def split_data(data):
    data['Date'] = pd.to_datetime(data['Date'])
    data['Month'] = data['Date'].dt.to_period('M')
    return {month: group for month, group in data.groupby('Month')}

def split_all_data_by_month():
    global monthly_data_by_file
    if not csv_data:
        print("Сначала загрузите CSV файлы!")
        return

    monthly_data_by_file = {file_name: split_data(data) for file_name, data in csv_data.items()}

    # Отладочный вывод данных, разделенных по месяцам
    print("Данные, разделенные по месяцам:")
    for file_name, months in monthly_data_by_file.items():
        print(f"{file_name}:")
        for month, data in months.items():
            print(f"{month}:\n{data.head()}")

# Функции для расчета индикаторов
def calculate_RSI(data):
    buy_threshold = 30
    sell_threshold = 70
    actions = []
    money = 10000  # Капитал
    stocks_amount = 0  # Количество акций
    period = 5  # Период для расчета
    start_index = 1  # Индекс, с которого начинаем расчет

    close = data['Close'].round(2)  # Цена закрытия

    delta = close.diff()  # Разница между закрытиями

    # Рост и падение цены
    gain = delta.where(delta > 0, 0).round(2)
    loss = -delta.where(delta < 0, 0).round(2)

    # Обрезаем до нужного индекса
    gain = gain[start_index:]
    loss = loss[start_index:]

    # Рассчитываем средние только начиная с нужного индекса
    avg_gain = gain.rolling(window=period, min_periods=period).mean().round(2)
    avg_loss = loss.rolling(window=period, min_periods=period).mean().round(2)

    rs = np.where(avg_loss != 0, avg_gain / avg_loss, 0)

    rsi = 100 - (100 / (1 + rs)).round(2)  # Рассчитываем RSI

    for i in range(len(rsi)):
        # Логика покупки и продажи на основе RSI
        if rsi[i] < buy_threshold:
            # Покупаем, если хватает денег
            if i + 1 < len(close) and money >= close.iloc[i + 1]:
                stocks_to_buy = math.floor(money // close.iloc[i + 1])  # Сколько акций можем купить
                stocks_amount += stocks_to_buy  # Увеличиваем количество акций
                money -= stocks_to_buy * close.iloc[i + 1]  # Списываем деньги за акции
                actions.append(f"День {i + 1}: Покупаем {stocks_to_buy} штук по цене {close.iloc[i]}, осталось {money}, "
                               f"RSI: {rsi[i]}")
            else:
                actions.append(f"День {i + 1}: Денег недостаточно для покупки, RSI: {rsi[i]}, баланс: {money}")

        elif rsi[i] > sell_threshold and stocks_amount > 0:
            # Продаем все акции, если RSI выше порога продажи
            if i + 1 < len(close):
                money += stocks_amount * close.iloc[i + 1]  # Добавляем деньги от продажи всех акций
                actions.append(f"День {i + 1}: Продаем {stocks_amount} штук по цене {close.iloc[i + 1]}, RSI: {rsi[i]}, "
                               f"баланс: {money}")
                stocks_amount = 0  # Обнуляем количество акций

        else:
            # Ожидание, если RSI не в зоне покупки или продажи
            actions.append(f"День {i + 1}: Ждём, RSI: {rsi[i]}, баланс: {money}")

    # Итоговая сумма после завершения торговли
    final_cash = (money + stocks_amount * close.iloc[-1]).round(2)  # Общая сумма после продажи всех акций

    return final_cash

def calculate_MAC(data):
    return data['Close'].mean() * 1.5

def calculate_SO(data):
    return data['Close'].max() - data['Close'].min()

# Функция для анализа портфеля
def markovits_analyze(data):
    capital = 100000

    # Заменяем нулевые значения на NaN и удаляем строки с NaN
    data = data.replace(0, np.nan).dropna()

    # Рассчитываем годовую доходность для каждого актива
    mu = expected_returns.mean_historical_return(data)

    # Рассчитываем ковариационную матрицу (риск активов) с использованием метода Ledoit-Wolf
    Sigma = risk_models.CovarianceShrinkage(data).ledoit_wolf()

    # Оптимизация портфеля: Максимальный коэффициент Шарпа
    ef = EfficientFrontier(mu, Sigma, weight_bounds=(0, 1))
    sharpe_pwt = ef.max_sharpe()
    sharpe_pwt = ef.clean_weights()

    # Оптимизация портфеля: Минимальная волатильность
    ef1 = EfficientFrontier(mu, Sigma, weight_bounds=(0, 1))
    minvol_pwt = ef1.min_volatility()
    minvol_pwt = ef1.clean_weights()

    # Печатаем показатели эффективности портфеля
    ef.portfolio_performance(verbose=True)

    # Получаем текущие цены для дискретного распределения активов
    latest_prices = data.iloc[-1]

    # Дискретное распределение активов для минимальной волатильности
    allocation_minv, rem_minv = DiscreteAllocation(minvol_pwt, latest_prices, capital).lp_portfolio()

    # Дискретное распределение активов для максимального коэффициента Шарпа
    allocation_shp, rem_shp = DiscreteAllocation(sharpe_pwt, latest_prices, capital).lp_portfolio()

    data_results = pd.DataFrame({
        '1': capital,
        '2': sharpe_pwt,
        '3': minvol_pwt,
        '4': allocation_minv,
        '5': rem_minv,
        '6': allocation_shp,
        '7': rem_shp
    }).T.round(2)

    return data_results

# Функции для обработки данных с индикаторами
def process_monthly_data_with_indicators():
    global monthly_data_by_file

    if not monthly_data_by_file:
        print("Сначала разделите данные по месяцам!")
        return [], [], []

    results_rsi = []
    results_mac = []
    results_so = []

    for company, months in monthly_data_by_file.items():
        for month, data in months.items():
            rsi_result = calculate_RSI(data)
            mac_result = calculate_MAC(data)
            so_result = calculate_SO(data)

            results_rsi.append({"Month": month, "Company": company, "Value": rsi_result})
            results_mac.append({"Month": month, "Company": company, "Value": mac_result})
            results_so.append({"Month": month, "Company": company, "Value": so_result})

    df_rsi = pd.DataFrame(results_rsi)
    df_mac = pd.DataFrame(results_mac)
    df_so = pd.DataFrame(results_so)

    return df_rsi, df_mac, df_so

# Функции для выбора лучших акций
def select_best_stocks_by_sum_rsi(dataframe, top_count=2):
    """
    Выбирает лучшие акции на основе суммы значений за все месяцы.

    :param dataframe: DataFrame с акциями в столбцах и месяцами в индексе.
    :param top_count: Количество лучших акций для выбора.
    :return: Список названий лучших акций.
    """
    # Суммируем значения по всем строкам для каждой акции
    stock_sums = dataframe.sum(axis=0)

    # Сортируем акции по их общей сумме (по убыванию) и выбираем top_count
    top_stocks = stock_sums.sort_values(ascending=False).head(top_count).index.tolist()

    return top_stocks

def select_best_stocks_by_mac(results_mac, top_count=2):
    stock_sums = {}
    for result in results_mac:
        company = result["Company"]
        value = result["Value"]
        stock_sums[company] = stock_sums.get(company, 0) + value

    sorted_stocks = sorted(stock_sums.items(), key=lambda x: x[1], reverse=True)
    return [stock for stock, _ in sorted_stocks[:top_count]]

def select_best_stocks_by_so(results_so, top_count=2):
    stock_sums = {}
    for result in results_so:
        company = result["Company"]
        value = result["Value"]
        stock_sums[company] = stock_sums.get(company, 0) + value

    sorted_stocks = sorted(stock_sums.items(), key=lambda x: x[1], reverse=True)
    return [stock for stock, _ in sorted_stocks[:top_count]]

# Основная функция для запуска анализа
def run_analysis():
    global results_rsi, results_mac, results_so
    global df_rsi_pivot, df_mac_pivot, df_so_pivot

    if not csv_data:
        result_label.config(text="Сначала загрузите CSV файлы.")
        return

    split_all_data_by_month()
    results_rsi, results_mac, results_so = process_monthly_data_with_indicators()

    df_rsi_pivot = results_rsi.pivot(index="Month", columns="Company", values="Value")
    df_mac_pivot = results_mac.pivot(index="Month", columns="Company", values="Value")
    df_so_pivot = results_so.pivot(index="Month", columns="Company", values="Value")

    print("RSI DataFrame:\n", df_rsi_pivot)
    print("MAC DataFrame:\n", df_mac_pivot)
    print("SO DataFrame:\n", df_so_pivot)

    rsi_label.config(text="Выберите компанию")
    mac_label.config(text="Выберите компанию")
    so_label.config(text="Выберите компанию")

    # Проверяем, что данные имеют правильный формат
    if not df_rsi_pivot.empty and not df_mac_pivot.empty and not df_so_pivot.empty:
        try:
            # Отладочный вывод перед вызовом markovits_analyze
            print("Данные для анализа RSI:\n", df_rsi_pivot)
            print("Данные для анализа MAC:\n", df_mac_pivot)
            print("Данные для анализа SO:\n", df_so_pivot)

            # Собираем данные о ценах активов для анализа портфеля
            price_data = pd.concat([csv_data[company][['Date', 'Close']].set_index('Date').rename(columns={'Close': company}) for company in csv_data.keys()], axis=1)

            # Выполняем анализ портфеля
            portfolio_results = markovits_analyze(price_data)

            markoviz_and_rsi_label.config(text=f"Марковиц с индикатором RSI:\n{portfolio_results}")
            markoviz_and_mac_label.config(text=f"Марковиц с индикатором MAC:\n{portfolio_results}")
            markoviz_and_so_label.config(text=f"Марковиц с индикатором SO:\n{portfolio_results}")

            # Выбор лучших акций
            best_rsi_stocks = select_best_stocks_by_sum_rsi(df_rsi_pivot)
            best_mac_stocks = select_best_stocks_by_sum_rsi(df_mac_pivot)
            best_so_stocks = select_best_stocks_by_sum_rsi(df_so_pivot)

            # Обновляем списки лучших акций
            best_rsi_label.config(text="Лучшие акции по RSI:\n" + ", ".join(best_rsi_stocks))
            best_mac_label.config(text="Лучшие акции по MAC:\n" + ", ".join(best_mac_stocks))
            best_so_label.config(text="Лучшие акции по SO:\n" + ", ".join(best_so_stocks))

        except Exception as e:
            result_label.config(text=f"Ошибка при анализе данных: {e}")
            print(f"Ошибка при анализе данных: {e}")
    else:
        result_label.config(text="Данные индикаторов пусты или имеют неправильный формат.")
        print("Данные индикаторов пусты или имеют неправильный формат.")
def update_labels(selected_company):

    if not selected_company:
        rsi_label.config(text="Нет данных")
        mac_label.config(text="Нет данных")
        so_label.config(text="Нет данных")
        return

    if selected_company in df_rsi_pivot.columns:
        rsi_data = [f"{month}: {value:.2f}" for month, value in df_rsi_pivot[selected_company].items()]
    else:
        rsi_data = ["Нет данных"]

    if selected_company in df_mac_pivot.columns:
        mac_data = [f"{month}: {value:.2f}" for month, value in df_mac_pivot[selected_company].items()]
    else:
        mac_data = ["Нет данных"]

    if selected_company in df_so_pivot.columns:
        so_data = [f"{month}: {value:.2f}" for month, value in df_so_pivot[selected_company].items()]
    else:
        so_data = ["Нет данных"]

    rsi_label.config(text="\n".join(rsi_data))
    mac_label.config(text="\n".join(mac_data))
    so_label.config(text="\n".join(so_data))

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip_window is not None:
            return

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

        label = tk.Label(self.tooltip_window, text=self.text, background="black", fg="white", relief="solid",
                         borderwidth=1, font=("Courier", 11, "bold"))
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def create_interface():
    global load_csv_label, rsi_label, mac_label, so_label, result_label, company_combobox, results
    global best_rsi_label, best_mac_label, best_so_label
    global markoviz_and_rsi_label, markoviz_and_mac_label, markoviz_and_so_label

    root = tk.Tk()
    root.title("Анализ CSV файлов")
    root.geometry("1800x1500")
    root.configure(bg="black")

    terminal_font = ("Courier", 11, "bold")
    fg_color = "White"

    frame_top = tk.Frame(root, bd=2, relief="raised", bg="grey")
    frame_top.pack(fill="x", pady=10)

    load_csv_btn = tk.Button(frame_top, text="Загрузка CSV файлов", width=20, height="1", command=load_csv,
                             font=terminal_font, fg=fg_color, bg="black", relief="groove")
    load_csv_btn.pack(side="left", pady=10)

    load_csv_label = tk.Label(frame_top, text="Файлы не загружены", anchor="w", font=terminal_font,
                              fg=fg_color, bg="black", relief="groove")
    load_csv_label.pack(side="right", padx=50)

    select_label = tk.Label(frame_top, text="Выберите компанию: ", width=20, height="1", font=terminal_font,
                            fg=fg_color, bg="Black", relief="groove")
    select_label.pack(side="left", padx=10)

    company_combobox = ttk.Combobox(frame_top, state="readonly", font=terminal_font, background="black", width=20,
                                    height="1")
    company_combobox.pack(side="left", padx=10)

    company_combobox.bind(
        "<<ComboboxSelected>>",
        lambda event: update_labels(company_combobox.get())
    )

    frame_middle_part_1 = tk.Frame(root, bd=2, relief="solid", bg="black")
    frame_middle_part_1.pack(fill="x", pady=10)

    rsi_label = tk.Label(frame_middle_part_1, text="RSI", width=60, height=15, font=terminal_font, fg=fg_color,
                         bg="black", bd=2, relief="groove")
    rsi_label.pack(side="left")
    Tooltip(rsi_label, "Индикатор RSI: Показатель силы тренда")

    mac_label = tk.Label(frame_middle_part_1, text="MAC", width=60, height=15, font=terminal_font, fg=fg_color,
                         bg="black", bd=2, relief="groove")
    mac_label.pack(side="left", padx=65)
    Tooltip(mac_label, "Индикатор MAC: Показатель скользящего среднего")

    so_label = tk.Label(frame_middle_part_1, text="SO", width=60, height=15, font=terminal_font, fg=fg_color,
                        bg="black", bd=2, relief="groove")
    so_label.pack(side="right")
    Tooltip(so_label, "Индикатор SO: Стохастический осциллятор")

    frame_middle_part_2 = tk.Frame(root, bd=2, relief="solid", bg="black")
    frame_middle_part_2.pack(fill="x", pady=10)

    best_rsi_label = tk.Label(frame_middle_part_2, text="Лучшие акции", width=60, height=5, font=terminal_font,
                              fg=fg_color, bg="black", bd=2, relief="groove")
    best_rsi_label.pack(side="left")
    Tooltip(best_rsi_label, "Лучшие акции с использованием метода RSI")

    best_mac_label = tk.Label(frame_middle_part_2, text="Лучшие акции", width=60, height=5, font=terminal_font,
                              fg=fg_color, bg="black", bd=2, relief="groove")
    best_mac_label.pack(side="left", padx=65)
    Tooltip(best_mac_label, "Лучшие акции с использованием метода MAC")

    best_so_label = tk.Label(frame_middle_part_2, text="Лучшие акции", width=60, height=5, font=terminal_font,
                             fg=fg_color, bg="black", bd=2, relief="groove")
    best_so_label.pack(side="right")
    Tooltip(best_so_label, "Лучшие акции с использованием метода SO")

    frame_middle_part_3 = tk.Frame(root, bd=2, relief="solid", bg="black")
    frame_middle_part_3.pack(fill="x", pady=10)

    markoviz_and_rsi_label = tk.Label(frame_middle_part_3, text="Марковиц с индикатором RSI", width=60, height=15,
                                      font=terminal_font, fg=fg_color, bg="black", bd=2, relief="groove")
    markoviz_and_rsi_label.pack(side="left")
    Tooltip(markoviz_and_rsi_label, "1-Капитал"
                                    "\n2-Портфель с коэф. Шарпа"
                                    "\n3-Портфель с минимальной волательностью"
                                    "\n4-Распределение с минимальной волтельностью"
                                    "\n5-Остаток с минимальной волатильностью"
                                    "\n6-Распределение активов с коэф. Шарпа"
                                    "\n7-Остаток с коэф. Шарпа")


    markoviz_and_mac_label = tk.Label(frame_middle_part_3, text="Марковиц с индикатором MAC", width=60, height=15,
                                      font=terminal_font, fg=fg_color, bg="black", bd=2, relief="groove")
    markoviz_and_mac_label.pack(side="left", padx=65)
    Tooltip(markoviz_and_mac_label, "1-Капитал"
                                    "\n2-Портфель с коэф. Шарпа"
                                    "\n3-Портфель с минимальной волательностью"
                                    "\n4-Распределение с минимальной волтельностью"
                                    "\n5-Остаток с минимальной волатильностью"
                                    "\n6-Распределение активов с коэф. Шарпа"
                                    "\n7-Остаток с коэф. Шарпа")

    markoviz_and_so_label = tk.Label(frame_middle_part_3, text="Марковиц с индикатором SO", width=60, height=15,
                                     font=terminal_font, fg=fg_color, bg="black", bd=2, relief="groove")
    markoviz_and_so_label.pack(side="right")

    frame_bottom = tk.Frame(root, bd=2, relief="flat", bg="black")
    frame_bottom.pack(fill="x", pady=10)
    Tooltip(markoviz_and_so_label, "1-Капитал"
                                    "\n2-Портфель с коэф. Шарпа"
                                    "\n3-Портфель с минимальной волательностью"
                                    "\n4-Распределение с минимальной волтельностью"
                                    "\n5-Остаток с минимальной волатильностью"
                                    "\n6-Распределение активов с коэф. Шарпа"
                                    "\n7-Остаток с коэф. Шарпа")

    analyze_btn = tk.Button(frame_bottom, text="Запуск анализа", width=50, height=2, command=run_analysis,
                            font=terminal_font, fg=fg_color, bg="black", relief="groove")
    analyze_btn.pack(pady=25, padx=65)

    result_label = tk.Label(frame_bottom, text="", anchor="w", font=terminal_font, fg=fg_color, bg="black")
    result_label.pack(pady=10)

    root.resizable(False, False)
    root.mainloop()

if __name__ == "__main__":
    create_interface()