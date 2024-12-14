from hdbcli import dbapi
import pandas as pd

# Подключение к базе данных
conn = dbapi.connect(
#
)

# Запрос
query = """

"""

try:
    # Выполнение запроса
    cursor = conn.cursor()
    cursor.execute(query)
    
    # Получение данных
    data = cursor.fetchall()
    
    # Создание DataFrame
    df = pd.DataFrame(data, columns=["CALWEEK", "ZDEPART", "ZPRDCLASS", "NetSales"])
    df["NetSales"] = df["NetSales"].astype(float)
    
    # Замена значений департаментов
    department_mapping = {
        "W": "Women",
        "U": "U&L",
        "M": "Men",
        "K": "Kids",
        "I": "IBM",
        "D": "Denim",
        "A": "A&P",
        "C": "Common"
    }
    df["ZDEPART"] = df["ZDEPART"].map(department_mapping).fillna(df["ZDEPART"])
    
    # Добавление сегмента
    segment_mapping = {
        "Kids": "Одежда",
        "Women": "Одежда",
        "Men": "Одежда",
        "Denim": "Одежда",
        "A&P": "Сопутка",
        "U&L": "Сопутка",
        "IBM": "Сопутка",
        "Common": "Маркетинг"
    }
    df["Сегмент"] = df["ZDEPART"].map(segment_mapping).fillna("Маркетинг")
    
    # Функция для вычисления метрик
    def calculate_department_metrics(group):
        total_sales = group["NetSales"].sum()
        
        # Вычисление долей продаж
        group["Доля продаж (%)"] = (group["NetSales"] / total_sales * 100).round(2)
        
        # Коррекция до 100%
        diff = 100 - group["Доля продаж (%)"].sum()
        if diff != 0:
            max_idx = group["NetSales"].idxmax()
            group.loc[max_idx, "Доля продаж (%)"] += diff
        
        # Сортировка и накопительный итог
        group = group.sort_values("NetSales", ascending=False)
        group["Накопительный итог (%)"] = group["Доля продаж (%)"].cumsum().round(2)
        
        # Преобразование в проценты для отображения
        group["Доля продаж (%)"] = group["Доля продаж (%)"].apply(lambda x: f"{x:.2f}%")
        group["Накопительный итог (%)"] = group["Накопительный итог (%)"].apply(lambda x: f"{x:.2f}%")
        
        group["Категория"] = group["Накопительный итог (%)"].apply(assign_abc_category)
        group["Ранг"] = range(1, len(group) + 1)
        
        return group

    # Функция для присваивания категории ABC
    def assign_abc_category(value):
        value = float(value.replace('%', ''))
        if value <= 80:
            return "А"
        elif 80 < value <= 95:
            return "В"
        else:
            return "С"

    # Применение группировки
    df = df.groupby(["CALWEEK", "ZDEPART"], group_keys=False).apply(calculate_department_metrics)
    
    # Замена точек на запятые
    df["NetSales"] = df["NetSales"].astype(str).str.replace('.', ',')
    df["Доля продаж (%)"] = df["Доля продаж (%)"].astype(str).str.replace('.', ',')
    df["Накопительный итог (%)"] = df["Накопительный итог (%)"].astype(str).str.replace('.', ',')
    
    # Сохранение результатов в Excel
    output_file = "query_results_with_segments.xlsx"
    df = df[["CALWEEK", "ZDEPART", "Сегмент", "ZPRDCLASS", "NetSales", "Ранг", "Доля продаж (%)", "Накопительный итог (%)", "Категория"]]
    df.to_excel(
        output_file, 
        index=False, 
        header=["Неделя", "Департамент", "Сегмент", "Категория", "Чистые продажи", "Ранг", "Доля продаж (%)", "Накопительный итог (%)", "Категория"]
    )
    
    print(f"Данные успешно сохранены в файл {output_file}.")
finally:
    cursor.close()
    conn.close()
