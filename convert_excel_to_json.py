import pandas as pd
import json
import os
import sys

EXCEL_FILE = "sales.xlsx"
JSON_FILE = "sales-data.json"

gradients = [
    "linear-gradient(135deg, #FFD700 0%, #FFA500 100%)",  # МАГ
    "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
    "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
    "linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)",
    "linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)",
    "linear-gradient(135deg, #fa709a 0%, #fee140 100%)",
    "linear-gradient(135deg, #30cfd0 0%, #330867 100%)",
    "linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)",
    "linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%)"
]

PERCENT_COLS = ['% Доля ACC', 'Доля Послуг', 'Конверсія ПК', 'Конверсія ПК Offline', 'Доля УДС']
COUNT_COLS = ['Шт.', 'Чеки', 'ПЧ']
MONEY_COLS = ['ТО', 'ASP', 'Ср. Чек', 'ACC', 'Послуги грн', 'УДС']

def normalize_number(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, str):
        val = val.replace('%', '').replace(',', '.').strip()
    try:
        return float(val)
    except:
        return 0.0

def build_metrics(row, metric_columns):
    metrics = {}
    for col in metric_columns:
        num = normalize_number(row[col])

        if col in PERCENT_COLS:
            value = round(num, 2)
            unit = "%"
        elif col in COUNT_COLS:
            value = int(num)
            unit = "шт"
        elif col in MONEY_COLS:
            value = round(num, 2)
            unit = "грн"
        else:
            value = round(num, 2)
            unit = ""

        metrics[col] = {
            "value": value,
            "label": col,
            "unit": unit
        }
    return metrics

def main():
    if not os.path.exists(EXCEL_FILE):
        print("❌ Немає sales.xlsx")
        sys.exit(1)

    df = pd.read_excel(EXCEL_FILE, header=2, engine="openpyxl")

    if "ПК" not in df.columns or "Посада" not in df.columns:
        print("❌ У файлі немає колонок ПК або Посада")
        print("👉 Є:", list(df.columns))
        sys.exit(1)

    metric_columns = list(df.columns[2:])
    sales_data = []

    # 🔹 РЯДОК 0 = МАГ (загальні показники)
    total_row = df.iloc[0]
    total_metrics = build_metrics(total_row, metric_columns)

    sales_data.append({
        "id": 0,
        "name": "Загальні показники магазину",
        "position": "Всі продавці",
        "initials": "МАГ",
        "gradient": gradients[0],
        "metrics": total_metrics
    })

    # 🔹 ПРОДАВЦІ
    seller_id = 1
    for i in range(1, len(df)):
        row = df.iloc[i]
        name = str(row["ПК"]).strip()

        if not name or name == "nan":
            continue

        parts = name.split()
        initials = "".join(p[0] for p in parts[:2]).upper()

        metrics = build_metrics(row, metric_columns)

        person = {
            "id": seller_id,
            "name": name,
            "position": str(row["Посада"]) if pd.notna(row["Посада"]) else "продавец-консультант",
            "initials": initials,
            "gradient": gradients[seller_id % len(gradients)],
            "metrics": metrics
        }

        sales_data.append(person)
        seller_id += 1

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(sales_data, f, ensure_ascii=False, indent=2)

    print("✅ JSON оновлено")
    print("👥 Записів:", len(sales_data))

if __name__ == "__main__":
    main()
