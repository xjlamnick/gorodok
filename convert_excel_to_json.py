import pandas as pd
import json
import os
import sys

EXCEL_FILE = "sales.xlsx"
JSON_FILE = "sales-data.json"

gradients = [
    'linear-gradient(135deg, #FFD700 0%, #FFA500 100%)',
    'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
    'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
    'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
    'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)',
    'linear-gradient(135deg, #fa709a 0%, #fee140 100%)',
    'linear-gradient(135deg, #30cfd0 0%, #330867 100%)',
    'linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)',
    'linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%)'
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


def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Файл {EXCEL_FILE} не знайдено")
        sys.exit(1)

    df = pd.read_excel(EXCEL_FILE, header=2, engine="openpyxl")

    required_columns = ["ПК", "Посада"]
    for col in required_columns:
        if col not in df.columns:
            print("❌ Немає колонки:", col)
            print("👉 Знайдені колонки:", list(df.columns))
            sys.exit(1)

    metric_columns = df.columns[2:]

    sales_data = []

    # ======================
    # ПРОДАВЦІ
    # ======================
    for idx, row in df.iterrows():
        name = str(row["ПК"]).strip()
        if not name or name == "nan":
            continue

        parts = name.split()
        initials = "".join(p[0] for p in parts[:2]).upper()

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

        sales_data.append({
            "id": len(sales_data) + 1,
            "name": name,
            "position": str(row["Посада"]) if pd.notna(row["Посада"]) else "продавец-консультант",
            "initials": initials,
            "gradient": gradients[(len(sales_data)) % len(gradients)],
            "metrics": metrics
        })

    # ======================
    # МАГАЗИН (ФОРМУЛИ)
    # ======================
    total_metrics = {}

    total_TO = df["ТО"].apply(normalize_number).sum()
    total_units = df["Шт."].apply(normalize_number).sum()
    total_checks = df["Чеки"].apply(normalize_number).sum()
    total_ACC = df["ACC"].apply(normalize_number).sum()
    total_services = df["Послуги грн"].apply(normalize_number).sum()
    total_UDS = df["УДС"].apply(normalize_number).sum()
    total_PCH = df["ПЧ"].apply(normalize_number).sum()

    avg_conv = df["Конверсія ПК"].apply(normalize_number).mean()
    avg_conv_off = df["Конверсія ПК Offline"].apply(normalize_number).mean()

    def safe_div(a, b):
        return round(a / b, 2) if b != 0 else 0

    computed = {
        "ТО": (round(total_TO, 2), "грн"),
        "Шт.": (int(total_units), "шт"),
        "Чеки": (int(total_checks), "шт"),
        "ASP": (safe_div(total_TO, total_units), "грн"),
        "Ср. Чек": (safe_div(total_TO, total_checks), "грн"),
        "КПЧ": (safe_div(total_units, total_checks), ""),
        "ACC": (round(total_ACC, 2), "грн"),
        "% Доля ACC": (safe_div(total_ACC * 100, total_TO), "%"),
        "Послуги грн": (round(total_services, 2), "грн"),
        "Доля Послуг": (safe_div(total_services * 100, total_TO), "%"),
        "ПЧ": (int(total_PCH), "шт"),
        "Конверсія ПК": (round(avg_conv, 2), "%"),
        "Конверсія ПК Offline": (round(avg_conv_off, 2), "%"),
        "УДС": (round(total_UDS, 2), "грн"),
        "Доля УДС": (safe_div(total_UDS * 100, total_TO), "%")
    }

    for key, (value, unit) in computed.items():
        total_metrics[key] = {
            "value": value,
            "label": key,
            "unit": unit
        }

    sales_data.insert(0, {
        "id": 0,
        "name": "Загальні показники магазину",
        "position": "Всі продавці",
        "initials": "МАГ",
        "gradient": gradients[0],
        "metrics": total_metrics
    })

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(sales_data, f, ensure_ascii=False, indent=2)

    print(f"✅ Успішно створено {JSON_FILE}")
    print(f"👥 Записів: {len(sales_data)}")


if __name__ == "__main__":
    main()
