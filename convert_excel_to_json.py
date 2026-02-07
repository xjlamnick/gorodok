import pandas as pd
import json
import os
import sys

EXCEL_FILE = "sales.xlsx"
JSON_FILE = "sales-data.json"

gradients = [
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

    # =========================
    # 🔹 ЗАГАЛЬНІ ПОКАЗНИКИ МАГАЗИНУ (ФОРМУЛИ)
    # =========================
    total_metrics = {}

    for col in metric_columns:
        values = df[col].apply(normalize_number)

        if col in PERCENT_COLS:
            # формули
            if col == '% Доля ACC':
                acc_sum = df['ACC'].apply(normalize_number).sum()
                to_sum = df['ТО'].apply(normalize_number).sum()
                value = (acc_sum / to_sum * 100) if to_sum else 0
            elif col == 'Доля Послуг':
                services_sum = df['Послуги грн'].apply(normalize_number).sum()
                to_sum = df['ТО'].apply(normalize_number).sum()
                value = (services_sum / to_sum * 100) if to_sum else 0
            elif col == 'Доля УДС':
                uds_sum = df['УДС'].apply(normalize_number).sum()
                to_sum = df['ТО'].apply(normalize_number).sum()
                value = (uds_sum / to_sum * 100) if to_sum else 0
            else:
                value = values.mean()

            unit = "%"
            value = round(value, 2)

        elif col in COUNT_COLS:
            value = int(values.sum())
            unit = "шт"

        elif col in MONEY_COLS:
            value = round(values.sum(), 2)
            unit = "грн"

        else:
            value = round(values.sum(), 2)
            unit = ""

        total_metrics[col] = {
            "value": value,
            "label": col,
            "unit": unit
        }

    sales_data.append({
        "id": 0,
        "name": "Загальні показники магазину",
        "position": "Всі продавці",
        "initials": "МАГ",
        "gradient": "linear-gradient(135deg, #FFD700 0%, #FFA500 100%)",
        "metrics": total_metrics
    })

    # =========================
    # 🔹 ПРОДАВЦІ
    # =========================
    for idx, row in df.iterrows():
        name = str(row["ПК"]).strip()
        if not name or name == "nan":
            continue

        parts = name.split()
        initials = "".join(p[0] for p in parts[:2]).upper()

        metrics = {}

        for col in metric_columns:
            raw_val = row[col]
            num = normalize_number(raw_val)

            if col in PERCENT_COLS:
                # фікс відсотків
                if num <= 1:
                    num = num * 100
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

        person = {
            "id": len(sales_data),
            "name": name,
            "position": str(row["Посада"]) if pd.notna(row["Посада"]) else "продавец-консультант",
            "initials": initials,
            "gradient": gradients[(len(sales_data) - 1) % len(gradients)],
            "metrics": metrics
        }

        sales_data.append(person)

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(sales_data, f, ensure_ascii=False, indent=2)

    print(f"✅ Успішно створено {JSON_FILE}")
    print(f"👥 Продавців: {len(sales_data) - 1}")


if __name__ == "__main__":
    main()
