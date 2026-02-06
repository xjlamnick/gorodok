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

def normalize_number(val):
    """Перетворює '6,933' або '16.56%' у число"""
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

    sales_data = []

    for idx, row in df.iterrows():
        name = str(row["ПК"]).strip()
        if not name or name == "nan":
            continue

        parts = name.split()
        initials = "".join(p[0] for p in parts[:2]).upper()

        metrics = {}

        for col in df.columns[2:]:
            raw_val = row[col]
            num = normalize_number(raw_val)

            if col in ['% Доля ACC', 'Доля Послуг', 'Конверсія ПК', 'Конверсія ПК Offline', 'Доля УДС']:
                value = round(num, 2)
                unit = "%"
            elif col in ['Шт.', 'Чеки', 'ПЧ']:
                value = int(num)
                unit = "шт"
            elif col in ['ТО', 'ASP', 'Ср. Чек', 'ACC', 'Послуги грн', 'УДС']:
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
            "id": len(sales_data) + 1,
            "name": name,
            "position": str(row["Посада"]) if pd.notna(row["Посада"]) else "продавец-консультант",
            "initials": initials,
            "gradient": gradients[len(sales_data) % len(gradients)],
            "metrics": metrics
        }

        sales_data.append(person)

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(sales_data, f, ensure_ascii=False, indent=2)

    print(f"✅ Успішно створено {JSON_FILE}")
    print(f"👥 Продавців: {len(sales_data)}")


if __name__ == "__main__":
    main()
