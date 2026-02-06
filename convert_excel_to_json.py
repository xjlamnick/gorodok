import pandas as pd
import json
import sys
import os

excel_file = 'sales.xlsx'
json_file = 'sales-data.json'

# 1️⃣ Перевірка наявності файлу
if not os.path.exists(excel_file):
    print(f"Error: File '{excel_file}' not found!")
    sys.exit(1)

try:
    # 2️⃣ Зчитуємо всі листи Excel
    all_sheets = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
    print(f"Found sheets: {list(all_sheets.keys())}")

    # 3️⃣ Виводимо перші 5 рядків кожного листа для перевірки
    for sheet_name, df in all_sheets.items():
        print(f"\nSheet '{sheet_name}' preview:")
        print(df.head())

    # 4️⃣ Конвертуємо лише перший лист у JSON (можна змінити)
    first_sheet_name = list(all_sheets.keys())[0]
    data = all_sheets[first_sheet_name].to_dict(orient='records')

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"\nSuccessfully converted '{first_sheet_name}' to '{json_file}'")

except Exception as e:
    print(f"Error converting Excel to JSON: {e}")
    sys.exit(1)
