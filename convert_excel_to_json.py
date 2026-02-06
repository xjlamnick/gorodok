import pandas as pd
import json
import sys
import os

excel_file = 'sales.xlsx'
json_file = 'sales-data.json'
sheet_name = 'Sheet1'  # змінити, якщо у Excel інша назва листа

# Перевіряємо, чи існує Excel
if not os.path.exists(excel_file):
    print(f"Error: File '{excel_file}' not found!")
    sys.exit(1)

try:
    # Читаємо Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
    
    # Конвертуємо у список словників
    data = df.to_dict(orient='records')
    
    # Записуємо у JSON
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    
    print(f"Successfully converted '{excel_file}' to '{json_file}'")

except Exception as e:
    print(f"Error converting Excel to JSON: {e}")
    sys.exit(1)
