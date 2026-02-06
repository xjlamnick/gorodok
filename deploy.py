import pandas as pd
import json

# Читаємо Excel
df = pd.read_excel('sales.xlsx', sheet_name='Sheet1')

# Конвертуємо у список словників
data = df.to_dict(orient='records')

# Записуємо у JSON
with open('sales-data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
