#!/usr/bin/env python3
"""
Скрипт для оновлення даних з Excel файлу
"""

import pandas as pd
import json
import sys
import os

def update_data_from_excel(excel_file='sales.xlsx'):
    """Оновлює sales-data.json з Excel файлу"""
    
    if not os.path.exists(excel_file):
        print(f"❌ Файл '{excel_file}' не знайдено!")
        return False
    
    print(f"📂 Читаю файл: {excel_file}")
    
    try:
        # Читаємо файл (рядок 3 - заголовки, рядок 4+ - дані)
        df = pd.read_excel(excel_file, header=2)
        
        # Градієнти
        gradients = [
            'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
            'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
            'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
            'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)',
            'linear-gradient(135deg, #fa709a 0%, #fee140 100%)',
            'linear-gradient(135deg, #30cfd0 0%, #330867 100%)',
            'linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)',
            'linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%)',
            'linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%)'
        ]
        
        sales_data = []
        
        for idx, row in df.iterrows():
            if pd.notna(row['ПК']):
                name = str(row['ПК'])
                
                # Генеруємо ініціали
                name_parts = name.split()
                if len(name_parts) >= 2:
                    initials = ''.join([p[0] for p in name_parts[:2]]).upper()
                else:
                    initials = name[0].upper()
                
                # Створюємо метрики (стовпці з 3-го)
                metrics = {}
                for col in df.columns[2:]:  # Починаємо з 3-го стовпця
                    val = row[col]
                    
                    # Визначаємо тип даних та одиниці
                    if pd.isna(val):
                        val = 0
                    
                    # Перевіряємо чи це відсоток (значення між 0 і 1)
                    if col in ['% Доля ACC', 'Доля Послуг', 'Конверсія ПК', 'Конверсія ПК Offline', 'Доля УДС']:
                        value = round(float(val) * 100, 2) if pd.notna(val) else 0
                        unit = '%'
                    elif col in ['Шт.', 'Чеки', 'ПЧ']:
                        value = int(val) if pd.notna(val) else 0
                        unit = 'шт'
                    elif col in ['ТО', 'ASP', 'Ср. Чек', 'ACC', 'Послуги грн', 'УДС']:
                        value = round(float(val), 2) if pd.notna(val) else 0
                        unit = 'грн'
                    else:
                        value = round(float(val), 2) if pd.notna(val) else 0
                        unit = ''
                    
                    metrics[col] = {
                        'value': value,
                        'label': col,
                        'unit': unit
                    }
                
                person = {
                    'id': len(sales_data) + 1,
                    'name': name,
                    'position': str(row['Посада']) if pd.notna(row['Посада']) else 'Менеджер з продажу',
                    'initials': initials,
                    'gradient': gradients[len(sales_data) % len(gradients)],
                    'metrics': metrics
                }
                sales_data.append(person)
        
        # Зберігаємо
        with open('sales-data.json', 'w', encoding='utf-8') as f:
            json.dump(sales_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ Оновлено дані для {len(sales_data)} продавців:")
        for p in sales_data:
            print(f"   • {p['name']}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Помилка: {e}")
        return False


if __name__ == "__main__":
    excel_file = sys.argv[1] if len(sys.argv) > 1 else 'sales.xlsx'
    
    print("\n" + "="*50)
    print("  ОНОВЛЕННЯ ДАНИХ")
    print("="*50 + "\n")
    
    if update_data_from_excel(excel_file):
        print("\n" + "="*50)
        print("  ✅ ГОТОВО!")
        print("="*50 + "\n")
    else:
        sys.exit(1)
