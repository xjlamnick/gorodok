#!/usr/bin/env python3
"""
Автоматичний скрипт для оновлення даних та завантаження на GitHub
"""

import subprocess
import sys
import os
from datetime import datetime

def run_command(cmd, description):
    """Виконує команду та виводить результат"""
    print(f"  {description}...")
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ❌ Помилка: {result.stderr}")
        return False
    return True

def main():
    print("\n" + "="*60)
    print("  ОНОВЛЕННЯ ТА ЗАВАНТАЖЕННЯ НА GITHUB")
    print("="*60 + "\n")
    
    # Перевірка наявності файлу
    if not os.path.exists('sales.xlsx'):
        print("❌ Файл sales.xlsx не знайдено!")
        print("   Переконайтесь, що файл знаходиться в цій папці\n")
        sys.exit(1)
    
    # Крок 1: Оновлення даних
    print("📊 Крок 1/4: Оновлення даних з Excel")
    if not run_command('python3 update_data.py sales.xlsx', 'Конвертація даних'):
        sys.exit(1)
    
    # Крок 2: Git add
    print("\n📦 Крок 2/4: Підготовка файлів")
    if not run_command('git add sales-data.json index.html', 'Додавання файлів до git'):
        sys.exit(1)
    
    # Крок 3: Git commit
    print("\n💾 Крок 3/4: Збереження змін")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_msg = f"Оновлення даних: {timestamp}"
    if not run_command(f'git commit -m "{commit_msg}"', 'Створення commit'):
        print("  ⚠️  Немає змін для збереження")
    
    # Крок 4: Git push
    print("\n🚀 Крок 4/4: Завантаження на GitHub")
    if not run_command('git push', 'Відправка на GitHub'):
        print("\n❌ Помилка при завантаженні на GitHub!")
        print("\n💡 Можливі причини:")
        print("   • Не налаштовано git remote")
        print("   • Потрібна авторизація")
        print("   • Немає інтернет з'єднання\n")
        sys.exit(1)
    
    print("\n" + "="*60)
    print("  ✅ УСПІХ!")
    print("="*60)
    print("\n📱 Ваш сайт оновлюється на GitHub Pages")
    print("⏱️  Зачекайте 1-2 хвилини, потім оновіть сторінку\n")
    print("🔗 Посилання:")
    print("   https://ваш-username.github.io/sales-team/\n")

if __name__ == "__main__":
    main()
