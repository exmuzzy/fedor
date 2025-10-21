#!/bin/bash
# Скрипт для быстрого запуска парсинга PDF

echo "🚀 Запуск парсинга PDF спецификаций..."
echo ""

# Активация виртуального окружения
source venv/bin/activate

# Запуск скрипта
python parse_pdfs.py

echo ""
echo "✅ Готово! Открываю файл specifications_full.xlsx"
open specifications_full.xlsx

