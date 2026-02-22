#!/usr/bin/env python3
"""
Веб-приложение для парсинга чек-листов с категориями
"""

import os
import uuid
from pathlib import Path
from datetime import datetime

from flask import Flask, request, render_template, jsonify

# Импортируем парсер с категориями
from parser.xml_parser import parse_checklist
from parser.models import Category

app = Flask(__name__)

# Конфигурация
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

# Создаём папки
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def allowed_file(filename):
    """Проверяет разрешённый тип файла"""
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


def format_result(data):
    """Форматирует результат для веб-интерфейса с категориями"""
    result = {
        'file_name': Path(data.file_path).name,
        'room_name': data.room_name or 'Не указан',
        'enterprise': data.enterprise or 'Не указано',
        'total_items': len(data.items),
        'checked_items': len(data.get_checked_items()),
        'categories': []
    }

    # Отладка в консоль
    print(f"\n📊 Всего элементов: {len(data.items)}")
    print(f"✅ Отмечено: {len(data.get_checked_items())}")

    # Покажем первые несколько элементов для проверки
    for i, item in enumerate(data.get_checked_items()[:10]):
        print(f"  {i + 1}. [{item.category.value}] {item.name}")

    grouped = data.group_checked_by_category()

    # Сортируем категории в логическом порядке (сверху вниз)
    category_order = [
        Category.SURFACE,  # 1. Поверхности (потолок, стены, пол)
        Category.PLUMBING,  # 2. Сантехника
        Category.SANITARY_POST,  # 3. Санпост
        Category.HOUSEHOLD_APPLIANCES,  # 4. Бытовая техника
        Category.THERMAL_EQUIPMENT,  # 5. Тепловое оборудование
        Category.REFRIGERATION_EQUIPMENT,  # 6. Холодильное оборудование
        Category.TECH_EQUIPMENT,  # 7. Технологическое оборудование
        Category.PACKAGING_EQUIPMENT,  # 8. Упаковочное оборудование
        Category.DISHWASHING_EQUIPMENT,  # 9. Посудомоечное оборудование
        Category.CLEANING_EQUIPMENT,  # 10. Моечный инвентарь
        Category.INVENTORY,  # 11. Инвентарь
        Category.FURNITURE,  # 12. Мебель
        Category.OFFICE_EQUIPMENT,  # 13. Офисная техника
        Category.DOSING_EQUIPMENT,  # 14. Дозирующее оборудование
        Category.PPE,  # 15. СИЗ
        Category.OTHER  # 16. Прочее
    ]

    for category in category_order:
        if category in grouped:
            items = grouped[category]
            category_data = {
                'name': category.value,
                'items': [{'name': item.name} for item in items]
            }
            result['categories'].append(category_data)

            # Вывод в консоль
            print(f"\n{category.value}: {len(items)} элементов")
            for item in items[:5]:  # Покажем первые 5 из каждой категории
                print(f"  • {item.name}")
            if len(items) > 5:
                print(f"  ... и ещё {len(items) - 5}")

    return result


@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Загрузка и парсинг файла"""

    if 'file' not in request.files:
        return jsonify({'error': 'Нет файла'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Разрешены только .docx файлы'}), 400

    try:
        # Сохраняем файл с уникальным именем
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = str(uuid.uuid4())[:8]
        filename = f"{timestamp}_{unique_id}_{file.filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        file.save(filepath)

        print(f"\n🔍 Начинаем обработку файла: {filepath}")

        # Парсим через парсер с категориями
        data = parse_checklist(filepath)

        # Форматируем результат
        result = format_result(data)

        print(f"\n✅ Обработка завершена. Найдено позиций: {len(data.items)}")

        return jsonify({
            'success': True,
            'data': result
        })

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/health', methods=['GET'])
def health():
    """Проверка работоспособности"""
    return jsonify({'status': 'ok', 'message': 'Сервер работает'})


if __name__ == '__main__':
    print("🚀 Запуск сервера на http://localhost:5000")
    print("📁 Файлы будут сохраняться в папку:", app.config['UPLOAD_FOLDER'])
    app.run(debug=True, host='0.0.0.0', port=5000)