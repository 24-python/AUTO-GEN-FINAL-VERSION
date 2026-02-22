#!/usr/bin/env python3
"""
Веб-приложение для парсинга чек-листов с элементами управления
Использует прямой парсинг XML структуры .docx
"""

import os
import uuid
from pathlib import Path
from datetime import datetime

from flask import Flask, request, render_template, jsonify

# Импортируем XML парсер
from parser.xml_parser import parse_checklist

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
    """Форматирует результат для веб-интерфейса"""
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

    for i, item in enumerate(data.get_checked_items()[:20]):
        print(f"  {i + 1}. {item.name}")

    grouped = data.group_checked_by_category()

    for category, items in grouped.items():
        category_data = {
            'name': category.value,
            'items': []
        }

        for item in items:
            category_data['items'].append({
                'name': item.name,
                'subcategory': item.subcategory,
                'markers': item.markers
            })

        result['categories'].append(category_data)

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
        # Сохраняем файл
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = str(uuid.uuid4())[:8]
        filename = f"{timestamp}_{unique_id}_{file.filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        file.save(filepath)

        # Парсим через XML парсер
        print(f"\n🔍 Парсинг файла: {filepath}")
        data = parse_checklist(filepath)

        result = format_result(data)

        return jsonify({
            'success': True,
            'data': result
        })

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)