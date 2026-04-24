#!/usr/bin/env python3
"""
Веб-приложение для парсинга чек-листов и генерации техкарт.
Поддерживает одиночные .docx, несколько .docx + .zip, выгрузку ZIP-архивом.
Добавлена история генераций с пагинацией, админ-панель с CRUD, автоочистка загрузок.
"""

import os
import uuid
import zipfile
import tempfile
import time
from pathlib import Path
from datetime import datetime, timedelta

from flask import Flask, request, render_template, jsonify, send_file

from parser.xml_parser import parse_checklist
from generator.docx_generator import TechCardGenerator
from db.database import SessionLocal
from db.models import Object, Instruction, Category as DBCategory

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['TECH_CARDS_FOLDER'] = 'tech_cards'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB
app.config['HISTORY_FILE'] = 'tech_cards/generation_history.json'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TECH_CARDS_FOLDER'], exist_ok=True)

# Совместимость Jinja2 с Vue.js
app.jinja_env.variable_start_string = '{?'
app.jinja_env.variable_end_string = '?}'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['docx', 'zip']


def cleanup_old_files(folder: str, hours: int = 24):
    """Удаляет файлы старше N часов"""
    now = time.time()
    folder_path = Path(folder)
    if not folder_path.exists():
        return
    count = 0
    for f in folder_path.glob("*"):
        if f.is_file() and (now - f.stat().st_mtime) > hours * 3600:
            f.unlink()
            count += 1
    if count > 0:
        print(f"🧹 Очищено {count} файлов из {folder} (старше {hours} ч)")


def load_history() -> list:
    """Загружает историю генераций"""
    import json
    history_path = Path(app.config['HISTORY_FILE'])
    if history_path.exists():
        with open(history_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_history(history: list):
    """Сохраняет историю генераций (не более 1000 записей)"""
    import json
    history_path = Path(app.config['HISTORY_FILE'])
    history = history[:1000]
    with open(history_path, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def add_to_history(filename: str, objects_count: int, room_name: str = ""):
    """Добавляет запись в историю"""
    history = load_history()
    file_path = Path(f"tech_cards/{filename}")
    size_kb = round(file_path.stat().st_size / 1024, 1) if file_path.exists() else 0

    history.insert(0, {
        'filename': filename,
        'display_name': room_name or Path(filename).stem,
        'objects': objects_count,
        'datetime': datetime.now().strftime('%d.%m.%Y %H:%M'),
        'timestamp': datetime.now().isoformat(),
        'download_url': f'/download/{filename}',
        'size_kb': size_kb
    })
    save_history(history)


def process_single_file(file, mode, generator):
    """Обрабатывает один .docx файл"""
    try:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = str(uuid.uuid4())[:8]
        filename = f"{timestamp}_{unique_id}_{file.filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        checklist_data = parse_checklist(filepath)
        room_name = checklist_data.room_name or Path(filepath).stem
        safe_name = "".join(c for c in room_name if c.isalnum() or c in (' ', '-', '_')).strip()
        output_filename = f"{safe_name}_tech_card.docx"
        output_path = os.path.join(app.config['TECH_CARDS_FOLDER'], output_filename)
        generator.generate(checklist_data, output_path, mode)

        add_to_history(output_filename, len(checklist_data.get_checked_items()), room_name)

        return {
            'filename': file.filename,
            'objects': len(checklist_data.get_checked_items()),
            'download_url': f'/download/{output_filename}'
        }
    except Exception as e:
        return {'filename': file.filename, 'error': str(e)}


def process_zip_file(file, mode, generator):
    """Обрабатывает ZIP архив"""
    results = []
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    unique_id = str(uuid.uuid4())[:8]
    filename = f"{timestamp}_{unique_id}_{file.filename}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(filepath, 'r') as zf:
            zf.extractall(tmpdir)

        docx_files = list(Path(tmpdir).glob("**/*.docx"))
        for docx_path in docx_files:
            if docx_path.name.startswith("~"):
                continue
            try:
                checklist_data = parse_checklist(str(docx_path))
                room_name = checklist_data.room_name or docx_path.stem
                safe_name = "".join(c for c in room_name if c.isalnum() or c in (' ', '-', '_')).strip()
                output_filename = f"{safe_name}_tech_card.docx"
                output_path = os.path.join(app.config['TECH_CARDS_FOLDER'], output_filename)
                generator.generate(checklist_data, output_path, mode)

                add_to_history(output_filename, len(checklist_data.get_checked_items()), room_name)

                results.append({
                    'filename': docx_path.name,
                    'objects': len(checklist_data.get_checked_items()),
                    'download_url': f'/download/{output_filename}'
                })
            except Exception as e:
                results.append({'filename': docx_path.name, 'error': str(e)})

    return results


def create_zip_archive(files_info):
    """Создаёт ZIP-архив из сгенерированных техкарт"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_filename = f"tech_cards_{timestamp}.zip"
    zip_path = os.path.join(app.config['TECH_CARDS_FOLDER'], zip_filename)

    with zipfile.ZipFile(zip_path, 'w') as zf:
        for f in files_info:
            docx_path = f['download_url'].replace('/download/', 'tech_cards/')
            if os.path.exists(docx_path):
                zf.write(docx_path, Path(docx_path).name)

    return f'/download/{zip_filename}'


# === АВТООЧИСТКА ПРИ ЗАПУСКЕ ===
cleanup_old_files(app.config['UPLOAD_FOLDER'], 24)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/admin')
def admin():
    return render_template('admin.html')


@app.route('/history')
def history():
    return render_template('history.html')


# ============================================================
# API: ГЕНЕРАЦИЯ
# ============================================================

@app.route('/api/generate', methods=['POST'])
def api_generate():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Нет файла'}), 400

    is_multi = request.form.get('multi') == 'true'
    mode = int(request.form.get('mode', 1))
    generator = TechCardGenerator()
    results = []

    try:
        if is_multi:
            files = [request.files['file']]
            i = 1
            while f'file_{i}' in request.files:
                files.append(request.files[f'file_{i}'])
                i += 1

            for file in files:
                ext = file.filename.rsplit('.', 1)[1].lower()
                if ext == 'zip':
                    zip_results = process_zip_file(file, mode, generator)
                    results.extend(zip_results)
                else:
                    result = process_single_file(file, mode, generator)
                    results.append(result)
        else:
            file = request.files['file']
            ext = file.filename.rsplit('.', 1)[1].lower()

            if ext == 'zip':
                results = process_zip_file(file, mode, generator)
            else:
                result = process_single_file(file, mode, generator)
                results.append(result)

        total_objects = sum(r.get('objects', 0) for r in results)
        errors = [r for r in results if 'error' in r]
        success_files = [r for r in results if 'error' not in r]

        zip_url = None
        if len(success_files) > 1:
            zip_url = create_zip_archive(success_files)

        return jsonify({
            'success': True,
            'total_files': len(results),
            'total_objects': total_objects,
            'files': results,
            'errors': errors,
            'zip_url': zip_url
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/download/<filename>')
def download(filename):
    path = f"tech_cards/{filename}"
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return jsonify({'error': 'Файл не найден'}), 404


# ============================================================
# API: ИСТОРИЯ
# ============================================================

@app.route('/api/history')
def api_history():
    page = int(request.args.get('page', 1))
    per_page = 100

    history = load_history()
    total = len(history)
    total_pages = max(1, (total + per_page - 1) // per_page)

    start = (page - 1) * per_page
    end = start + per_page

    return jsonify({
        'history': history[start:end],
        'page': page,
        'total_pages': total_pages,
        'total': total
    })


@app.route('/api/history/delete', methods=['POST'])
def api_delete_history():
    data = request.get_json()
    filenames = data.get('filenames', [])
    before_date = data.get('before_date', None)

    deleted = 0
    history = load_history()

    if before_date:
        cutoff = datetime.strptime(before_date, '%Y-%m-%d')
        to_delete = []
        for item in history:
            item_date = datetime.fromisoformat(item['timestamp'])
            if item_date < cutoff:
                to_delete.append(item['filename'])
        filenames = to_delete

    for filename in filenames:
        path = os.path.join(app.config['TECH_CARDS_FOLDER'], filename)
        if os.path.exists(path):
            os.remove(path)
            deleted += 1

    history = [h for h in history if h['filename'] not in filenames]
    save_history(history)

    return jsonify({'success': True, 'deleted': deleted})


# ============================================================
# API: КАТЕГОРИИ
# ============================================================

@app.route('/api/categories')
def api_categories():
    session = SessionLocal()
    cats = session.query(DBCategory).order_by(DBCategory.sort_order).all()
    result = [{'id': c.id, 'name': c.name, 'sort_order': c.sort_order} for c in cats]
    session.close()
    return jsonify({'categories': result})


# ============================================================
# API: ОБЪЕКТЫ
# ============================================================

@app.route('/api/objects')
def api_objects():
    session = SessionLocal()
    objects = session.query(Object).order_by(Object.sort_priority, Object.display_name).all()
    cats = {c.id: c.name for c in session.query(DBCategory).all()}
    result = [{
        'id': o.id,
        'category_id': o.category_id,
        'category_name': cats.get(o.category_id, ''),
        'display_name': o.display_name,
        'normalized_name': o.normalized_name,
        'base_name': o.base_name,
        'modifier': o.modifier or '',
        'sort_priority': o.sort_priority,
        'instruction_count': len(o.instructions)
    } for o in objects]
    session.close()
    return jsonify({'objects': result})


@app.route('/api/objects', methods=['POST'])
def api_create_object():
    data = request.get_json()
    session = SessionLocal()
    obj = Object(
        display_name=data.get('display_name', ''),
        normalized_name=data.get('normalized_name', data.get('display_name', '')),
        base_name=data.get('base_name', data.get('display_name', '')),
        modifier=data.get('modifier'),
        sort_priority=data.get('sort_priority', 0),
        category_id=data.get('category_id', 1)
    )
    session.add(obj)
    session.commit()
    session.refresh(obj)
    obj_id = obj.id
    session.close()
    return jsonify({'success': True, 'id': obj_id})


@app.route('/api/objects/<int:obj_id>', methods=['PUT'])
def api_update_object(obj_id):
    data = request.get_json()
    session = SessionLocal()
    obj = session.query(Object).get(obj_id)
    if obj:
        for key in ['display_name', 'normalized_name', 'base_name', 'modifier', 'sort_priority', 'category_id']:
            if key in data:
                setattr(obj, key, data[key])
        session.commit()
        session.close()
        return jsonify({'success': True})
    session.close()
    return jsonify({'success': False, 'error': 'Не найден'}), 404


@app.route('/api/objects/<int:obj_id>', methods=['DELETE'])
def api_delete_object(obj_id):
    session = SessionLocal()
    obj = session.query(Object).get(obj_id)
    if obj:
        session.delete(obj)
        session.commit()
        session.close()
        return jsonify({'success': True})
    session.close()
    return jsonify({'success': False, 'error': 'Не найден'}), 404


# ============================================================
# API: ИНСТРУКЦИИ
# ============================================================

@app.route('/api/objects/<int:obj_id>/instructions')
def api_object_instructions(obj_id):
    session = SessionLocal()
    obj = session.query(Object).get(obj_id)
    if obj:
        result = [{
            'id': i.id,
            'cleaning_method': i.cleaning_method or '',
            'product_name': i.product_name or '',
            'cleaning_technique': i.cleaning_technique or '',
            'concentration': i.concentration or '',
            'temperature': i.temperature or '',
            'exposure_time': i.exposure_time or '',
            'inventory': i.inventory or '',
            'frequency': i.frequency or '',
            'executor': i.executor or '',
            'control_method': i.control_method or '',
            'instruction_number': i.instruction_number or ''
        } for i in obj.instructions]
        session.close()
        return jsonify({'instructions': result})
    session.close()
    return jsonify({'instructions': []})


@app.route('/api/objects/<int:obj_id>/instructions', methods=['POST'])
def api_create_instruction(obj_id):
    data = request.get_json()
    session = SessionLocal()
    instr = Instruction(
        object_id=obj_id,
        cleaning_method=data.get('cleaning_method', ''),
        product_name=data.get('product_name', ''),
        cleaning_technique=data.get('cleaning_technique', ''),
        concentration=data.get('concentration', ''),
        temperature=data.get('temperature', ''),
        exposure_time=data.get('exposure_time', ''),
        inventory=data.get('inventory', ''),
        frequency=data.get('frequency', ''),
        executor=data.get('executor', ''),
        control_method=data.get('control_method', ''),
        instruction_number=data.get('instruction_number', '')
    )
    session.add(instr)
    session.commit()
    session.refresh(instr)
    instr_id = instr.id
    session.close()
    return jsonify({'success': True, 'id': instr_id})


@app.route('/api/instructions/<int:instr_id>', methods=['PUT'])
def api_update_instruction(instr_id):
    data = request.get_json()
    session = SessionLocal()
    instr = session.query(Instruction).get(instr_id)
    if instr:
        fields = ['cleaning_method', 'product_name', 'cleaning_technique', 'concentration',
                   'temperature', 'exposure_time', 'inventory', 'frequency', 'executor',
                   'control_method', 'instruction_number']
        for key in fields:
            if key in data:
                setattr(instr, key, data[key])
        session.commit()
        session.close()
        return jsonify({'success': True})
    session.close()
    return jsonify({'success': False, 'error': 'Не найдена'}), 404


@app.route('/api/instructions/<int:instr_id>', methods=['DELETE'])
def api_delete_instruction(instr_id):
    session = SessionLocal()
    instr = session.query(Instruction).get(instr_id)
    if instr:
        session.delete(instr)
        session.commit()
        session.close()
        return jsonify({'success': True})
    session.close()
    return jsonify({'success': False, 'error': 'Не найдена'}), 404


# ============================================================
# API: ЭКСПОРТ CSV
# ============================================================

@app.route('/api/export_csv')
def api_export_csv():
    import csv
    from io import StringIO

    session = SessionLocal()
    objects = session.query(Object).order_by(Object.sort_priority, Object.display_name).all()
    cats = {c.id: c.name for c in session.query(DBCategory).all()}

    si = StringIO()
    writer = csv.writer(si, delimiter=';')
    writer.writerow(['object_id', 'category_name', 'normalized_name', 'display_name', 'base_name',
                     'modifier', 'sort_priority', 'cleaning_method', 'product_name', 'concentration',
                     'temperature', 'exposure_time', 'inventory', 'frequency', 'executor', 'control_method'])

    for obj in objects:
        if obj.instructions:
            for instr in obj.instructions:
                writer.writerow([obj.id, cats.get(obj.category_id, ''), obj.normalized_name,
                               obj.display_name, obj.base_name, obj.modifier or '', obj.sort_priority,
                               instr.cleaning_method or '', instr.product_name or '',
                               instr.concentration or '', instr.temperature or '',
                               instr.exposure_time or '', instr.inventory or '',
                               instr.frequency or '', instr.executor or '', instr.control_method or ''])
        else:
            writer.writerow([obj.id, cats.get(obj.category_id, ''), obj.normalized_name,
                           obj.display_name, obj.base_name, obj.modifier or '', obj.sort_priority,
                           '', '', '', '', '', '', '', '', ''])

    session.close()

    output = si.getvalue().encode('utf-8-sig')
    si.close()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
        tmp.write(output)
        return send_file(tmp.name, as_attachment=True, download_name='db_export.csv')


if __name__ == '__main__':
    print("🚀 Запуск сервера на http://localhost:5000")
    print(f"📂 Папка техкарт: {app.config['TECH_CARDS_FOLDER']}")
    print(f"🧹 Автоочистка uploads: 24 часа")
    app.run(debug=True, host='0.0.0.0', port=5000)