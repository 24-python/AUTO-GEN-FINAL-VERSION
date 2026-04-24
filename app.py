#!/usr/bin/env python3
"""
Веб-приложение для парсинга чек-листов и генерации техкарт.
Поддерживает одиночные .docx, несколько .docx + .zip, выгрузку ZIP-архивом.
"""

import os
import uuid
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime
from io import BytesIO

from flask import Flask, request, render_template, jsonify, send_file

from parser.xml_parser import parse_checklist
from generator.docx_generator import TechCardGenerator
from db.database import SessionLocal
from db.models import Object, Category as DBCategory

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('tech_cards', exist_ok=True)

# Совместимость Jinja2 с Vue.js
app.jinja_env.variable_start_string = '{?'
app.jinja_env.variable_end_string = '?}'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['docx', 'zip']


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
        output_path = f"tech_cards/{safe_name}_tech_card.docx"
        generator.generate(checklist_data, output_path, mode)

        return {
            'filename': file.filename,
            'objects': len(checklist_data.get_checked_items()),
            'download_url': f'/download/{safe_name}_tech_card.docx'
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
                output_path = f"tech_cards/{safe_name}_tech_card.docx"
                generator.generate(checklist_data, output_path, mode)
                results.append({
                    'filename': docx_path.name,
                    'objects': len(checklist_data.get_checked_items()),
                    'download_url': f'/download/{safe_name}_tech_card.docx'
                })
            except Exception as e:
                results.append({'filename': docx_path.name, 'error': str(e)})

    return results


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/admin')
def admin():
    return render_template('admin.html')


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

        # Если несколько успешных файлов — создаём ZIP-архив
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


def create_zip_archive(files_info):
    """Создаёт ZIP-архив из сгенерированных техкарт"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    zip_filename = f"tech_cards_{timestamp}.zip"
    zip_path = f"tech_cards/{zip_filename}"

    with zipfile.ZipFile(zip_path, 'w') as zf:
        for f in files_info:
            docx_path = f['download_url'].replace('/download/', 'tech_cards/')
            if os.path.exists(docx_path):
                zf.write(docx_path, Path(docx_path).name)

    return f'/download/{zip_filename}'


@app.route('/download/<filename>')
def download(filename):
    path = f"tech_cards/{filename}"
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return jsonify({'error': 'Файл не найден'}), 404


@app.route('/api/objects')
def api_objects():
    session = SessionLocal()
    objects = session.query(Object).order_by(Object.sort_priority, Object.display_name).all()
    cats = {c.id: c.name for c in session.query(DBCategory).all()}
    result = [{
        'id': o.id,
        'category_name': cats.get(o.category_id, ''),
        'display_name': o.display_name,
        'normalized_name': o.normalized_name,
        'sort_priority': o.sort_priority,
        'instruction_count': len(o.instructions)
    } for o in objects]
    session.close()
    return jsonify({'objects': result})


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
    return jsonify({'success': False, 'error': 'Объект не найден'}), 404


@app.route('/api/export_csv')
def api_export_csv():
    import csv
    from io import StringIO

    session = SessionLocal()
    objects = session.query(Object).order_by(Object.sort_priority, Object.display_name).all()
    cats = {c.id: c.name for c in session.query(DBCategory).all()}

    si = StringIO()
    writer = csv.writer(si, delimiter=';')
    writer.writerow(['object_id', 'category_name', 'normalized_name', 'display_name', 'base_name', 'modifier', 'sort_priority', 'cleaning_method', 'product_name', 'concentration', 'temperature', 'exposure_time', 'inventory', 'frequency', 'executor', 'control_method'])

    for obj in objects:
        if obj.instructions:
            for instr in obj.instructions:
                writer.writerow([obj.id, cats.get(obj.category_id, ''), obj.normalized_name, obj.display_name, obj.base_name, obj.modifier or '', obj.sort_priority, instr.cleaning_method or '', instr.product_name or '', instr.concentration or '', instr.temperature or '', instr.exposure_time or '', instr.inventory or '', instr.frequency or '', instr.executor or '', instr.control_method or ''])
        else:
            writer.writerow([obj.id, cats.get(obj.category_id, ''), obj.normalized_name, obj.display_name, obj.base_name, obj.modifier or '', obj.sort_priority, '', '', '', '', '', '', '', '', ''])

    session.close()

    output = si.getvalue().encode('utf-8-sig')
    si.close()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
        tmp.write(output)
        return send_file(tmp.name, as_attachment=True, download_name='db_export.csv')


if __name__ == '__main__':
    print("🚀 Запуск сервера на http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)