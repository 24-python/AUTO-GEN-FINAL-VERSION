#!/usr/bin/env python3
"""
Универсальный скрипт для обновления периодичности и исполнителя в CSV-файле
на основе графика санитарной обработки (.docx).

Не изменяет базу данных — только обновляет CSV-файл.

Использование:
    python update_frequency_from_schedule.py "производственных помещений.docx"
    python update_frequency_from_schedule.py "производственных помещений.docx" --csv=db_export_all1111.csv
    python update_frequency_from_schedule.py file1.docx file2.docx --csv=db_export_all1111.csv
"""

import sys
import zipfile
import csv
import re
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def extract_text(cell) -> str:
    """Извлекает текст из ячейки"""
    texts = cell.findall('.//w:t', NAMESPACES)
    return ' '.join(t.text or '' for t in texts).strip()


def extract_category_name(root) -> str:
    """Извлекает категорию помещения из заголовка"""
    rows = root.findall('.//w:tr', NAMESPACES)
    if rows:
        first_cell_text = extract_text(rows[0])

        room_categories = [
            'Общего назначения', 'Производственное', 'Бытовое', 'Складское',
            'Санитарное', 'Вспомогательное', 'Моечное', 'Техническое', 'Офисное'
        ]
        for rc in room_categories:
            if rc.lower() in first_cell_text.lower():
                return rc

        match = re.match(r'(\w+)\w*\s+помещения', first_cell_text, re.IGNORECASE)
        if match:
            base = match.group(1).lower()
            mapping = {
                'производствен': 'Производственное',
                'бытов': 'Бытовое',
                'склад': 'Складское',
                'санитар': 'Санитарное',
                'вспомогатель': 'Вспомогательное',
                'моеч': 'Моечное',
                'техническ': 'Техническое',
                'офис': 'Офисное',
                'общего': 'Общего назначения',
            }
            for key, value in mapping.items():
                if key in base:
                    return value

    return None


def normalize_object_name(name: str) -> str:
    """Упрощённая нормализация объекта для сопоставления"""
    name = name.lower().strip()
    name = re.sub(r'[^\w\s]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name.strip()


def parse_schedule(file_path: str) -> dict:
    """
    Парсит график санитарной обработки.
    Возвращает {object_name: {data: {level: {cleaning_method: frequency}}, executor: ...}}
    """
    result = {}

    with zipfile.ZipFile(file_path, 'r') as z:
        with z.open('word/document.xml') as f:
            root = etree.fromstring(f.read())

    tables = root.findall('.//w:tbl', NAMESPACES)
    if not tables:
        print("❌ Таблица не найдена")
        return result

    table = tables[0]
    rows = table.findall('.//w:tr', NAMESPACES)

    current_object = None
    current_method = None
    current_executor = None
    collected_data = {}

    for r_idx, row in enumerate(rows[4:], start=4):
        cells = row.findall('.//w:tc', NAMESPACES)
        if len(cells) < 3:
            continue

        cell_texts = [extract_text(c) for c in cells]

        col0 = cell_texts[0] if len(cell_texts) > 0 else ''
        col1 = cell_texts[1] if len(cell_texts) > 1 else ''
        col2 = cell_texts[2] if len(cell_texts) > 2 else ''
        col3 = cell_texts[3] if len(cell_texts) > 3 else ''
        col4 = cell_texts[4] if len(cell_texts) > 4 else ''

        # Пропускаем строки-разделители секций
        if col0 and not col1:
            if current_object and collected_data:
                result[current_object] = {
                    'data': collected_data.copy(),
                    'executor': current_executor
                }
                current_object = None
                collected_data = {}
            continue

        # Новый объект: есть номер пункта
        if col0 and re.match(r'\d+\.\d+\.?', col0):
            if current_object and collected_data:
                result[current_object] = {
                    'data': collected_data.copy(),
                    'executor': current_executor
                }

            current_object = col1.strip()
            collected_data = {}
            current_executor = None

            if col2 in ['основная', 'поддерживающая', 'генеральная']:
                current_method = col2
            else:
                current_method = 'основная'

            collected_data[current_method] = {
                'мойка': col3 if col3 else None,
                'дезинфекция': col4 if col4 else None,
            }

            if len(cell_texts) > 4 and cell_texts[4]:
                current_executor = cell_texts[4]

        # Продолжение объекта: новый уровень обработки
        elif col2 in ['основная', 'поддерживающая', 'генеральная']:
            current_method = col2
            if current_object is not None:
                collected_data[current_method] = {
                    'мойка': col3 if col3 else None,
                    'дезинфекция': col4 if col4 else None,
                }
            if len(cell_texts) > 4 and cell_texts[4]:
                current_executor = cell_texts[4]

        # Продолжение объекта: добавляем текст к имени
        elif not col0 and col1 and current_object:
            current_object += ' ' + col1

        # Обновляем исполнителя
        if len(cell_texts) > 4 and cell_texts[4]:
            current_executor = cell_texts[4]

    # Сохраняем последний объект
    if current_object and collected_data:
        result[current_object] = {
            'data': collected_data.copy(),
            'executor': current_executor
        }

    return result


def get_best_frequency(data: dict, cleaning_method: str) -> str:
    """
    Выбирает периодичность по правилу:
    основная → поддерживающая → генеральная
    """
    for level in ['основная', 'поддерживающая', 'генеральная']:
        if level in data:
            freq = data[level].get(cleaning_method)
            if freq:
                return freq
    return None


def find_rows_in_csv(rows: list, obj_name: str, room_category_name: str) -> list:
    """
    Ищет строки в CSV по имени объекта и категории помещения.
    Приоритет: точное совпадение > частичное совпадение.
    """
    norm = normalize_object_name(obj_name)

    exact_matches = []
    partial_matches = []

    for row in rows:
        csv_norm = normalize_object_name(row.get('normalized_name', ''))
        csv_disp = normalize_object_name(row.get('display_name', ''))
        csv_base = normalize_object_name(row.get('base_name', ''))
        csv_room = row.get('room_category_name', '').strip()

        # Категория должна совпадать
        if csv_room != room_category_name:
            continue

        # Точное совпадение
        if norm == csv_norm or norm == csv_disp or norm == csv_base:
            exact_matches.append(row)
        # Частичное совпадение (имя объекта содержится в CSV-имени)
        elif (norm in csv_norm or norm in csv_disp or norm in csv_base or
              csv_norm in norm or csv_disp in norm or csv_base in norm):
            partial_matches.append(row)

    return exact_matches if exact_matches else partial_matches


def update_csv(csv_path: str, schedule_data: dict, room_category_name: str, output_path: str = None):
    """
    Обновляет frequency и executor в CSV-файле на основе графика.
    """
    if output_path is None:
        p = Path(csv_path)
        output_path = str(p.parent / f"{p.stem}_updated{p.suffix}")

    rows = []
    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')
        fieldnames = reader.fieldnames
        for row in reader:
            rows.append(row)

    updated = 0
    not_found = 0
    found_count = 0

    for obj_name, obj_data in schedule_data.items():
        executor = obj_data['executor']
        data = obj_data['data']

        found_rows = find_rows_in_csv(rows, obj_name, room_category_name)

        if not found_rows:
            not_found += 1
            print(f"  ⚠️ Не найден: {obj_name}")
            continue

        found_count += 1

        for row in found_rows:
            cleaning_method = row.get('cleaning_method', '').strip().lower()
            if cleaning_method in ['мойка', 'дезинфекция']:
                frequency = get_best_frequency(data, cleaning_method)
                if frequency:
                    old_freq = row.get('frequency', '')
                    if old_freq != frequency:
                        row['frequency'] = frequency
                        updated += 1

                if executor:
                    row['executor'] = executor

    with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        writer.writerows(rows)

    print(f"  ✅ Найдено объектов: {found_count}")
    print(f"  ✅ Обновлено строк: {updated}")
    print(f"  ⚠️ Не найдено объектов: {not_found}")
    print(f"  📁 Сохранено: {output_path}")

    return output_path


def main():
    if len(sys.argv) < 2:
        print("Использование:")
        print("  python update_frequency_from_schedule.py <файл1.docx> [файл2.docx ...] [--csv=db_export_all1111.csv]")
        print()
        print("Пример:")
        print("  python update_frequency_from_schedule.py 'производственных помещений.docx' --csv=db_export_all1111.csv")
        print("  python update_frequency_from_schedule.py file1.docx file2.docx --csv=db_export_all1111.csv")
        return

    csv_path = "db_export_all1111.csv"
    docx_files = []

    for a in sys.argv[1:]:
        if a.startswith('--csv='):
            csv_path = a.split('=', 1)[1]
        elif a.endswith('.docx'):
            docx_files.append(a)

    if not docx_files:
        print("❌ Укажите хотя бы один .docx файл")
        return

    all_schedules = {}

    for file_path in docx_files:
        if not Path(file_path).exists():
            print(f"❌ Файл не найден: {file_path}")
            continue

        print(f"\n{'='*60}")
        print(f"📄 Обработка: {file_path}")

        schedule_data = parse_schedule(file_path)
        print(f"  📋 Объектов в графике: {len(schedule_data)}")

        with zipfile.ZipFile(file_path, 'r') as z:
            with z.open('word/document.xml') as f:
                root = etree.fromstring(f.read())
        room_category_name = extract_category_name(root)
        print(f"  🏢 Категория помещения: {room_category_name}")

        if not room_category_name:
            print("  ❌ Не удалось определить категорию помещения")
            continue

        all_schedules[room_category_name] = schedule_data

    print(f"\n{'='*60}")
    print(f"📝 Обновление CSV: {csv_path}")

    current_csv = csv_path
    for room_cat_name, sched_data in all_schedules.items():
        print(f"\n🔧 Применяем график для: {room_cat_name}")
        current_csv = update_csv(current_csv, sched_data, room_cat_name)

    print(f"\n{'='*60}")
    print(f"🎉 Итоговый файл: {current_csv}")


if __name__ == "__main__":
    main()