# diagnose_schedule.py
"""
Диагностика сопоставления графика и CSV.
Показывает, какие объекты из графика найдены/не найдены в CSV.
"""

import sys
import csv
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from update_frequency_from_schedule import (
    parse_schedule,
    extract_category_name,
    normalize_object_name,
    get_best_frequency,
)
import zipfile
from lxml import etree

NAMESPACES = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Конфигурация
DOCX_FILE = "../производственных помещений.docx"
CSV_FILE = "../db_export_all.csv"


def main():
    if len(sys.argv) >= 2:
        docx_file = sys.argv[1]
    else:
        docx_file = DOCX_FILE

    if len(sys.argv) >= 3:
        csv_file = sys.argv[2]
    else:
        csv_file = CSV_FILE

    print(f"📄 График: {docx_file}")
    print(f"📋 CSV: {csv_file}")
    print()

    # Парсим график
    schedule = parse_schedule(docx_file)
    print(f"Объектов в графике: {len(schedule)}")

    # Категория помещения
    with zipfile.ZipFile(docx_file, 'r') as z:
        with z.open('word/document.xml') as f:
            root = etree.fromstring(f.read())
    room_cat = extract_category_name(root)
    print(f"Категория помещения: {room_cat}")
    print()

    # Читаем CSV
    csv_rows = []
    with open(csv_file, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter=';')
        for row in reader:
            csv_rows.append(row)
    print(f"Строк в CSV: {len(csv_rows)}")
    print()

    # Сопоставление
    found_list = []
    not_found_list = []
    partial_list = []  # Найдены, но другая категория

    for obj_name, obj_data in schedule.items():
        norm = normalize_object_name(obj_name)
        executor = obj_data['executor']
        data = obj_data['data']

        matched = False
        wrong_category = False

        for row in csv_rows:
            csv_norm = normalize_object_name(row.get('normalized_name', ''))
            csv_disp = normalize_object_name(row.get('display_name', ''))
            csv_base = normalize_object_name(row.get('base_name', ''))

            if (norm in csv_norm or csv_norm in norm or
                    norm in csv_disp or csv_disp in norm or
                    norm in csv_base or csv_base in norm):

                csv_room = row.get('room_category_name', '').strip()
                if csv_room == room_cat:
                    # Полное совпадение — выводим детали
                    cleaning_method = row.get('cleaning_method', '')
                    csv_freq = row.get('frequency', '')
                    new_freq = get_best_frequency(data, cleaning_method.lower())

                    if not matched:
                        found_list.append({
                            'name': obj_name,
                            'norm': norm,
                            'csv_name': row.get('display_name', ''),
                            'executor': executor,
                            'instructions': []
                        })
                        matched = True

                    found_list[-1]['instructions'].append({
                        'method': cleaning_method,
                        'csv_freq': csv_freq,
                        'new_freq': new_freq,
                        'levels': {k: v for k, v in data.items()}
                    })
                else:
                    wrong_category = True

        if not matched:
            if wrong_category:
                partial_list.append(obj_name)
            else:
                not_found_list.append(obj_name)

    # Вывод найденных
    print("=" * 70)
    print(f"✅ НАЙДЕНО В CSV: {len(found_list)}")
    print("=" * 70)
    for item in found_list[:10]:
        print(f"\n📦 {item['name']}")
        print(f"   CSV-объект: {item['csv_name']}")
        print(f"   Исполнитель из графика: {item['executor']}")
        print(f"   Инструкции:")
        for instr in item['instructions']:
            print(f"      {instr['method']}: было='{instr['csv_freq']}' → станет='{instr['new_freq']}'")
            for level, methods in instr['levels'].items():
                m = methods.get(instr['method'].lower())
                print(f"         {level}: {m}")

    # Вывод не найденных
    print()
    print("=" * 70)
    print(f"❌ НЕ НАЙДЕНО В CSV: {len(not_found_list)}")
    print("=" * 70)
    for name in not_found_list[:20]:
        print(f"   - {name}")

    # Найдены в другой категории
    if partial_list:
        print()
        print("=" * 70)
        print(f"⚠️ НАЙДЕНЫ В ДРУГОЙ КАТЕГОРИИ: {len(partial_list)}")
        print("=" * 70)
        for name in partial_list[:10]:
            print(f"   - {name}")

    print()
    print(f"Итого: найдено {len(found_list)}, не найдено {len(not_found_list)}, другая категория {len(partial_list)}")


if __name__ == "__main__":
    main()