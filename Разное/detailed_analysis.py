#!/usr/bin/env python3
"""
detailed_analysis.py

Детальный анализ всех объектов по категориям.
Показывает содержимое Unicode-ячеек.
"""

import zipfile
import sys
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def extract_objects_from_unicode_cell(cell) -> list:
    """Извлекает объекты из ячейки с Unicode-символами"""
    texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
    full_text = ''.join([t.text for t in texts if t.text])

    if not full_text:
        return []

    objects = []

    # Ищем все чек-боксы и текст после них
    import re
    parts = re.split(r'([☐☒])', full_text)

    for i in range(1, len(parts), 2):
        if i < len(parts):
            checkbox = parts[i]
            text = parts[i + 1].strip() if i + 1 < len(parts) else ""

            # Обрезаем до следующего чек-бокса или до конца
            next_cb = text.find('☒')
            if next_cb == -1:
                next_cb = text.find('☐')
            if next_cb != -1:
                text = text[:next_cb].strip()

            if text:
                objects.append({
                    'name': text,
                    'checked': (checkbox == '☒')
                })

    return objects


def analyze_document_detailed(docx_path: str):
    """Детальный анализ документа"""

    print(f"\n{'=' * 80}")
    print(f"🔧 ДЕТАЛЬНЫЙ АНАЛИЗ ЧЕК-ЛИСТА")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {docx_path}\n")

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)
            print(f"📊 Всего ячеек в документе: {len(cells)}")

            # Категории и их объекты
            categories = {}
            current_category = "БЕЗ КАТЕГОРИИ"

            # Список всех возможных категорий (из чек-листа)
            all_category_names = [
                'Поверхности', 'Бытовая техника', 'Тепловое оборудование',
                'Упаковочное оборудование', 'Инвентарь', 'Технологическое оборудование',
                'Холодильное оборудование', 'Санитарный пост', 'Мебель',
                'Офисная техника', 'Многоразовые резиновые СИЗ', 'Моечный',
                'Посудомоечное', 'Сантехническое', 'Дозирующее'
            ]

            unicode_cells_content = []

            for idx, cell in enumerate(cells):
                # Получаем текст ячейки
                texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
                full_text = ''.join([t.text for t in texts if t.text])

                # Определяем категорию
                for cat_name in all_category_names:
                    if cat_name in full_text and ':' in full_text:
                        current_category = cat_name
                        if current_category not in categories:
                            categories[current_category] = {
                                'sdt_objects': [],
                                'unicode_objects': []
                            }
                        break

                # Проверяем SDT
                sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
                has_sdt = len(sdt_elements) > 0

                # Проверяем Unicode
                has_unicode = '☒' in full_text or '☐' in full_text

                # Если есть Unicode, но нет SDT - это "чистый" Unicode
                if has_unicode and not has_sdt:
                    objects = extract_objects_from_unicode_cell(cell)
                    unicode_cells_content.append({
                        'cell_idx': idx,
                        'category': current_category,
                        'text': full_text[:100],
                        'objects': objects
                    })

                    if current_category not in categories:
                        categories[current_category] = {
                            'sdt_objects': [],
                            'unicode_objects': []
                        }

                    for obj in objects:
                        categories[current_category]['unicode_objects'].append(obj)

    # Вывод результатов
    print(f"\n{'=' * 80}")
    print(f"📊 ЯЧЕЙКИ ТОЛЬКО С UNICODE-СИМВОЛАМИ ({len(unicode_cells_content)} шт.)")
    print(f"{'=' * 80}")

    for cell_info in unicode_cells_content:
        print(f"\n📍 Ячейка #{cell_info['cell_idx']} [{cell_info['category']}]")
        print(f"   Текст: '{cell_info['text']}'")
        for obj in cell_info['objects']:
            status = "✅" if obj['checked'] else "⬜"
            print(f"      {status} {obj['name']}")

    print(f"\n{'=' * 80}")
    print(f"📊 ПОЛНАЯ СТАТИСТИКА ПО ВСЕМ КАТЕГОРИЯМ")
    print(f"{'=' * 80}")

    total_objects = 0
    total_checked = 0

    for cat_name in sorted(categories.keys()):
        cat_data = categories[cat_name]
        sdt_count = len(cat_data['sdt_objects'])
        unicode_count = len(cat_data['unicode_objects'])
        cat_total = sdt_count + unicode_count

        total_objects += cat_total

        print(f"\n📁 {cat_name}:")
        print(f"   SDT объектов: {sdt_count}")
        print(f"   Unicode объектов: {unicode_count}")
        print(f"   ВСЕГО: {cat_total}")

    print(f"\n{'=' * 40}")
    print(f"🎯 ОБЩЕЕ КОЛИЧЕСТВО ОБЪЕКТОВ: {total_objects}")

    # Показываем все Unicode-объекты списком
    print(f"\n{'=' * 80}")
    print(f"📋 СПИСОК ВСЕХ UNICODE-ОБЪЕКТОВ")
    print(f"{'=' * 80}")

    for cat_name in sorted(categories.keys()):
        unicode_objs = categories[cat_name]['unicode_objects']
        if unicode_objs:
            print(f"\n{cat_name}:")
            for obj in unicode_objs:
                status = "✅" if obj['checked'] else "⬜"
                print(f"   {status} {obj['name']}")


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    analyze_document_detailed(file_path)


if __name__ == "__main__":
    main()