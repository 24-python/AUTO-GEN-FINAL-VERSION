#!/usr/bin/env python3
"""
compare_parser_with_db.py

Сравнивает normalized_name из чек-листа с БД.
Показывает:
- Какие объекты есть в БД, а какие нет
- Какие категории в БД соответствуют объектам
- Расхождения в категориях
"""

import zipfile
import re
import sys
from pathlib import Path
from collections import defaultdict
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

# Импортируем БД
from db.database import SessionLocal
from db.models import Object, Category

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


# ============================================================
# НОРМАЛИЗАЦИЯ (ТОЧНО ТАКАЯ ЖЕ, КАК В ПАРСЕРЕ)
# ============================================================
def normalize_name(name: str) -> str:
    """Нормализует имя для поиска в БД"""
    if not name:
        return ""
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


def normalize_potolok_name(name: str) -> str:
    """Специальная нормализация для потолка"""
    if not name:
        return ""
    normalized = re.sub(r'\([^)]*\)', '()', name)
    normalized = re.sub(r'[^\w\s\-\(\)]', '', normalized)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


def smart_normalize(name: str) -> str:
    """Умная нормализация (как в парсере)"""
    if 'потолок' in name.lower():
        return normalize_potolok_name(name)
    return normalize_name(name)


# ============================================================
# ФУНКЦИЯ ОПРЕДЕЛЕНИЯ СОСТОЯНИЯ ЧЕК-БОКСА
# ============================================================
def get_checkbox_state(sdt_element) -> bool:
    """Определяет состояние чек-бокса по символу внутри SDT"""
    texts = sdt_element.xpath('.//w:t', namespaces=NAMESPACES)
    for t in texts:
        if t.text:
            if '☒' in t.text:
                return True
            if '☐' in t.text:
                return False
    return False


def clean_text(text: str) -> str:
    """Очищает текст"""
    if not text:
        return ""
    text = ' '.join(text.split())
    return text.strip()


# ============================================================
# ПАРСИНГ ЧЕК-ЛИСТА
# ============================================================
def parse_checklist_for_comparison(docx_path: str) -> list:
    """Парсит чек-лист и возвращает все объекты с их normalized_name"""

    all_objects = []

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())
            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)

            for cell in cells:
                # Парсим SDT
                sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
                if not sdt_elements:
                    continue

                elements = []

                for para in cell.xpath('.//w:p', namespaces=NAMESPACES):
                    for child in para.getchildren():
                        tag = child.tag.split('}')[-1]

                        if tag == 'sdt':
                            elements.append({
                                'type': 'checkbox',
                                'checked': get_checkbox_state(child)
                            })
                        elif tag == 'r':
                            texts = child.xpath('.//w:t', namespaces=NAMESPACES)
                            for t in texts:
                                if t.text:
                                    text = t.text.strip()
                                    if text and text not in ['☒', '☐']:
                                        elements.append({
                                            'type': 'text',
                                            'value': text
                                        })

                if not elements:
                    continue

                current_obj = None
                current_state = False
                text_parts = []

                for elem in elements:
                    if elem['type'] == 'checkbox':
                        if text_parts:
                            full_text = clean_text(''.join(text_parts))
                            if full_text:
                                if current_obj is None:
                                    current_obj = {
                                        'name': full_text,
                                        'checked': current_state,
                                        'normalized': smart_normalize(full_text),
                                        'modifiers': []
                                    }
                                    all_objects.append(current_obj)
                                else:
                                    mod_full_name = f"{current_obj['name']} {full_text}"
                                    current_obj['modifiers'].append({
                                        'name': full_text,
                                        'full_name': mod_full_name,
                                        'checked': current_state,
                                        'normalized': smart_normalize(mod_full_name)
                                    })
                            text_parts = []
                        current_state = elem['checked']
                    elif elem['type'] == 'text':
                        text_parts.append(elem['value'])

                # Последний текст
                if text_parts and current_obj is None:
                    full_text = clean_text(''.join(text_parts))
                    if full_text:
                        current_obj = {
                            'name': full_text,
                            'checked': current_state,
                            'normalized': smart_normalize(full_text),
                            'modifiers': []
                        }
                        all_objects.append(current_obj)

    return all_objects


# ============================================================
# СРАВНЕНИЕ С БД
# ============================================================
def compare_with_db(checklist_objects: list, output_path: str):
    """Сравнивает объекты из чек-листа с БД"""

    session = SessionLocal()

    # Загружаем все объекты из БД
    db_objects = session.query(Object).all()
    db_categories = {cat.id: cat.name for cat in session.query(Category).all()}

    # Создаем словарь БД: normalized_name -> (display_name, category_name)
    db_dict = {}
    for obj in db_objects:
        cat_name = db_categories.get(obj.category_id, "НЕИЗВЕСТНО")
        db_dict[obj.normalized_name] = {
            'display_name': obj.display_name,
            'category': cat_name,
            'id': obj.id
        }

    # Собираем все normalized_name из чек-листа
    checklist_names = set()
    for obj in checklist_objects:
        checklist_names.add(obj['normalized'])
        for mod in obj['modifiers']:
            checklist_names.add(mod['normalized'])

    # Находим совпадения и расхождения
    in_both = checklist_names & set(db_dict.keys())
    only_in_checklist = checklist_names - set(db_dict.keys())
    only_in_db = set(db_dict.keys()) - checklist_names

    # Группируем по категориям из БД
    by_db_category = defaultdict(list)
    for name in in_both:
        cat = db_dict[name]['category']
        by_db_category[cat].append({
            'normalized': name,
            'display': db_dict[name]['display_name']
        })

    # Сохраняем в TXT
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("=" * 100 + "\n")
        f.write("СРАВНЕНИЕ ОБЪЕКТОВ ЧЕК-ЛИСТА С БАЗОЙ ДАННЫХ\n")
        f.write("=" * 100 + "\n\n")

        # Статистика
        f.write("СТАТИСТИКА\n")
        f.write("-" * 50 + "\n")
        f.write(f"Всего объектов в чек-листе:     {len(checklist_names)}\n")
        f.write(f"Всего объектов в БД:             {len(db_dict)}\n")
        f.write(f"Есть в обоих:                    {len(in_both)}\n")
        f.write(f"Только в чек-листе (нет в БД):   {len(only_in_checklist)}\n")
        f.write(f"Только в БД (нет в чек-листе):   {len(only_in_db)}\n")
        f.write(f"Процент покрытия:                {len(in_both) / len(checklist_names) * 100:.1f}%\n")
        f.write("\n")

        # Объекты, которых НЕТ в БД (самое важное!)
        if only_in_checklist:
            f.write("=" * 100 + "\n")
            f.write(f"⚠️ ОБЪЕКТЫ, КОТОРЫХ НЕТ В БД ({len(only_in_checklist)} шт.)\n")
            f.write("=" * 100 + "\n\n")
            f.write("Эти объекты нужно ДОБАВИТЬ в БД:\n\n")

            for name in sorted(only_in_checklist):
                # Находим оригинальное имя из чек-листа
                original_name = name
                for obj in checklist_objects:
                    if obj['normalized'] == name:
                        original_name = obj['name']
                        break
                    for mod in obj['modifiers']:
                        if mod['normalized'] == name:
                            original_name = mod['full_name']
                            break

                f.write(f"   normalized: {name}\n")
                f.write(f"   original:   {original_name}\n")
                f.write(f"   → нужно определить категорию и display_name\n\n")

        # Объекты, которые есть в обоих (с категориями из БД)
        f.write("\n" + "=" * 100 + "\n")
        f.write(f"✅ ОБЪЕКТЫ, КОТОРЫЕ ЕСТЬ В БД (с категориями) ({len(in_both)} шт.)\n")
        f.write("=" * 100 + "\n\n")

        for cat in sorted(by_db_category.keys()):
            items = by_db_category[cat]
            f.write(f"\n{cat} ({len(items)} шт.):\n")
            f.write("-" * 60 + "\n")
            for item in sorted(items, key=lambda x: x['normalized']):
                f.write(f"   {item['normalized']:<40} → {item['display']}\n")

        # Сводка по категориям (таблица)
        f.write("\n" + "=" * 100 + "\n")
        f.write("СВОДКА ПО КАТЕГОРИЯМ (из БД)\n")
        f.write("=" * 100 + "\n\n")

        f.write(f"{'Категория':<40} {'Найдено':<10} {'% от чек-листа':<15}\n")
        f.write("-" * 70 + "\n")

        total_in_checklist = len(checklist_names)
        for cat in sorted(by_db_category.keys()):
            count = len(by_db_category[cat])
            percent = count / total_in_checklist * 100 if total_in_checklist > 0 else 0
            f.write(f"{cat:<40} {count:<10} {percent:.1f}%\n")

        # Объекты только в БД (лишние)
        if only_in_db:
            f.write("\n" + "=" * 100 + "\n")
            f.write(f"📦 ОБЪЕКТЫ ТОЛЬКО В БД (нет в чек-листе) ({len(only_in_db)} шт.)\n")
            f.write("=" * 100 + "\n\n")

            by_cat_db = defaultdict(list)
            for name in only_in_db:
                cat = db_dict[name]['category']
                by_cat_db[cat].append({
                    'normalized': name,
                    'display': db_dict[name]['display_name']
                })

            for cat in sorted(by_cat_db.keys()):
                items = by_cat_db[cat]
                f.write(f"\n{cat} ({len(items)} шт.):\n")
                for item in sorted(items, key=lambda x: x['normalized'])[:30]:
                    f.write(f"   {item['normalized']}\n")
                if len(items) > 30:
                    f.write(f"   ... и ещё {len(items) - 30}\n")

    session.close()

    print(f"\n✅ Результаты сохранены в: {output_path}")

    # Вывод в консоль
    print(f"\n{'=' * 80}")
    print(f"📊 КРАТКАЯ СТАТИСТИКА")
    print(f"{'=' * 80}")
    print(f"Объектов в чек-листе:     {len(checklist_names)}")
    print(f"Объектов в БД:             {len(db_dict)}")
    print(f"Есть в обоих:              {len(in_both)}")
    print(f"❌ НЕТ В БД:               {len(only_in_checklist)}")
    print(f"Процент покрытия:          {len(in_both) / len(checklist_names) * 100:.1f}%")

    if only_in_checklist:
        print(f"\n⚠️ ПЕРВЫЕ 20 ОБЪЕКТОВ, КОТОРЫХ НЕТ В БД:")
        for name in sorted(only_in_checklist)[:20]:
            print(f"   - {name}")
        if len(only_in_checklist) > 20:
            print(f"   ... и ещё {len(only_in_checklist) - 20}")


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу чек-листа")
    parser_arg.add_argument("--output", "-o", default="db_comparison.txt", help="Выходной TXT файл")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    print(f"\n{'=' * 80}")
    print(f"🔧 СРАВНЕНИЕ ЧЕК-ЛИСТА С БАЗОЙ ДАННЫХ")
    print(f"{'=' * 80}")
    print(f"📄 Чек-лист: {file_path}")
    print(f"🗄️ БД: tech_cards.db\n")

    # Парсим чек-лист
    print("📊 Парсинг чек-листа...")
    checklist_objects = parse_checklist_for_comparison(file_path)
    print(f"   Найдено объектов: {len(checklist_objects)}")

    # Сравниваем с БД
    print("\n🔍 Сравнение с БД...")
    compare_with_db(checklist_objects, args.output)

    print(f"\n{'=' * 80}")
    print(f"✅ СРАВНЕНИЕ ЗАВЕРШЕНО")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()