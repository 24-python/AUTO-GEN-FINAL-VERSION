"""
Парсер таблицы используемых средств из формата .docx
Извлекает: название средства, форму фасовки, концентрацию, метод разведения,
отмеченные типы поверхностей, примечание.
"""

import zipfile
from pathlib import Path
from lxml import etree
from typing import List, Dict, Optional
from dataclasses import dataclass, field

# Пространства имён Word
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# Типы поверхностей (индексы колонок)
SURFACE_TYPES = {
    4: "потолок",
    5: "стены",
    6: "полы",
    7: "трапы",
    8: "окна, зеркала",
    9: "сантехника",
    10: "оборудование",
    11: "рабочие поверхности",
    12: "локальные загрязнения",
    13: "посуда, производственный инвентарь",
    14: "уборочный инвентарь",
}

# Маппинг типа поверхности → категория в БД
SURFACE_TO_CATEGORY = {
    "потолок": "Поверхности",
    "стены": "Поверхности",
    "полы": "Поверхности",
    "трапы": "Поверхности",
    "окна, зеркала": "Поверхности",
    "сантехника": "Сантехническое оборудование",
    "оборудование": "Оборудование",  # объединяет: холодильное, тепловое, технологическое
    "рабочие поверхности": "Мебель",
    "локальные загрязнения": "Поверхности",
    "посуда, производственный инвентарь": "Инвентарь, посуда и т.д.",
    "уборочный инвентарь": "Моечный, уборочный инвентарь и оборудование",
}


@dataclass
class ProductEntry:
    """Одна строка из таблицы средств"""
    product_name: str
    packaging: str           # Форма фасовки
    concentration: str       # Концентрация
    application_method: str  # Метод разведения
    surfaces: List[str] = field(default_factory=list)  # Отмеченные типы поверхностей
    note: Optional[str] = None  # Примечание (последняя колонка)


def _get_text_from_cell(cell) -> str:
    """Извлекает текст из ячейки (без символов чек-боксов)"""
    texts = cell.findall('.//w:t', namespaces=NAMESPACES)
    result = ' '.join(t.text or '' for t in texts).strip()
    # Удаляем символы чек-боксов
    result = result.replace('☒', '').replace('☐', '').strip()
    return result


def _get_checkbox_state(cell) -> bool:
    """Определяет, отмечен ли чек-бокс в ячейке"""
    texts = cell.findall('.//w:t', namespaces=NAMESPACES)
    for t in texts:
        if t.text and '☒' in t.text:
            return True
    return False


def parse_products_table(file_path: str) -> List[ProductEntry]:
    """
    Парсит таблицу используемых средств из .docx файла.
    Возвращает список ProductEntry.
    """
    file_path = Path(file_path)
    entries = []

    with zipfile.ZipFile(file_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

    # Находим таблицу
    tables = root.findall('.//w:tbl', namespaces=NAMESPACES)
    if not tables:
        print("❌ Таблица не найдена в документе")
        return entries

    table = tables[0]
    rows = table.findall('.//w:tr', namespaces=NAMESPACES)

    if len(rows) < 3:
        print("❌ Слишком мало строк в таблице")
        return entries

    print(f"📊 Найдено строк: {len(rows)} (строки 1-2 — заголовки)")

    # Обрабатываем строки данных (начиная с 3-й)
    for r_idx in range(2, len(rows)):
        row = rows[r_idx]
        cells = row.findall('.//w:tc', namespaces=NAMESPACES)

        if len(cells) < 15:
            continue

        # Извлекаем данные
        product_name = _get_text_from_cell(cells[0])
        packaging = _get_text_from_cell(cells[1])
        concentration = _get_text_from_cell(cells[2])
        application_method = _get_text_from_cell(cells[3])
        note = _get_text_from_cell(cells[15]) if len(cells) > 15 else None

        if not product_name:
            continue

        # Определяем отмеченные поверхности
        surfaces = []
        for col_idx, surface_name in SURFACE_TYPES.items():
            if col_idx < len(cells):
                if _get_checkbox_state(cells[col_idx]):
                    surfaces.append(surface_name)

        entry = ProductEntry(
            product_name=product_name,
            packaging=packaging,
            concentration=concentration,
            application_method=application_method,
            surfaces=surfaces,
            note=note if note else None,
        )

        entries.append(entry)

        # Вывод для отладки
        print(f"\n📦 {product_name}")
        print(f"   Концентрация: {concentration}")
        print(f"   Метод: {application_method}")
        print(f"   Поверхности: {', '.join(surfaces) if surfaces else 'нет'}")
        if note:
            print(f"   Примечание: {note}")

    print(f"\n✅ Всего записей: {len(entries)}")
    return entries


def export_to_csv(entries: List[ProductEntry], output_path: str = "products_export.csv"):
    """Экспортирует распарсенные данные в CSV для ручной проверки"""
    import csv

    with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            'product_name', 'packaging', 'concentration', 'application_method',
            'surfaces', 'note', 'categories'
        ])

        for entry in entries:
            categories = [SURFACE_TO_CATEGORY.get(s, 'Неизвестно') for s in entry.surfaces]
            writer.writerow([
                entry.product_name,
                entry.packaging,
                entry.concentration,
                entry.application_method,
                ', '.join(entry.surfaces),
                entry.note or '',
                ', '.join(set(categories)),
            ])

    print(f"\n📁 Экспортировано в {output_path}")


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Парсер таблицы используемых средств")
    parser.add_argument("file", help="Путь к .docx файлу с таблицей средств")
    parser.add_argument("--csv", "-c", help="Экспорт в CSV", action="store_true")
    parser.add_argument("--output", "-o", default="products_export.csv", help="Путь для CSV экспорта")

    args = parser.parse_args()

    entries = parse_products_table(args.file)

    if args.csv and entries:
        export_to_csv(entries, args.output)


if __name__ == "__main__":
    main()