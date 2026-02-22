"""
Парсер с категориями из исходного чек-листа
"""

import zipfile
from pathlib import Path
from lxml import etree
import re

from parser.models import ChecklistData, ChecklistItem, Category


class CategoryParser:
    """Парсер с определением категорий по заголовкам"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    # Категории из исходного чек-листа
    CATEGORY_HEADERS = {
        'Поверхности:': Category.SURFACE,
        'Бытовая техника:': Category.HOUSEHOLD_APPLIANCES,
        'Тепловое оборудование:': Category.THERMAL_EQUIPMENT,
        'Упаковочное оборудование:': Category.PACKAGING_EQUIPMENT,
        'Технологическое оборудование:': Category.TECH_EQUIPMENT,
        'Инвентарь, посуда и т.д.': Category.INVENTORY,
        'Моечный, уборочный инвентарь и оборудование:': Category.CLEANING_EQUIPMENT,
        'Посудомоечное оборудование:': Category.DISHWASHING_EQUIPMENT,
        'Холодильное оборудование:': Category.REFRIGERATION_EQUIPMENT,
        'Дозирующее оборудование': Category.DOSING_EQUIPMENT,
        'Сантехническое оборудование:': Category.PLUMBING,
        'Мебель:': Category.FURNITURE,
        'Офисная техника:': Category.OFFICE_EQUIPMENT,
        'Санитарный пост': Category.SANITARY_POST,
        'Многоразовые резиновые СИЗ': Category.PPE,
    }

    def __init__(self):
        self.current_category = Category.OTHER

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                # Находим все ячейки таблиц
                cells = root.xpath('.//w:tc', namespaces=self.NAMESPACES)
                print(f"📊 Найдено ячеек: {len(cells)}")

                for cell in cells:
                    self._process_cell(cell, data)

        # Группируем по категориям для вывода
        self._print_statistics(data)

        print(f"\n✅ Найдено элементов: {len(data.items)}")
        return data

    def _process_cell(self, cell, data: ChecklistData):
        """Обрабатывает одну ячейку таблицы"""

        # Получаем все текстовые элементы в ячейке
        text_elements = cell.xpath('.//w:t', namespaces=self.NAMESPACES)

        # Собираем полный текст ячейки
        full_text = ''.join([t.text for t in text_elements if t.text])

        # Проверяем, не является ли ячейка заголовком категории
        for header, category in self.CATEGORY_HEADERS.items():
            if header in full_text:
                print(f"\n📌 Найдена категория: {category.value}")
                self.current_category = category
                return

        # Если в ячейке есть ☒
        if '☒' in full_text:
            # Разбиваем по чек-боксам внутри ячейки
            parts = re.split(r'([☐☒])', full_text)

            i = 0
            while i < len(parts):
                if parts[i] == '☒':  # Нашли отмеченный
                    if i + 1 < len(parts):
                        # Берём текст до следующего чек-бокса или до конца ячейки
                        next_text = parts[i + 1]

                        # Ищем следующий чек-бокс в этом же тексте
                        next_checkbox_pos = -1
                        for j, char in enumerate(next_text):
                            if char in ['☐', '☒']:
                                next_checkbox_pos = j
                                break

                        if next_checkbox_pos != -1:
                            # Обрезаем до следующего чек-бокса
                            item_text = next_text[:next_checkbox_pos].strip()
                        else:
                            # Берём весь текст до конца ячейки
                            item_text = next_text.strip()

                        # Очищаем
                        item_text = re.sub(r'[^\w\s\-\(\)]', '', item_text)
                        item_text = re.sub(r'\s+', ' ', item_text).strip()

                        # Убираем одиночные буквы в конце
                        item_text = re.sub(r'\s+[А-Я]$', '', item_text)

                        if item_text and len(item_text) > 1:
                            print(f"  [{self.current_category.value}] ✅ {item_text}")
                            data.items.append(ChecklistItem(
                                name=item_text,
                                category=self.current_category,
                                checked=True
                            ))
                i += 1

    def _print_statistics(self, data: ChecklistData):
        """Выводит статистику по категориям"""
        print("\n📊 Статистика по категориям:")

        grouped = data.group_checked_by_category()
        for category, items in grouped.items():
            print(f"  {category.value}: {len(items)} элементов")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = CategoryParser()
    return parser.parse(file_path)