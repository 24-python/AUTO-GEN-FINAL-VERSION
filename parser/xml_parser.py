"""
Парсер с определением категорий по базе данных
"""

import zipfile
from pathlib import Path
from lxml import etree
import re
import sys

from parser.models import ChecklistData, ChecklistItem, Category

# Добавляем корень проекта в путь для импорта БД
sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Object as DBObject


class CategoryParser:
    """Парсер с определением категорий из базы данных"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }

    # Составные позиции, требующие специальной обработки (без скобок, соответствует БД)
    COMPOUND_PATTERNS = [
        {
            'base': r'потолок\s*\([^)]+\)',
            'modifiers': [
                {'marker': 'П', 'name': 'П', 'format': '{base} {mod}'},
                {'marker': 'ОК', 'name': 'ОК', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'вытяжные\s+зонты',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} {mod}'},
                {'marker': 'А', 'name': 'А', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'формы\s+для\s+выпечки',
            'modifiers': [
                {'marker': 'С', 'name': 'С', 'format': '{base} {mod}'},
                {'marker': 'Н', 'name': 'Н', 'format': '{base} {mod}'},
                {'marker': 'А', 'name': 'А', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'листы\s+для\s+выпечки',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} {mod}'},
                {'marker': 'А', 'name': 'А', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'съёмные\s+детали\s+оборудования',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} {mod}'},
                {'marker': 'А', 'name': 'А', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'ПММ',
            'category': Category.DISHWASHING_EQUIPMENT,
            'modifiers': [
                {'marker': 'купольная', 'name': 'купольная', 'format': '{base} {mod}'},
                {'marker': 'туннельная', 'name': 'туннельная', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'камеры',
            'modifiers': [
                {'marker': r'холд\.?', 'name': 'холд.', 'format': '{base} {mod}'},
                {'marker': r'мороз\.?', 'name': 'мороз.', 'format': '{base} {mod}'},
                {'marker': r'шок\.?\s+замор\.?', 'name': 'шок. замор.', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'плиты',
            'modifiers': [
                {'marker': r'индук\.?', 'name': 'индук.', 'format': '{base} {mod}'},
                {'marker': r'элек\.?', 'name': 'элек.', 'format': '{base} {mod}'},
                {'marker': r'газ\.?', 'name': 'газ.', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'производственные\s+столы',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} {mod}'},
                {'marker': 'Д', 'name': 'Д', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'весы',
            'modifiers': [
                {'marker': r'напольные', 'name': 'напольные', 'format': '{base} {mod}'},
                {'marker': r'настольные', 'name': 'настольные', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'просеиватели',
            'modifiers': [
                {'marker': 'мука', 'name': 'мука', 'format': '{base} {mod}'},
                {'marker': 'сахар', 'name': 'сахар', 'format': '{base} {mod}'}
            ]
        }
    ]

    # Маппинг названий категорий из БД в Enum Category
    DB_CATEGORY_TO_ENUM = {
        "Поверхности": Category.SURFACE,
        "Сантехническое оборудование": Category.PLUMBING,
        "Санитарный пост": Category.SANITARY_POST,
        "Мебель": Category.FURNITURE,
        "Офисная техника": Category.OFFICE_EQUIPMENT,
        "Многоразовые резиновые СИЗ": Category.PPE,
        "Бытовая техника": Category.HOUSEHOLD_APPLIANCES,
        "Инвентарь, посуда и т.д.": Category.INVENTORY,
        "Моечный, уборочный инвентарь и оборудование": Category.CLEANING_EQUIPMENT,
        "Посудомоечное оборудование": Category.DISHWASHING_EQUIPMENT,
        "Холодильное оборудование": Category.REFRIGERATION_EQUIPMENT,
        "Дозирующее оборудование": Category.DOSING_EQUIPMENT,
        "Тепловое оборудование": Category.THERMAL_EQUIPMENT,
        "Технологическое оборудование": Category.TECH_EQUIPMENT,
        "Упаковочное оборудование": Category.PACKAGING_EQUIPMENT,
    }

    def __init__(self):
        self.current_category = Category.OTHER
        self._category_cache = {}  # Кэш для ускорения {имя_объекта: категория}

    def _get_category_from_db(self, item_name: str) -> Category:
        """Получает категорию из БД по точному имени объекта"""

        # Проверяем кэш
        if item_name in self._category_cache:
            return self._category_cache[item_name]

        session = SessionLocal()
        try:
            obj = session.query(DBObject).filter(DBObject.name == item_name).first()
            if obj and obj.category:
                category_enum = self.DB_CATEGORY_TO_ENUM.get(obj.category.name, Category.OTHER)
                self._category_cache[item_name] = category_enum
                return category_enum
        except Exception as e:
            print(f"⚠️ Ошибка при запросе к БД: {e}")
        finally:
            session.close()

        return Category.OTHER

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))
        self.current_category = Category.OTHER

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                self._parse_header_from_first_row(root, data)

                cells = root.xpath('.//w:tc', namespaces=self.NAMESPACES)
                print(f"📊 Найдено ячеек: {len(cells)}")

                for cell in cells:
                    self._process_cell(cell, data)

        self._print_statistics(data)
        print(f"\n✅ Найдено элементов: {len(data.items)}")

        return data

    def _parse_header_from_first_row(self, root, data: ChecklistData):
        """Парсит предприятие и участок из первой строки"""
        first_row = root.xpath('.//w:tr[1]', namespaces=self.NAMESPACES)

        if first_row:
            cells = first_row[0].xpath('.//w:tc', namespaces=self.NAMESPACES)

            if len(cells) >= 2:
                enterprise_text = self._get_cell_text(cells[0])
                if enterprise_text and not data.enterprise:
                    data.enterprise = enterprise_text.strip()
                    print(f"🏭 Найдено предприятие: {data.enterprise}")

                room_text = self._get_cell_text(cells[1])
                if room_text and not data.room_name:
                    data.room_name = room_text.strip()
                    print(f"🏢 Найден участок: {data.room_name}")

    def _get_cell_text(self, cell) -> str:
        """Извлекает текст из ячейки"""
        texts = []
        for text_elem in cell.xpath('.//w:t', namespaces=self.NAMESPACES):
            if text_elem.text:
                texts.append(text_elem.text)
        return ' '.join(texts)

    def _is_checkbox_before(self, text: str, position: int) -> bool:
        """Проверяет, есть ли символ чек-бокса перед указанной позицией"""
        pos = position - 1
        while pos >= 0 and text[pos] in [' ', '\t', ':', ';', ',', '.', '-', '—', '–']:
            pos -= 1
        return pos >= 0 and text[pos] == '☒'

    def _process_compound_cell(self, full_text: str, data: ChecklistData) -> bool:
        """Специальная обработка составных ячеек"""

        for pattern in self.COMPOUND_PATTERNS:
            base_match = re.search(pattern['base'], full_text, re.IGNORECASE)
            if not base_match:
                continue

            base_name = base_match.group(0)
            base_start = base_match.start()
            base_checked = self._is_checkbox_before(full_text, base_start)

            checked_modifiers = []

            for modifier in pattern['modifiers']:
                for mod_match in re.finditer(modifier['marker'], full_text, re.IGNORECASE):
                    mod_pos = mod_match.start()
                    if mod_pos < base_match.end():
                        continue
                    if mod_pos > 0 and full_text[mod_pos - 1].isalpha() and mod_pos - 1 >= base_match.end():
                        continue
                    if self._is_checkbox_before(full_text, mod_pos):
                        checked_modifiers.append(modifier)
                        break

            if checked_modifiers:
                for modifier in checked_modifiers:
                    if 'format' in modifier:
                        item_name = modifier['format'].format(
                            base=base_name,
                            mod=modifier['name']
                        )
                    else:
                        item_name = f"{base_name} {modifier['name']}"

                    item_name = re.sub(r'\s+', ' ', item_name).strip()

                    # Получаем категорию из БД
                    if 'category' in pattern:
                        category = pattern['category']
                    else:
                        category = self._get_category_from_db(item_name)

                    print(f"  [{category.value}] ✅ {item_name}")
                    data.items.append(ChecklistItem(
                        name=item_name,
                        category=category,
                        checked=True
                    ))
                return True

            elif base_checked:
                if 'category' in pattern:
                    category = pattern['category']
                else:
                    category = self._get_category_from_db(base_name)

                print(f"  [{category.value}] ✅ {base_name}")
                data.items.append(ChecklistItem(
                    name=base_name,
                    category=category,
                    checked=True
                ))
                return True

        return False

    def _process_cell(self, cell, data: ChecklistData):
        """Обрабатывает одну ячейку"""

        text_elements = cell.xpath('.//w:t', namespaces=self.NAMESPACES)
        full_text = ''.join([t.text for t in text_elements if t.text])

        if '☒' in full_text:
            if self._process_compound_cell(full_text, data):
                return

            parts = re.split(r'([☐☒])', full_text)

            i = 0
            while i < len(parts):
                if parts[i] == '☒':
                    if i + 1 < len(parts):
                        next_text = parts[i + 1]

                        next_checkbox_pos = -1
                        for j, char in enumerate(next_text):
                            if char in ['☐', '☒']:
                                next_checkbox_pos = j
                                break

                        if next_checkbox_pos != -1:
                            item_text = next_text[:next_checkbox_pos].strip()
                        else:
                            item_text = next_text.strip()

                        item_text = re.sub(r'[^\w\s\-\(\)/]', '', item_text)  # слеш разрешен
                        item_text = re.sub(r'\s+', ' ', item_text).strip()
                        item_text = re.sub(r'\s+[А-Я]$', '', item_text)

                        if item_text and len(item_text) > 1:
                            category = self._get_category_from_db(item_text)
                            print(f"  [{category.value}] ✅ {item_text}")
                            data.items.append(ChecklistItem(
                                name=item_text,
                                category=category,
                                checked=True
                            ))
                i += 1

    def _print_statistics(self, data: ChecklistData):
        """Выводит статистику по категориям"""
        print("\n📊 Статистика по категориям:")

        grouped = data.group_checked_by_category()
        for category, items in grouped.items():
            print(f"  {category.value}: {len(items)} элементов")
            for item in items[:3]:
                print(f"    • {item.name}")
            if len(items) > 3:
                print(f"    ... и ещё {len(items) - 3}")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = CategoryParser()
    return parser.parse(file_path)