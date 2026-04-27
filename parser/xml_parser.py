"""
Парсер чек-листов с поддержкой SDT и Unicode чек-боксов.
Логика:
- 1 чек-бокс в ячейке → одиночный объект → в свою категорию
- 2+ чек-боксов → объект с модификаторами
  - Первый чек-бокс = базовый объект
  - Все остальные чек-боксы = модификаторы (сколько угодно)
  - Базовый ✅ + модификаторы ✅ → каждый модификатор в свою категорию
  - Базовый ✅ + модификаторы ⬜ → базовый в "Прочее"
  - Базовый ⬜ → ничего не добавляем

Добавлено: парсинг категории помещения из выпадающего списка (dropDownList).
"""

import zipfile
from pathlib import Path
from lxml import etree
import re
import sys

from parser.models import ChecklistData, ChecklistItem, Category

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Object as DBObject


# ============================================================
# НОРМАЛИЗАЦИЯ
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
    """Умная нормализация"""
    if 'потолок' in name.lower():
        return normalize_potolok_name(name)
    return normalize_name(name)


# ============================================================
# SDT ПАРСЕР
# ============================================================
class SDTChecklistParser:
    """Парсер для чек-листов с SDT и Unicode чек-боксами"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }

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
        self._category_cache = {}

    def _get_checkbox_state(self, sdt_element) -> bool:
        """Определяет состояние чек-бокса по символу внутри SDT"""
        texts = sdt_element.xpath('.//w:t', namespaces=self.NAMESPACES)
        for t in texts:
            if t.text:
                if '☒' in t.text:
                    return True
                if '☐' in t.text:
                    return False
        return False

    def _clean_text(self, text: str) -> str:
        """Очищает текст от лишних пробелов и символов"""
        if not text:
            return ""
        # Удаляем символы чек-боксов, если они попали в текст
        text = text.replace('☒', '').replace('☐', '')
        return ' '.join(text.split()).strip()

    def _get_category_from_db(self, normalized_name: str) -> Category:
        """Получает категорию из БД по normalized_name"""
        if normalized_name in self._category_cache:
            return self._category_cache[normalized_name]

        session = SessionLocal()
        try:
            obj = session.query(DBObject).filter(DBObject.normalized_name == normalized_name).first()
            if obj and obj.category:
                category_enum = self.DB_CATEGORY_TO_ENUM.get(obj.category.name, Category.OTHER)
                self._category_cache[normalized_name] = category_enum
                return category_enum
        except Exception:
            pass
        finally:
            session.close()

        return Category.OTHER

    def _parse_sdt_cell(self, cell, data: ChecklistData):
        """Парсит ячейку с SDT чек-боксами"""
        sdt_elements = cell.xpath('.//w:sdt', namespaces=self.NAMESPACES)
        if not sdt_elements:
            return

        elements = []
        checkbox_count = 0

        for para in cell.xpath('.//w:p', namespaces=self.NAMESPACES):
            for child in para.getchildren():
                tag = child.tag.split('}')[-1]

                if tag == 'sdt':
                    checkbox_count += 1
                    elements.append({
                        'type': 'checkbox',
                        'checked': self._get_checkbox_state(child)
                    })
                elif tag == 'r':
                    texts = child.xpath('.//w:t', namespaces=self.NAMESPACES)
                    for t in texts:
                        if t.text:
                            text = t.text.strip()
                            if text and text not in ['☒', '☐']:
                                elements.append({
                                    'type': 'text',
                                    'value': text
                                })

        if not elements:
            return

        # Собираем объект
        current_obj = None
        current_state = False
        text_parts = []

        for elem in elements:
            if elem['type'] == 'checkbox':
                # Сохраняем накопленный текст перед новым чек-боксом
                if text_parts:
                    full_text = self._clean_text(''.join(text_parts))
                    if full_text:
                        if current_obj is None:
                            current_obj = {
                                'name': full_text,
                                'checked': current_state,
                                'modifiers': []
                            }
                        else:
                            # Добавляем модификатор
                            current_obj['modifiers'].append({
                                'name': full_text,
                                'checked': current_state
                            })
                    text_parts = []
                current_state = elem['checked']
            elif elem['type'] == 'text':
                text_parts.append(elem['value'])

        # ВАЖНО: Сохраняем последний накопленный текст (последний модификатор)
        if text_parts:
            full_text = self._clean_text(''.join(text_parts))
            if full_text:
                if current_obj is None:
                    current_obj = {
                        'name': full_text,
                        'checked': current_state,
                        'modifiers': []
                    }
                else:
                    # Добавляем последний модификатор
                    current_obj['modifiers'].append({
                        'name': full_text,
                        'checked': current_state
                    })

        if current_obj:
            self._add_object_to_data(current_obj, checkbox_count, data)

    def _parse_unicode_cell(self, cell, data: ChecklistData):
        """Парсит ячейку с Unicode-символами ☒/☐ (без SDT)"""
        texts = cell.xpath('.//w:t', namespaces=self.NAMESPACES)
        full_text = ''.join([t.text for t in texts if t.text])

        if not full_text or ('☒' not in full_text and '☐' not in full_text):
            return

        # Считаем количество чек-боксов
        checkbox_count = full_text.count('☒') + full_text.count('☐')

        parts = re.split(r'([☐☒])', full_text)
        current_obj = None
        text_parts = []
        current_state = False

        for i in range(1, len(parts)):
            if parts[i] in ['☒', '☐']:
                # Сохраняем накопленный текст
                if text_parts:
                    full_text = self._clean_text(''.join(text_parts))
                    if full_text:
                        if current_obj is None:
                            current_obj = {
                                'name': full_text,
                                'checked': current_state,
                                'modifiers': []
                            }
                        else:
                            current_obj['modifiers'].append({
                                'name': full_text,
                                'checked': current_state
                            })
                    text_parts = []
                current_state = (parts[i] == '☒')
            else:
                text_parts.append(parts[i])

        # Сохраняем последний текст
        if text_parts:
            full_text = self._clean_text(''.join(text_parts))
            if full_text:
                if current_obj is None:
                    current_obj = {
                        'name': full_text,
                        'checked': current_state,
                        'modifiers': []
                    }
                else:
                    current_obj['modifiers'].append({
                        'name': full_text,
                        'checked': current_state
                    })

        if current_obj:
            self._add_object_to_data(current_obj, checkbox_count, data)

    def _add_object_to_data(self, obj: dict, checkbox_count: int, data: ChecklistData):
        """
        Добавляет объект в ChecklistData с правильной логикой.

        Логика:
        - checkbox_count == 1 → одиночный объект → в свою категорию (если отмечен)
        - checkbox_count >= 2 → объект с модификаторами
          - Базовый ✅ + модификаторы ✅ → каждый модификатор в свою категорию
          - Базовый ✅ + модификаторы ⬜ → базовый в "Прочее"
          - Базовый ⬜ → ничего не добавляем
        """
        base_name = obj['name']
        base_checked = obj['checked']
        modifiers = obj['modifiers']

        # Если базовый объект НЕ отмечен → ничего не добавляем
        if not base_checked:
            return

        # Нормализуем базовое имя
        normalized_base = smart_normalize(base_name)

        if checkbox_count == 1:
            # ОДИНОЧНЫЙ ОБЪЕКТ → в свою категорию
            category = self._get_category_from_db(normalized_base)
            print(f"  [{category.value}] ✅ {base_name}")
            data.items.append(ChecklistItem(
                name=normalized_base,
                category=category,
                checked=True,
                markers=[]
            ))
            return

        # ОБЪЕКТ С МОДИФИКАТОРАМИ (checkbox_count >= 2)
        checked_modifiers = [m for m in modifiers if m['checked']]

        if checked_modifiers:
            # Есть отмеченные модификаторы → добавляем каждый
            for mod in checked_modifiers:
                full_name = f"{base_name} {mod['name']}"
                normalized_full = smart_normalize(full_name)

                # Пытаемся получить категорию для полного имени
                mod_category = self._get_category_from_db(normalized_full)
                if mod_category == Category.OTHER:
                    mod_category = self._get_category_from_db(normalized_base)

                print(f"  [{mod_category.value}] ✅ {full_name}")
                data.items.append(ChecklistItem(
                    name=normalized_full,
                    category=mod_category,
                    checked=True,
                    markers=[mod['name']]
                ))
        else:
            # Нет отмеченных модификаторов → базовый в "Прочее"
            print(f"  [Прочее] ⚠️ {base_name} (модификаторы не отмечены)")
            data.items.append(ChecklistItem(
                name=normalized_base,
                category=Category.OTHER,
                checked=True,
                markers=[]
            ))

    def _parse_header(self, root, data: ChecklistData):
        """Парсит предприятие и участок"""
        first_row = root.xpath('.//w:tr[1]', namespaces=self.NAMESPACES)
        if first_row:
            cells = first_row[0].xpath('.//w:tc', namespaces=self.NAMESPACES)
            if len(cells) >= 2:
                texts = cells[0].xpath('.//w:t', namespaces=self.NAMESPACES)
                if texts:
                    data.enterprise = ' '.join([t.text for t in texts if t.text]).strip()
                texts = cells[1].xpath('.//w:t', namespaces=self.NAMESPACES)
                if texts:
                    data.room_name = ' '.join([t.text for t in texts if t.text]).strip()

    def _parse_room_category(self, root, data: ChecklistData):
        """
        Извлекает категорию помещения из выпадающего списка (dropDownList SDT) в чек-листе.
        Ищет SDT с тегом w:dropDownList и читает выбранное значение.
        """
        # Ищем все SDT элементы в документе
        sdt_elements = root.xpath('.//w:sdt', namespaces=self.NAMESPACES)

        for sdt in sdt_elements:
            sdt_pr = sdt.find('.//w:sdtPr', namespaces=self.NAMESPACES)
            if sdt_pr is None:
                continue

            # Проверяем, есть ли выпадающий список в этом SDT
            dropdown = sdt_pr.find('.//w:dropDownList', namespaces=self.NAMESPACES)
            if dropdown is None:
                continue

            # Нашли выпадающий список — извлекаем выбранное значение
            sdt_content = sdt.find('.//w:sdtContent', namespaces=self.NAMESPACES)
            if sdt_content is not None:
                texts = sdt_content.findall('.//w:t', namespaces=self.NAMESPACES)
                value = ''.join(t.text or '' for t in texts).strip()
                if value:
                    data.room_category = value
                    print(f"  📋 Категория помещения (из выпадающего списка): «{value}»")
                    return

        print("  ⚠️ Категория помещения не найдена (выпадающий список отсутствует в чек-листе)")

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                # Парсим заголовок (предприятие и участок)
                self._parse_header(root, data)

                # Парсим категорию помещения из выпадающего списка
                self._parse_room_category(root, data)

                # Парсим объекты (чек-боксы)
                cells = root.xpath('.//w:tc', namespaces=self.NAMESPACES)
                print(f"📊 Найдено ячеек: {len(cells)}")

                for cell in cells:
                    sdt_elements = cell.xpath('.//w:sdt', namespaces=self.NAMESPACES)
                    if sdt_elements:
                        self._parse_sdt_cell(cell, data)
                    else:
                        self._parse_unicode_cell(cell, data)

        self._print_statistics(data)
        print(f"\n✅ Найдено элементов: {len(data.items)}")

        return data

    def _print_statistics(self, data: ChecklistData):
        """Выводит статистику"""
        print("\n📊 Статистика по категориям:")
        grouped = data.group_checked_by_category()

        main_total = 0
        other_total = 0

        for category, items in grouped.items():
            if category == Category.OTHER:
                other_total = len(items)
            else:
                main_total += len(items)
            print(f"  {category.value}: {len(items)} элементов")
            for item in items[:3]:
                marker_info = f" [модификатор: {item.markers[0]}]" if item.markers else ""
                print(f"    • {item.name}{marker_info}")
            if len(items) > 3:
                print(f"    ... и ещё {len(items) - 3}")

        print(f"\n📋 ИТОГО:")
        print(f"   В основных категориях: {main_total}")
        print(f"   В Прочее: {other_total}")
        if data.room_category:
            print(f"   Категория помещения: {data.room_category}")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = SDTChecklistParser()
    return parser.parse(file_path)