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
Добавлено: парсинг моющих и дезинфицирующих средств из раздела "Дополнительная информация".
Добавлено: парсинг специализированных моющих средств (пол/трапы, тепловое оборудование, стекло/зеркала/мониторы).
Добавлено: парсинг цвета инвентаря из выпадающего списка.
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

    # Категории помещений (в нижнем регистре для единообразия)
    ROOM_CATEGORIES = [
        'производственное', 'бытовое', 'складское', 'санитарное',
        'вспомогательное', 'моечное', 'техническое', 'офисное',
        'общего назначения'
    ]

    # Разрешённые цвета инвентаря (для фильтрации)
    ALLOWED_INVENTORY_COLORS = [
        'черный', 'красный', 'желтый', 'зеленый', 'синий', 'голубой'
    ]

    # Заглушки, которые могут быть выбраны вместо цвета инвентаря
    INVENTORY_COLOR_PLACEHOLDERS = {
        "Выберите элемент.", "Не выбрано", "Выберите элемент",
        "", None
    }

    # Паттерны, по которым определяем, что значение выпадающего списка – не название средства,
    # а концентрация или метод разведения
    NON_PRODUCT_PATTERNS = [
        r'%',                           # процент
        r'дозирующая система',        # дозирующая система ProMax и т.п.
        r'ручн(ой|ое|ая|ые)',         # ручной метод
        r'система',                    # система
        r'протирание',                 # метод уборки
        r'щётка',                      # метод
        r'ветошь',                     # метод
        r'губка',                      # метод
        r'погружение',                 # метод
        r'замачивание',                # метод
        r'орошение',                   # метод
        r'распыление',                 # метод
        r'протирка',                   # метод
        r'обработка',                  # общее слово
        r'раствор',                    # может быть в концентрации
    ]

    def __init__(self):
        self._category_cache = {}

    def _is_product_name(self, value: str) -> bool:
        """Проверяет, похоже ли значение на название средства (а не концентрацию/метод/заглушку)"""
        if not value:
            return False
        value_lower = value.lower().strip()
        # Если значение совпадает с заглушками – не продукт
        if value_lower in ('выберите элемент.', 'не выбрано', 'выберите элемент', ''):
            return False
        # Проверяем паттерны
        for pattern in self.NON_PRODUCT_PATTERNS:
            if re.search(pattern, value_lower):
                return False
        # Дополнительно: если содержит только цифры, символы процента, запятые, точки, дефисы – вероятно, концентрация
        if re.match(r'^[0-9.,%\-–\s]+$', value_lower):
            return False
        return True

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

        current_obj = None
        current_state = False
        text_parts = []

        for elem in elements:
            if elem['type'] == 'checkbox':
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
                current_state = elem['checked']
            elif elem['type'] == 'text':
                text_parts.append(elem['value'])

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

    def _parse_unicode_cell(self, cell, data: ChecklistData):
        """Парсит ячейку с Unicode-символами ☒/☐ (без SDT)"""
        texts = cell.xpath('.//w:t', namespaces=self.NAMESPACES)
        full_text = ''.join([t.text for t in texts if t.text])

        if not full_text or ('☒' not in full_text and '☐' not in full_text):
            return

        checkbox_count = full_text.count('☒') + full_text.count('☐')

        parts = re.split(r'([☐☒])', full_text)
        current_obj = None
        text_parts = []
        current_state = False

        for i in range(1, len(parts)):
            if parts[i] in ['☒', '☐']:
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
        """Добавляет объект в ChecklistData с правильной логикой"""
        base_name = obj['name']
        base_checked = obj['checked']
        modifiers = obj['modifiers']

        if not base_checked:
            return

        normalized_base = smart_normalize(base_name)

        if checkbox_count == 1:
            category = self._get_category_from_db(normalized_base)
            print(f"  [{category.value}] ✅ {base_name}")
            data.items.append(ChecklistItem(
                name=normalized_base,
                category=category,
                checked=True,
                markers=[]
            ))
            return

        checked_modifiers = [m for m in modifiers if m['checked']]

        if checked_modifiers:
            for mod in checked_modifiers:
                full_name = f"{base_name} {mod['name']}"
                normalized_full = smart_normalize(full_name)
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
        """Извлекает категорию помещения из выпадающего списка (dropDownList SDT)"""
        sdt_elements = root.xpath('.//w:sdt', namespaces=self.NAMESPACES)

        for sdt in sdt_elements:
            sdt_pr = sdt.find('.//w:sdtPr', namespaces=self.NAMESPACES)
            if sdt_pr is None:
                continue

            dropdown = sdt_pr.find('.//w:dropDownList', namespaces=self.NAMESPACES)
            if dropdown is None:
                continue

            sdt_content = sdt.find('.//w:sdtContent', namespaces=self.NAMESPACES)
            if sdt_content is not None:
                texts = sdt_content.findall('.//w:t', namespaces=self.NAMESPACES)
                value = ''.join(t.text or '' for t in texts).strip()
                if value:
                    # Приводим к нижнему регистру
                    value_lower = value.lower()
                    if value_lower in self.ROOM_CATEGORIES:
                        data.room_category = value_lower
                        print(f"  📋 Категория помещения (из выпадающего списка): «{value_lower}»")
                        return

        print("  ⚠️ Категория помещения не найдена (выпадающий список отсутствует в чек-листе)")

    def _parse_additional_info(self, root, data: ChecklistData):
        """
        Извлекает моющие и дезинфицирующие средства из раздела "Дополнительная информация",
        а также цвет инвентаря из выпадающего списка.
        Порядок списков фиксирован: 1 цвет + 15 средств (5 групп по 3).
        Пустые и заглушечные значения сохраняются, чтобы индексы групп не смещались.
        """
        sdt_elements = root.xpath('.//w:sdt', namespaces=self.NAMESPACES)

        all_values = []

        for sdt in sdt_elements:
            sdt_pr = sdt.find('.//w:sdtPr', namespaces=self.NAMESPACES)
            if sdt_pr is None:
                continue

            dropdown = sdt_pr.find('.//w:dropDownList', namespaces=self.NAMESPACES)
            if dropdown is None:
                continue

            sdt_content = sdt.find('.//w:sdtContent', namespaces=self.NAMESPACES)
            if sdt_content is not None:
                texts = sdt_content.findall('.//w:t', namespaces=self.NAMESPACES)
                value = ''.join(t.text or '' for t in texts).strip()
                # Пропускаем категории помещений
                if value.lower() in self.ROOM_CATEGORIES:
                    continue
                # Сохраняем значение как есть (может быть пустым)
                all_values.append(value if value else None)

        print(f"  📦 Найдено выпадающих списков: {len(all_values)}")

        # ------------------- ЦВЕТ ИНВЕНТАРЯ -------------------
        inventory_color = None
        if all_values:
            first_val = all_values[0]
            if first_val and first_val.lower() in self.ALLOWED_INVENTORY_COLORS:
                inventory_color = first_val.lower()
                print(f"  🎨 Найден цвет инвентаря: {inventory_color}")
            elif first_val in self.INVENTORY_COLOR_PLACEHOLDERS or first_val is None:
                print(f"  🎨 Цвет инвентаря не выбран (заглушка), будет 'промаркированный'")
            else:
                print(f"  ⚠️ Не удалось определить цвет инвентаря: '{first_val}', будет 'промаркированный'")
        data.inventory_color = inventory_color
        if inventory_color is None:
            print("  🎨 Цвет инвентаря не указан (будет 'промаркированный')")

        # ------------------- СРЕДСТВА -------------------
        # Оставшиеся значения (индексы 1..15) – 15 полей для средств
        dropdown_values = all_values[1:] if len(all_values) > 1 else []
        # Гарантируем ровно 15 элементов (дополняем None, если не хватает)
        while len(dropdown_values) < 15:
            dropdown_values.append(None)
        # Используем только первые 15 (на случай, если больше)
        dropdown_values = dropdown_values[:15]

        # Функция для извлечения тройки (без изменения индексов)
        def extract_triplet_from_slice(triplet):
            product = None
            concentration = None
            method = None
            # Ищем продукт в тройке
            for i, val in enumerate(triplet):
                if val and self._is_product_name(val):
                    product = val
                    # Концентрация и метод – следующие элементы, если они не продукты
                    if i + 1 < len(triplet) and triplet[i+1] and not self._is_product_name(triplet[i+1]):
                        concentration = triplet[i+1]
                    if i + 2 < len(triplet) and triplet[i+2] and not self._is_product_name(triplet[i+2]):
                        method = triplet[i+2]
                    break
            return product, concentration, method

        # Фиксированные группы
        groups = [
            ('cleaning', 0),      # общее моющее
            ('disinfection', 3),  # дезинфекция
            ('floor', 6),         # пол/трапы
            ('thermal', 9),       # тепловое
            ('glass', 12)         # стекло/зеркала
        ]

        for group_name, start_idx in groups:
            triplet = dropdown_values[start_idx:start_idx+3]
            product, concentration, method = extract_triplet_from_slice(triplet)
            setattr(data, f'{group_name}_product', product)
            setattr(data, f'{group_name}_concentration', concentration)
            setattr(data, f'{group_name}_method_text', method)

        # Вывод информации
        if data.cleaning_product:
            print(f"  🧴 Общее моющее средство: {data.cleaning_product}")
            print(f"     Концентрация: {data.cleaning_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.cleaning_method_text or 'не указан'}")
        else:
            print("  🧴 Общее моющее средство не выбрано")
        if data.disinfection_product:
            print(f"  🦠 Общее дезинфицирующее средство: {data.disinfection_product}")
            print(f"     Концентрация: {data.disinfection_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.disinfection_method_text or 'не указан'}")
        else:
            print("  🦠 Общее дезинфицирующее средство не выбрано")
        if data.floor_cleaning_product:
            print(f"  🧽 Моющее для пола/трапов: {data.floor_cleaning_product}")
            print(f"     Концентрация: {data.floor_cleaning_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.floor_cleaning_method_text or 'не указан'}")
        else:
            print("  🧽 Моющее для пола/трапов не выбрано")
        if data.thermal_cleaning_product:
            print(f"  🔥 Моющее для теплового оборудования: {data.thermal_cleaning_product}")
            print(f"     Концентрация: {data.thermal_cleaning_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.thermal_cleaning_method_text or 'не указан'}")
        else:
            print("  🔥 Моющее для теплового оборудования не выбрано")
        if data.glass_cleaning_product:
            print(f"  🪞 Моющее для стекол/зеркал/мониторов: {data.glass_cleaning_product}")
            print(f"     Концентрация: {data.glass_cleaning_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.glass_cleaning_method_text or 'не указан'}")
        else:
            print("  🪞 Моющее для стекол/зеркал/мониторов не выбрано")

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                self._parse_header(root, data)
                self._parse_room_category(root, data)
                self._parse_additional_info(root, data)

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
        if data.cleaning_product:
            print(f"   Общее моющее средство: {data.cleaning_product}")
        if data.disinfection_product:
            print(f"   Общее дезинфицирующее средство: {data.disinfection_product}")
        if data.floor_cleaning_product:
            print(f"   Моющее для пола/трапов: {data.floor_cleaning_product}")
        if data.thermal_cleaning_product:
            print(f"   Моющее для теплового оборудования: {data.thermal_cleaning_product}")
        if data.glass_cleaning_product:
            print(f"   Моющее для стекол/зеркал: {data.glass_cleaning_product}")
        if data.inventory_color:
            print(f"   Цвет инвентаря: {data.inventory_color}")
        else:
            print("   Цвет инвентаря: промаркированный")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = SDTChecklistParser()
    return parser.parse(file_path)