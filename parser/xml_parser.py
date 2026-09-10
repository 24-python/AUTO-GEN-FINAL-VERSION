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
Добавлено: парсинг специализированных моющих средств (пол/трапы, технологическое оборудование, тепловое оборудование, стекло/зеркала/мониторы).
Добавлено: парсинг цвета инвентаря из выпадающего списка.
Доработано: категории помещений и цвета инвентаря загружаются из БД (справочники).
Добавлено: парсинг зональных исполнителей (поверхности выше 2 м, до 2 м, оборудование).

ИСПРАВЛЕНИЯ БЕЗОПАСНОСТИ:
- Защита от XXE (XML External Entity): отключены DTD, внешние сущности и сетевые запросы
  при парсинге XML через lxml.
"""

import zipfile
from pathlib import Path
from lxml import etree
import re
import sys

from parser.models import ChecklistData, ChecklistItem, Category

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Object as DBObject, RoomCategory, InventoryColor


# ============================================================
# БЕЗОПАСНЫЙ ПАРСЕР XML (защита от XXE)
# ============================================================
def _create_safe_xml_parser() -> etree.XMLParser:
    """
    Создаёт XML-парсер с отключёнными DTD, внешними сущностями и сетью.
    Защита от XXE (XML External Entity): чтения локальных файлов,
    SSRF-атак и DoS через «billion laughs».
    """
    return etree.XMLParser(
        load_dtd=False,           # не загружать DTD
        no_network=True,          # запретить сетевые запросы
        resolve_entities=False,   # не разрешать внешние сущности
        huge_tree=False,          # защита от гигантских деревьев (DoS)
        dtd_validation=False,     # не валидировать по DTD
        attribute_defaults=False, # не подставлять значения из DTD
    )


def _safe_fromstring(xml_content: bytes) -> etree._Element:
    """Безопасный парсинг XML из байтов (защита от XXE)."""
    parser = _create_safe_xml_parser()
    return etree.fromstring(xml_content, parser=parser)


# ============================================================
# НОРМАЛИЗАЦИЯ
# ============================================================
def normalize_name(name: str) -> str:
    """Нормализует имя для поиска в БД, гарантируя единообразие пробелов вокруг скобок"""
    if not name:
        return ""
    # Вставляем пробел перед открывающей скобкой, если его нет
    name = re.sub(r'(?<!\s)\(', ' (', name)
    # Вставляем пробел после закрывающей скобки, если его нет
    name = re.sub(r'\)(?!\s)', ') ', name)
    # Удаляем всё, кроме букв, цифр, пробелов, дефиса, подчёркивания, скобок
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    # Схлопываем множественные пробелы
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    return normalized


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
        "Контактные поверхности": Category.CONTACT_SURFACES,
    }

    # ===== УДАЛЕНЫ ЖЁСТКИЕ СПИСКИ ROOM_CATEGORIES и ALLOWED_INVENTORY_COLORS =====
    # Теперь они загружаются из БД в __init__

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
        # ===== ДОБАВЛЕНО: кэши для категорий и цветов =====
        self._room_categories = set()      # допустимые категории помещений (нижний регистр)
        self._inventory_colors = set()     # допустимые цвета инвентаря (нижний регистр)
        self._load_room_categories()
        self._load_inventory_colors()

    def _load_room_categories(self):
        """Загружает список допустимых категорий помещений из БД."""
        session = SessionLocal()
        try:
            categories = session.query(RoomCategory).all()
            self._room_categories = {c.name.lower() for c in categories}
        finally:
            session.close()

    def _load_inventory_colors(self):
        """Загружает список допустимых цветов инвентаря из БД."""
        session = SessionLocal()
        try:
            colors = session.query(InventoryColor).all()
            self._inventory_colors = {c.name.lower() for c in colors}
        finally:
            session.close()

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

    # ===================== ИСПРАВЛЕННЫЕ МЕТОДЫ СБОРКИ ТЕКСТА =====================
    def _parse_sdt_cell(self, cell, data: ChecklistData):
        """Парсит ячейку с SDT чек-боксами (сбор полного текста, а не по runs)"""
        sdt_elements = cell.xpath('.//w:sdt', namespaces=self.NAMESPACES)
        if not sdt_elements:
            return

        # Собираем полный текст ячейки через XPath string()
        full_text = cell.xpath('string()', namespaces=self.NAMESPACES)
        full_text = full_text.strip()

        if not full_text or ('☒' not in full_text and '☐' not in full_text):
            return

        checkbox_count = full_text.count('☒') + full_text.count('☐')

        # Разбиваем строку на чекбоксы и текст
        parts = re.split(r'([☐☒])', full_text)
        current_obj = None
        text_parts = []
        current_state = False
        modifiers = []

        for part in parts:
            if part in ('☒', '☐'):
                if text_parts:
                    text = self._clean_text(''.join(text_parts))
                    if text:
                        if current_obj is None:
                            current_obj = {'name': text, 'checked': current_state, 'modifiers': []}
                        else:
                            modifiers.append({'name': text, 'checked': current_state})
                    text_parts = []
                current_state = (part == '☒')
            else:
                text_parts.append(part)

        # Последний фрагмент после последнего чекбокса
        if text_parts:
            text = self._clean_text(''.join(text_parts))
            if text:
                if current_obj is None:
                    current_obj = {'name': text, 'checked': current_state, 'modifiers': []}
                else:
                    modifiers.append({'name': text, 'checked': current_state})

        if current_obj:
            current_obj['modifiers'] = modifiers
            self._add_object_to_data(current_obj, checkbox_count, data)

    def _parse_unicode_cell(self, cell, data: ChecklistData):
        """Парсит ячейку с Unicode-символами ☒/☐ (полный текст)"""
        full_text = cell.xpath('string()', namespaces=self.NAMESPACES).strip()
        if not full_text or ('☒' not in full_text and '☐' not in full_text):
            return

        checkbox_count = full_text.count('☒') + full_text.count('☐')

        parts = re.split(r'([☐☒])', full_text)
        current_obj = None
        text_parts = []
        current_state = False
        modifiers = []

        for part in parts:
            if part in ('☒', '☐'):
                if text_parts:
                    text = self._clean_text(''.join(text_parts))
                    if text:
                        if current_obj is None:
                            current_obj = {'name': text, 'checked': current_state, 'modifiers': []}
                        else:
                            modifiers.append({'name': text, 'checked': current_state})
                    text_parts = []
                current_state = (part == '☒')
            else:
                text_parts.append(part)

        if text_parts:
            text = self._clean_text(''.join(text_parts))
            if text:
                if current_obj is None:
                    current_obj = {'name': text, 'checked': current_state, 'modifiers': []}
                else:
                    modifiers.append({'name': text, 'checked': current_state})

        if current_obj:
            current_obj['modifiers'] = modifiers
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
        """
        Извлекает категорию помещения из выпадающего списка (dropDownList SDT).
        Проверяет значение по БД (таблица room_categories).
        """
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
                    value_lower = value.lower()
                    # ===== ЗАМЕНА ЖЁСТКОГО СПИСКА НА ПРОВЕРКУ ПО БД =====
                    if value_lower in self._room_categories:
                        data.room_category = value_lower
                        print(f"  📋 Категория помещения (из выпадающего списка): «{value_lower}»")
                        return

        print("  ⚠️ Категория помещения не найдена (выпадающий список отсутствует или значение не в БД)")

    def _parse_additional_info(self, root, data: ChecklistData):
        """
        Извлекает моющие и дезинфицирующие средства, цвет инвентаря.
        Строгий порядок: 1 цвет + 18 средств (6 групп по 3).
        В каждой тройке: [0] – средство, [1] – концентрация, [2] – метод.
        Цвет инвентаря проверяется по БД (таблица inventory_colors).
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
                # Пропускаем значения, которые являются категориями помещений
                if value.lower() in self._room_categories:
                    continue
                all_values.append(value if value else None)

        print(f"  📦 Найдено выпадающих списков: {len(all_values)}")

        # --- Цвет инвентаря ---
        inventory_color = None
        if all_values:
            first_val = all_values[0]
            # ===== ЗАМЕНА ЖЁСТКОГО СПИСКА НА ПРОВЕРКУ ПО БД =====
            if first_val and first_val.lower() in self._inventory_colors:
                inventory_color = first_val.lower()
                print(f"  🎨 Найден цвет инвентаря: {inventory_color}")
            elif first_val in self.INVENTORY_COLOR_PLACEHOLDERS or first_val is None:
                print(f"  🎨 Цвет инвентаря не выбран (заглушка), будет 'промаркированный'")
            else:
                print(f"  ⚠️ Не удалось определить цвет инвентаря: '{first_val}', будет 'промаркированный'")
        data.inventory_color = inventory_color
        if inventory_color is None:
            print("  🎨 Цвет инвентаря не указан (будет 'промаркированный')")

        # --- Средства (18 элементов = 6 групп по 3) ---
        dropdown_values = all_values[1:] if len(all_values) > 1 else []
        while len(dropdown_values) < 18:
            dropdown_values.append(None)
        dropdown_values = dropdown_values[:18]

        # Группы с правильными именами полей
        groups = [
            ('cleaning', 0, False),      # общее моющее
            ('disinfection', 3, False),  # дезинфицирующее
            ('floor', 6, True),          # пол/трапы
            ('tech', 9, True),           # ДОБАВЛЕНО: технологическое оборудование
            ('thermal', 12, True),       # тепловое оборудование (было 9, стало 12)
            ('glass', 15, True)          # стекло/зеркала (было 12, стало 15)
        ]

        for group_name, start_idx, use_cleaning_suffix in groups:
            product = dropdown_values[start_idx]
            concentration = dropdown_values[start_idx + 1] if start_idx + 1 < len(dropdown_values) else None
            method = dropdown_values[start_idx + 2] if start_idx + 2 < len(dropdown_values) else None

            # Если средство пустое или заглушка – считаем, что ничего не выбрано
            if not product or product in self.INVENTORY_COLOR_PLACEHOLDERS:
                product = None
                concentration = None
                method = None

            # Присваиваем данные
            if use_cleaning_suffix:
                setattr(data, f'{group_name}_cleaning_product', product)
                setattr(data, f'{group_name}_cleaning_concentration', concentration)
                setattr(data, f'{group_name}_cleaning_method_text', method)
            else:
                setattr(data, f'{group_name}_product', product)
                setattr(data, f'{group_name}_concentration', concentration)
                setattr(data, f'{group_name}_method_text', method)

        # Вывод результатов
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
        if data.tech_cleaning_product:  # ДОБАВЛЕНО
            print(f"  ⚙️ Моющее для технологического оборудования: {data.tech_cleaning_product}")
            print(f"     Концентрация: {data.tech_cleaning_concentration or 'не указана'}")
            print(f"     Способ разведения: {data.tech_cleaning_method_text or 'не указан'}")
        else:
            print("  ⚙️ Моющее для технологического оборудования не выбрано")
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

    # ===================== ДОБАВЛЕНО: ПАРСИНГ ЗОНАЛЬНЫХ ИСПОЛНИТЕЛЕЙ =====================
    def _parse_executors(self, root, data: ChecklistData):
        """
        Извлекает зональных исполнителей из строк вида:
        «поверхности выше 2 м – Иванов И.И.»
        «поверхности до 2 м – Петров П.П.»
        «оборудование – Сидоров С.С.»

        Ищет по всему документу (в параграфах).
        """
        # Ищем все текстовые узлы (параграфы) в документе
        paragraphs = root.xpath('.//w:p', namespaces=self.NAMESPACES)
        for p in paragraphs:
            text = p.xpath('string()', namespaces=self.NAMESPACES).strip()
            if not text:
                continue

            # Приводим к нижнему регистру для поиска шаблонов, но сохраняем оригинал для извлечения
            text_lower = text.lower()

            # Шаблон: поверхности выше 2 м
            if 'поверхности выше 2 м' in text_lower:
                # Ищем разделитель – или :
                parts = re.split(r'[–\-:]', text, maxsplit=1)
                if len(parts) > 1:
                    val = parts[1].strip()
                    if val:
                        data.executor_high = val
                        print(f"  👤 Исполнитель для поверхностей выше 2 м: {val}")

            # Шаблон: поверхности до 2 м
            if 'поверхности до 2 м' in text_lower:
                parts = re.split(r'[–\-:]', text, maxsplit=1)
                if len(parts) > 1:
                    val = parts[1].strip()
                    if val:
                        data.executor_low = val
                        print(f"  👤 Исполнитель для поверхностей до 2 м: {val}")

            # Шаблон: оборудование (ищем только если есть разделитель)
            if 'оборудование' in text_lower and ('–' in text or '-' in text or ':' in text):
                parts = re.split(r'[–\-:]', text, maxsplit=1)
                if len(parts) > 1:
                    val = parts[1].strip()
                    if val:
                        data.executor_equipment = val
                        print(f"  👤 Исполнитель для оборудования: {val}")

        print(f"🔍 Парсинг исполнителей, найдено параграфов: {len(paragraphs)}")
        for p in paragraphs:
            text = p.xpath('string()', namespaces=self.NAMESPACES).strip()
            if text:
                print(f"   Текст: {text[:100]}")


    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                # ===== ИСПРАВЛЕНИЕ: безопасный парсинг XML (защита от XXE) =====
                root = _safe_fromstring(xml_content)

                self._parse_header(root, data)
                self._parse_room_category(root, data)
                self._parse_additional_info(root, data)
                # ===== ДОБАВЛЕНО: парсинг исполнителей =====
                self._parse_executors(root, data)

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
        if data.tech_cleaning_product:  # ДОБАВЛЕНО
            print(f"   Моющее для технологического оборудования: {data.tech_cleaning_product}")
        if data.thermal_cleaning_product:
            print(f"   Моющее для теплового оборудования: {data.thermal_cleaning_product}")
        if data.glass_cleaning_product:
            print(f"   Моющее для стекол/зеркал: {data.glass_cleaning_product}")
        if data.inventory_color:
            print(f"   Цвет инвентаря: {data.inventory_color}")
        else:
            print("   Цвет инвентаря: промаркированный")
        if data.executor_high:
            print(f"   Исполнитель (выше 2 м): {data.executor_high}")
        if data.executor_low:
            print(f"   Исполнитель (до 2 м): {data.executor_low}")
        if data.executor_equipment:
            print(f"   Исполнитель (оборудование): {data.executor_equipment}")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = SDTChecklistParser()
    return parser.parse(file_path)