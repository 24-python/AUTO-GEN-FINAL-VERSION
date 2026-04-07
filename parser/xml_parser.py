"""
Парсер с определением категорий по ключевым словам и обработкой составных позиций
"""

import zipfile
from pathlib import Path
from lxml import etree
import re

from parser.models import ChecklistData, ChecklistItem, Category


class CategoryParser:
    """Парсер с определением категорий по ключевым словам и обработкой составных позиций"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }

    # Составные позиции, требующие специальной обработки
    COMPOUND_PATTERNS = [
        {
            'base': r'потолок\s*\([^)]+\)',
            'modifiers': [
                {'marker': 'П', 'name': 'П', 'format': '{base} ({mod})'},
                {'marker': 'ОК', 'name': 'ОК', 'format': '{base} ({mod})'}
            ]
        },
        {
            'base': r'вытяжные\s+зонты',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} ({mod})'},
                {'marker': 'А', 'name': 'А', 'format': '{base} ({mod})'}
            ]
        },
        {
            'base': r'формы\s+для\s+выпечки',
            'modifiers': [
                {'marker': 'С', 'name': 'С', 'format': '{base} ({mod})'},
                {'marker': 'Н', 'name': 'Н', 'format': '{base} ({mod})'},
                {'marker': 'А', 'name': 'А', 'format': '{base} ({mod})'}
            ]
        },
        {
            'base': r'листы\s+для\s+выпечки',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} ({mod})'},
                {'marker': 'А', 'name': 'А', 'format': '{base} ({mod})'}
            ]
        },
        {
            'base': r'съёмные\s+детали\s+оборудования',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} ({mod})'},
                {'marker': 'А', 'name': 'А', 'format': '{base} ({mod})'}
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
                {'marker': r'холд\.?', 'name': 'холодильные', 'format': '{base} {mod}'},
                {'marker': r'мороз\.?', 'name': 'морозильные', 'format': '{base} {mod}'},
                {'marker': r'шок\.?\s+замор\.?', 'name': 'шоковой заморозки', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'плиты',
            'modifiers': [
                {'marker': r'индук\.?', 'name': 'индукционные', 'format': '{base} {mod}'},
                {'marker': r'элек\.?', 'name': 'электрические', 'format': '{base} {mod}'},
                {'marker': r'газ\.?', 'name': 'газовые', 'format': '{base} {mod}'}
            ]
        },
        {
            'base': r'производственные\s+столы',
            'modifiers': [
                {'marker': 'Н', 'name': 'Н', 'format': '{base} ({mod})'},
                {'marker': 'Д', 'name': 'Д', 'format': '{base} ({mod})'}
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
                {'marker': 'мука', 'name': 'для муки', 'format': '{base} {mod}'},
                {'marker': 'сахар', 'name': 'для сахара', 'format': '{base} {mod}'}
            ]
        }
    ]

    # Полный словарь категорий по ключевым словам (РАСШИРЕННЫЙ)
    KEYWORD_CATEGORIES = [
        # Бытовая техника
        ('чайники', Category.HOUSEHOLD_APPLIANCES),
        ('термопоты', Category.HOUSEHOLD_APPLIANCES),
        ('кулера', Category.HOUSEHOLD_APPLIANCES),
        ('микроволновки', Category.HOUSEHOLD_APPLIANCES),
        ('холодильники', Category.HOUSEHOLD_APPLIANCES),
        ('стиральные машины', Category.HOUSEHOLD_APPLIANCES),
        ('сушильные машины', Category.HOUSEHOLD_APPLIANCES),

        # Упаковочное оборудование
        ('вакуумные упаковщики', Category.PACKAGING_EQUIPMENT),
        ('запайщики', Category.PACKAGING_EQUIPMENT),
        ('этикетировщики', Category.PACKAGING_EQUIPMENT),
        ('термоусадочные туннели', Category.PACKAGING_EQUIPMENT),
        ('автоматический клипсаторы', Category.PACKAGING_EQUIPMENT),
        ('flow pack', Category.PACKAGING_EQUIPMENT),
        ('принтера для этикеток', Category.PACKAGING_EQUIPMENT),

        # Тепловое оборудование
        ('плиты', Category.THERMAL_EQUIPMENT),
        ('плиты индук', Category.THERMAL_EQUIPMENT),
        ('плиты элек', Category.THERMAL_EQUIPMENT),
        ('плиты газ', Category.THERMAL_EQUIPMENT),
        ('варочные котлы', Category.THERMAL_EQUIPMENT),
        ('сковороды', Category.THERMAL_EQUIPMENT),
        ('фритюры', Category.THERMAL_EQUIPMENT),
        ('грили', Category.THERMAL_EQUIPMENT),
        ('пароконвектоматы', Category.THERMAL_EQUIPMENT),
        ('печи подовые', Category.THERMAL_EQUIPMENT),
        ('печи ротационные', Category.THERMAL_EQUIPMENT),
        ('расстоечные шкафы', Category.THERMAL_EQUIPMENT),
        ('вафельницы', Category.THERMAL_EQUIPMENT),

        # Инвентарь, посуда
        ('доски', Category.INVENTORY),
        ('посуда', Category.INVENTORY),
        ('инвентарь', Category.INVENTORY),
        ('ножи', Category.INVENTORY),
        ('мусаты', Category.INVENTORY),
        ('секачи', Category.INVENTORY),
        ('шампура', Category.INVENTORY),
        ('крючки', Category.INVENTORY),
        ('вешала', Category.INVENTORY),
        ('гастроёмкости', Category.INVENTORY),
        ('формы для выпечки', Category.INVENTORY),
        ('листы для выпечки', Category.INVENTORY),
        ('съёмные детали оборудования', Category.INVENTORY),
        ('силапеновые коврики', Category.INVENTORY),
        ('внутрицеховая тара', Category.INVENTORY),
        ('дежи', Category.INVENTORY),
        ('передвижные ёмкости', Category.INVENTORY),
        ('ёмкости для перетаривания', Category.INVENTORY),
        ('корзины для расстойки теста', Category.INVENTORY),
        ('ёмкости для сыпучих продуктов', Category.INVENTORY),
        ('шпильки', Category.INVENTORY),
        ('тележки', Category.INVENTORY),
        ('листы от шпилек', Category.INVENTORY),
        ('тележки подкатные', Category.INVENTORY),
        ('оборотная тара', Category.INVENTORY),
        ('изотермические контейнеры', Category.INVENTORY),
        ('расстоечные термочехлы', Category.INVENTORY),
        ('отсадочные мешки', Category.INVENTORY),

        # Поверхности (РАСШИРЕНО)
        ('потолок', Category.SURFACE),
        ('светильники', Category.SURFACE),
        ('вытяжные зонты', Category.SURFACE),
        ('вентиляционные трубы', Category.SURFACE),
        ('диффузоры', Category.SURFACE),
        ('кондиционер', Category.SURFACE),
        ('сплит-системы', Category.SURFACE),
        ('инсектицидные лампы', Category.SURFACE),
        ('стены', Category.SURFACE),
        ('бактерицидные лампы', Category.SURFACE),
        ('стерилизатор', Category.SURFACE),
        ('металлоконструкции', Category.SURFACE),
        ('кабель-каналы', Category.SURFACE),
        ('водопроводные трубы', Category.SURFACE),
        ('выключатели', Category.SURFACE),
        ('розетки', Category.SURFACE),
        ('дверные ручки', Category.SURFACE),
        ('окна внешние', Category.SURFACE),
        ('окна внутрицеховые', Category.SURFACE),
        ('электрощиты', Category.SURFACE),
        ('пожарный щит', Category.SURFACE),
        ('отопительные приборы', Category.SURFACE),
        ('бойлера', Category.SURFACE),
        ('двери', Category.SURFACE),
        ('ворота/роллета', Category.SURFACE),
        ('завесы ПВХ', Category.SURFACE),
        ('завесы тепловые', Category.SURFACE),
        ('демосистема', Category.SURFACE),
        ('лестницы', Category.SURFACE),
        ('пол', Category.SURFACE),
        ('мусорные корзины', Category.SURFACE),
        ('контейнеры для отходов', Category.SURFACE),
        ('трапы', Category.SURFACE),
        ('стоки', Category.SURFACE),
        ('резиновые коврики', Category.SURFACE),
        ('гидравлические тележки', Category.SURFACE),
        ('погрузчики', Category.SURFACE),
        ('штабелёры', Category.SURFACE),
        ('пластиковые паллеты', Category.SURFACE),
        ('пластиковые подкаты', Category.SURFACE),
        ('держатели для ножей', Category.SURFACE),
        ('держатели для досок', Category.SURFACE),
        ('часы', Category.SURFACE),
        ('вентиляторы', Category.SURFACE),
        ('держатели для инвентаря', Category.SURFACE),

        # Моечный инвентарь
        ('мопы', Category.CLEANING_EQUIPMENT),
        ('ветошь', Category.CLEANING_EQUIPMENT),
        ('щётки', Category.CLEANING_EQUIPMENT),
        ('сгоны', Category.CLEANING_EQUIPMENT),
        ('ведра', Category.CLEANING_EQUIPMENT),
        ('АВД', Category.CLEANING_EQUIPMENT),
        ('пылесосы', Category.CLEANING_EQUIPMENT),
        ('поломоечная машина', Category.CLEANING_EQUIPMENT),

        # Технологическое оборудование (РАСШИРЕНО)
        ('производственные столы', Category.TECH_EQUIPMENT),
        ('весы', Category.TECH_EQUIPMENT),
        ('весы напольные', Category.TECH_EQUIPMENT),
        ('весы настольные', Category.TECH_EQUIPMENT),
        ('блендеры', Category.TECH_EQUIPMENT),
        ('миксеры планетарные', Category.TECH_EQUIPMENT),
        ('бисквиторезки', Category.TECH_EQUIPMENT),
        ('тестомесы', Category.TECH_EQUIPMENT),
        ('тестоделители', Category.TECH_EQUIPMENT),
        ('тестоокруглители', Category.TECH_EQUIPMENT),
        ('тестораскатки', Category.TECH_EQUIPMENT),
        ('прессы для теста', Category.TECH_EQUIPMENT),
        ('тарталетницы', Category.TECH_EQUIPMENT),
        ('пневматические распылители', Category.TECH_EQUIPMENT),
        ('термощупы', Category.TECH_EQUIPMENT),
        ('дробилки', Category.TECH_EQUIPMENT),
        ('машина для резки конд. изделий', Category.TECH_EQUIPMENT),
        ('ультразвуковые нарезки', Category.TECH_EQUIPMENT),
        ('водяные бани', Category.TECH_EQUIPMENT),
        ('минифилы', Category.TECH_EQUIPMENT),
        ('дозаторы для жидкостей', Category.TECH_EQUIPMENT),
        ('распылители для желе и сиропов', Category.TECH_EQUIPMENT),
        ('просеиватели', Category.TECH_EQUIPMENT),
        ('просеиватели мука', Category.TECH_EQUIPMENT),
        ('просеиватели сахар', Category.TECH_EQUIPMENT),
        ('солодоварки', Category.TECH_EQUIPMENT),
        ('ферментаторы', Category.TECH_EQUIPMENT),
        ('рентгеновские системы контроля', Category.TECH_EQUIPMENT),
        ('вакуумные роторные шприцы', Category.TECH_EQUIPMENT),
        ('овощерезки', Category.TECH_EQUIPMENT),
        ('овощечистки', Category.TECH_EQUIPMENT),
        ('измельчители', Category.TECH_EQUIPMENT),
        ('слайсера', Category.TECH_EQUIPMENT),
        ('протирочные машины', Category.TECH_EQUIPMENT),
        ('картофелечистки', Category.TECH_EQUIPMENT),
        ('депозитор волюметрический', Category.TECH_EQUIPMENT),
        ('металлодетектор', Category.TECH_EQUIPMENT),

        # Сантехника
        ('раковина', Category.PLUMBING),
        ('смеситель', Category.PLUMBING),
        ('унитаз', Category.PLUMBING),
        ('писсуар', Category.PLUMBING),
        ('душевая кабина', Category.PLUMBING),
        ('душевые секции', Category.PLUMBING),
        ('технологическая мойка', Category.PLUMBING),

        # Посудомоечное оборудование
        ('ПММ', Category.DISHWASHING_EQUIPMENT),
        ('таромоечная машина', Category.DISHWASHING_EQUIPMENT),

        # Холодильное оборудование
        ('камеры', Category.REFRIGERATION_EQUIPMENT),
        ('камеры холд', Category.REFRIGERATION_EQUIPMENT),
        ('камеры мороз', Category.REFRIGERATION_EQUIPMENT),
        ('камеры шок', Category.REFRIGERATION_EQUIPMENT),
        ('морозильный ларь', Category.REFRIGERATION_EQUIPMENT),
        ('холод. столы', Category.REFRIGERATION_EQUIPMENT),
        ('холод. шкафы', Category.REFRIGERATION_EQUIPMENT),
        ('ледогенератор', Category.REFRIGERATION_EQUIPMENT),

        # Мебель
        ('столы', Category.FURNITURE),
        ('стулья', Category.FURNITURE),
        ('полки', Category.FURNITURE),
        ('стеллажи', Category.FURNITURE),
        ('шкафы', Category.FURNITURE),
        ('вешалки', Category.FURNITURE),
        ('скамейки', Category.FURNITURE),
        ('зеркала', Category.FURNITURE),
        ('сушилка для обуви', Category.FURNITURE),
        ('обувницы', Category.FURNITURE),

        # Офисная техника
        ('компьютеры', Category.OFFICE_EQUIPMENT),
        ('мониторы', Category.OFFICE_EQUIPMENT),
        ('принтера', Category.OFFICE_EQUIPMENT),
        ('телефоны', Category.OFFICE_EQUIPMENT),

        # Санитарный пост
        ('санитарный пост', Category.SANITARY_POST),
        ('корзины/диспенсеры для СИЗ', Category.SANITARY_POST),
        ('диспенсер для полотенец', Category.SANITARY_POST),
        ('диспенсер для т/б', Category.SANITARY_POST),
        ('дозаторы для мыла/антисептика', Category.SANITARY_POST),
        ('сушилка для рук', Category.SANITARY_POST),
        ('модульный санпропускник', Category.SANITARY_POST),
        ('дезинфицирующие маты', Category.SANITARY_POST),
        ('бак для грязной одежды', Category.SANITARY_POST),

        # СИЗ
        ('фартуки', Category.PPE),
        ('сапоги', Category.PPE),
        ('перчатки', Category.PPE),
        ('нарукавники', Category.PPE),
    ]

    def __init__(self):
        self.current_category = Category.OTHER

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))
        self.current_category = Category.OTHER

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                # Парсим информацию об участке и предприятии из первой строки
                self._parse_header_from_first_row(root, data)

                # Находим все ячейки
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

                if len(cells) >= 2:
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

    def _get_category_by_keywords(self, item_text: str) -> Category:
        """Определяет категорию по ключевым словам"""
        item_lower = item_text.lower()

        # Специальное правило для посудомоечного оборудования
        if 'пмм' in item_lower:
            return Category.DISHWASHING_EQUIPMENT
        if 'таромоечная' in item_lower:
            return Category.DISHWASHING_EQUIPMENT

        # Специальное правило для машины для резки
        if 'машина для резки' in item_lower:
            return Category.TECH_EQUIPMENT

        # Специальное правило для форм для выпечки
        if 'форма для выпечки' in item_lower:
            return Category.INVENTORY

        # Специальное правило для листов для выпечки
        if 'лист для выпечки' in item_lower:
            return Category.INVENTORY

        # Для каждого ключевого слова проверяем наличие в тексте
        for keyword, category in self.KEYWORD_CATEGORIES:
            if keyword in item_lower:
                return category

        return Category.OTHER

    def _is_checkbox_before(self, text: str, position: int) -> bool:
        """
        Проверяет, есть ли символ чек-бокса перед указанной позицией.
        Пропускает пробелы, табуляцию и знаки пунктуации (:, ;, и т.д.).
        """
        pos = position - 1

        # Пропускаем пробелы, табуляцию, двоеточия и другие разделители
        while pos >= 0 and text[pos] in [' ', '\t', ':', ';', ',', '.', '-', '—', '–']:
            pos -= 1

        # Проверяем, есть ли чек-бокс
        return pos >= 0 and text[pos] == '☒'

    def _process_compound_cell(self, full_text: str, data: ChecklistData) -> bool:
        """
        Специальная обработка составных ячеек.
        Возвращает True, если ячейка обработана как составная.
        """
        # Перебираем все паттерны составных позиций
        for pattern in self.COMPOUND_PATTERNS:
            # Ищем базовый объект
            base_match = re.search(pattern['base'], full_text, re.IGNORECASE)
            if not base_match:
                continue

            base_name = base_match.group(0)
            base_start = base_match.start()

            # Проверяем, отмечен ли базовый объект
            base_checked = self._is_checkbox_before(full_text, base_start)

            # Собираем все отмеченные модификаторы
            checked_modifiers = []

            # Для каждого модификатора проверяем, отмечен ли он
            for modifier in pattern['modifiers']:
                # Ищем все вхождения модификатора в тексте
                for mod_match in re.finditer(modifier['marker'], full_text, re.IGNORECASE):
                    mod_pos = mod_match.start()

                    # Убеждаемся, что этот модификатор находится ПОСЛЕ базового объекта
                    if mod_pos < base_match.end():
                        continue

                    # Проверяем окружение модификатора
                    if mod_pos > 0 and full_text[mod_pos - 1].isalpha() and mod_pos - 1 >= base_match.end():
                        continue

                    # Проверяем, есть ли чек-бокс перед модификатором
                    if self._is_checkbox_before(full_text, mod_pos):
                        checked_modifiers.append(modifier)
                        break

            # Формируем результат
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

                    if 'category' in pattern:
                        category = pattern['category']
                    else:
                        category = self._get_category_by_keywords(item_name)

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
                    category = self._get_category_by_keywords(base_name)

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

                        item_text = re.sub(r'[^\w\s\-\(\)]', '', item_text)
                        item_text = re.sub(r'\s+', ' ', item_text).strip()
                        item_text = re.sub(r'\s+[А-Я]$', '', item_text)

                        if item_text and len(item_text) > 1:
                            category = self._get_category_by_keywords(item_text)

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