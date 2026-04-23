# generator/docx_generator.py

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from parser.models import ChecklistData
from db.models import Instruction, Category as DBCategory, Object as DBObject
from db.database import SessionLocal
from collections import defaultdict
import re


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    PRODUCT_COLORS = {
        "ХИМИТЕК ПОЛИДЕЗ®-СУПЕР": "FFFFCC",
        "ХИМИТЕК УНИВЕРСАЛ-ПД-Н": "99FF99",
        "ХИМИТЕК ЧУДОДЕЙ®-КОМБИ-ПЕНАКТИВ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-CIP": "99CCFF",
        "ХИМИТЕК ПОЛИКОР®": "FFCCCC",
        "ХИМИТЕК ЧАРОЙТ®-СПРЕЙ": "99FF99",
        "ХИМИТЕК СВЕЖЕСТЬ-АНТИСЕПТИК": "FFFFCC",
    }

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE

    def _normalize_product_name(self, name: str) -> str:
        if not name:
            return ""
        normalized = re.sub(r'[^\w\s\-]', '', name)
        normalized = normalized.lower()
        normalized = re.sub(r'\s+', ' ', normalized)
        return normalized.strip()

    def _get_color_for_product(self, product_name: str) -> str:
        if not product_name:
            return None
        normalized_input = self._normalize_product_name(product_name)
        for key, color in self.PRODUCT_COLORS.items():
            if self._normalize_product_name(key) == normalized_input:
                return color
        return None

    def _set_cell_background(self, cell, hex_color: str):
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), hex_color)
        cell._tc.get_or_add_tcPr().append(shading)

    def _merge_cells_horizontal(self, row, start_col: int, end_col: int):
        if start_col >= end_col:
            return
        start_cell = row.cells[start_col]
        for col in range(start_col + 1, end_col + 1):
            start_cell.merge(row.cells[col])

    def _merge_cells_vertical(self, table, col: int, start_row: int, end_row: int):
        """Объединяет ячейки по вертикали"""
        if start_row >= end_row:
            return
        start_cell = table.cell(start_row, col)
        for row in range(start_row + 1, end_row + 1):
            start_cell.merge(table.cell(row, col))

    def _cells_are_equal(self, table, col: int, start_row: int, end_row: int) -> bool:
        """Проверяет, одинаковые ли значения в ячейках колонки"""
        first_text = table.cell(start_row, col).text.strip()
        for row in range(start_row + 1, end_row + 1):
            if table.cell(row, col).text.strip() != first_text:
                return False
        return True

    def _get_category_priority(self, session) -> dict:
        categories = session.query(DBCategory).order_by(DBCategory.sort_order).all()
        return {cat.name: cat.sort_order for cat in categories}

    def _get_instruction_signature(self, instructions: list) -> tuple:
        """
        Создаёт сигнатуру для полного набора инструкций объекта.
        Сигнатура — это кортеж из отсортированных сигнатур каждой инструкции.
        """
        if not instructions:
            return (("empty",),)

        signatures = []
        for instr in sorted(instructions, key=lambda x: x.cleaning_method or ""):
            sig = (
                instr.cleaning_method or "",
                instr.instruction_number or "",
                instr.product_name or "",
                instr.cleaning_technique or "",
                instr.concentration or "",
                instr.temperature or "",
                instr.exposure_time or "",
                instr.inventory or "",
                instr.frequency or "",
                instr.executor or "",
                instr.control_method or "",
            )
            signatures.append(sig)

        return tuple(signatures)

    def _group_by_full_instructions(self, object_instructions: list) -> list:
        """
        Группирует объекты с ПОЛНОСТЬЮ ИДЕНТИЧНЫМИ наборами инструкций.
        object_instructions: список [(obj_name, [instructions]), ...]
        Возвращает: список [(merged_names, [instructions]), ...]
        """
        groups = {}

        for obj_name, instructions in object_instructions:
            signature = self._get_instruction_signature(instructions)

            if signature not in groups:
                groups[signature] = {
                    "names": [],
                    "instructions": instructions
                }
            groups[signature]["names"].append(obj_name)

        result = []
        for signature, data in groups.items():
            merged_name = ", ".join(sorted(data["names"]))
            result.append((merged_name, data["instructions"]))

        return result

    def generate(self, checklist_data: ChecklistData, output_path: str, mode: int = 1) -> str:
        """
        Генерирует техкарту.
        mode: 1 - по категориям (с заголовками), 2 - по приоритету (единый список)
        """
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # Заполняем помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '______________'}"
            if room_cell.paragraphs and room_cell.paragraphs[0].runs:
                room_cell.paragraphs[0].runs[0].font.bold = True

        # === ЗАГРУЗКА ДАННЫХ ИЗ БД ===
        session = SessionLocal()
        category_priority = self._get_category_priority(session)

        checked_items = checklist_data.get_checked_items()
        all_names = [item.name for item in checked_items]

        db_objects = session.query(DBObject).filter(DBObject.normalized_name.in_(all_names)).all()
        objects_dict = {obj.normalized_name: obj for obj in db_objects}

        object_ids = [obj.id for obj in db_objects]
        db_instructions = session.query(Instruction).filter(Instruction.object_id.in_(object_ids)).all()

        instructions_dict = defaultdict(list)
        for instr in db_instructions:
            instructions_dict[instr.object_id].append(instr)

        # Собираем данные: для каждого объекта — его полный набор инструкций
        category_object_instructions = defaultdict(list)  # для режима 1
        all_object_instructions = []  # для режима 2
        category_order = {}

        for item in checked_items:
            obj = objects_dict.get(item.name)
            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            display_name = obj.display_name if obj else item.name
            sort_priority = obj.sort_priority if obj else 999

            if obj and obj.id in instructions_dict:
                instructions = instructions_dict[obj.id]
            else:
                instructions = []

            category_object_instructions[cat_name].append((display_name, instructions))
            all_object_instructions.append((display_name, instructions, sort_priority))

        session.close()

        # === ПОДГОТОВКА ДАННЫХ В ЗАВИСИМОСТИ ОТ РЕЖИМА ===
        rows_data = []
        merge_info = []  # (start_row, end_row) для вертикального объединения

        if mode == 1:
            # Режим 1: по категориям
            sorted_categories = sorted(category_object_instructions.keys(), key=lambda x: category_order.get(x, 999))

            for cat_name in sorted_categories:
                items = category_object_instructions[cat_name]
                if not items:
                    continue

                # Сортируем объекты внутри категории по алфавиту
                items.sort(key=lambda x: x[0])

                # Группируем объекты с полностью идентичными наборами инструкций
                grouped = self._group_by_full_instructions(items)

                # Добавляем заголовок категории
                rows_data.append(('category', cat_name, None))

                # Добавляем строки для каждой группы инструкций
                for obj_name, instructions in grouped:
                    group_start_row = len(rows_data)

                    if instructions:
                        for i, instr in enumerate(instructions):
                            # Только первая строка содержит имя объекта
                            display_name = obj_name if i == 0 else ""
                            rows_data.append(('object', display_name, instr))
                    else:
                        rows_data.append(('object', obj_name, None))

                    group_end_row = len(rows_data) - 1

                    # Если в группе больше одной строки — запоминаем для объединения
                    if group_end_row > group_start_row:
                        merge_info.append((group_start_row, group_end_row))
        else:
            # Режим 2: по приоритету (единый список)
            # Сортируем по sort_priority, затем по display_name
            all_object_instructions.sort(key=lambda x: (x[2], x[0]))

            # Группируем объекты с полностью идентичными наборами инструкций
            items_for_grouping = [(name, instrs) for name, instrs, _ in all_object_instructions]
            grouped = self._group_by_full_instructions(items_for_grouping)

            # Добавляем строки для каждой группы инструкций
            for obj_name, instructions in grouped:
                group_start_row = len(rows_data)

                if instructions:
                    for i, instr in enumerate(instructions):
                        # Только первая строка содержит имя объекта
                        display_name = obj_name if i == 0 else ""
                        rows_data.append(('object', display_name, instr))
                else:
                    rows_data.append(('object', obj_name, None))

                group_end_row = len(rows_data) - 1

                # Если в группе больше одной строки — запоминаем для объединения
                if group_end_row > group_start_row:
                    merge_info.append((group_start_row, group_end_row))

        # === ОЧИСТКА ТАБЛИЦЫ ===
        start_row = 6
        tbl = main_table._tbl
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tr_elements = tbl.findall('.//w:tr', namespaces=ns)
        while len(tr_elements) > start_row:
            tbl.remove(tr_elements[-1])
            tr_elements = tbl.findall('.//w:tr', namespaces=ns)

        # === ДОБАВЛЕНИЕ СТРОК ===
        for _ in range(len(rows_data)):
            main_table.add_row()

        # === ЗАПОЛНЕНИЕ ===
        current_row = start_row
        for row_info in rows_data:
            row = main_table.rows[current_row]

            if row_info[0] == 'category':
                self._merge_cells_horizontal(row, 0, 11)
                row.cells[0].text = row_info[1]
                if row.cells[0].paragraphs and row.cells[0].paragraphs[0].runs:
                    row.cells[0].paragraphs[0].runs[0].font.bold = True

            elif row_info[0] == 'object':
                obj_name, instr = row_info[1], row_info[2]
                row.cells[0].text = obj_name

                if instr:
                    row.cells[1].text = instr.cleaning_method or ""
                    row.cells[2].text = instr.instruction_number or ""
                    row.cells[3].text = instr.product_name or ""
                    row.cells[4].text = instr.cleaning_technique or ""
                    row.cells[5].text = instr.concentration or ""
                    row.cells[6].text = instr.temperature or ""
                    row.cells[7].text = instr.exposure_time or ""
                    row.cells[8].text = instr.inventory or ""
                    row.cells[9].text = instr.frequency or ""
                    row.cells[10].text = instr.executor or ""
                    row.cells[11].text = instr.control_method or ""

                    if instr.product_name:
                        color = self._get_color_for_product(instr.product_name)
                        if color:
                            self._set_cell_background(row.cells[3], color)
                else:
                    for col in range(1, 12):
                        row.cells[col].text = ""

            current_row += 1

        # === ВЕРТИКАЛЬНОЕ ОБЪЕДИНЕНИЕ ===
        # Корректируем индексы с учётом start_row
        for group_start, group_end in merge_info:
            actual_start = start_row + group_start
            actual_end = start_row + group_end

            if actual_end > actual_start:
                # Очищаем ячейки в первой колонке (кроме первой) ДО объединения
                for row in range(actual_start + 1, actual_end + 1):
                    main_table.cell(row, 0).text = ""
                # Объединяем первую колонку
                self._merge_cells_vertical(main_table, 0, actual_start, actual_end)

                # Объединяем колонки 9, 10, 11, 12 (индексы 8, 9, 10, 11) если значения одинаковые
                for col in [8, 9, 10, 11]:
                    if self._cells_are_equal(main_table, col, actual_start, actual_end):
                        # Очищаем ячейки (кроме первой) ДО объединения
                        for row in range(actual_start + 1, actual_end + 1):
                            main_table.cell(row, col).text = ""
                        # Объединяем
                        self._merge_cells_vertical(main_table, col, actual_start, actual_end)

        # Сохраняем
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()

        doc.save(str(output_file))
        return str(output_file)