from docx import Document
from docx.shared import Pt
from pathlib import Path
from parser.models import ChecklistData
from db.models import Instruction, Category as DBCategory
from db.database import SessionLocal
from generator.mapper import find_object_and_instructions
from collections import defaultdict


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE

    def _set_cell_font(self, cell, text: str, font_name: str = 'Arial', size_pt: int = 9, bold: bool = False):
        """Устанавливает текст и шрифт в ячейке"""
        cell.text = text
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(size_pt)
                run.font.bold = bold

    def _apply_font_to_row(self, row, font_name: str = 'Arial', size_pt: int = 9, bold: bool = False):
        """Применяет шрифт ко всем ячейкам строки"""
        for col in range(len(row.cells)):
            for paragraph in row.cells[col].paragraphs:
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(size_pt)
                    run.font.bold = bold

    def _merge_cells_horizontal(self, row, start_col: int, end_col: int):
        """Объединяет ячейки в строке по горизонтали"""
        if start_col >= end_col:
            return
        start_cell = row.cells[start_col]
        for col in range(start_col + 1, end_col + 1):
            start_cell.merge(row.cells[col])

    def _get_category_priority(self, session) -> dict:
        """Загружает приоритеты категорий из БД"""
        categories = session.query(DBCategory).order_by(DBCategory.sort_order).all()
        return {cat.name: cat.sort_order for cat in categories}

    def generate(self, checklist_data: ChecklistData, output_path: str) -> str:
        doc = Document(self.template_path)

        # Основная таблица - вторая в документе (индекс 1)
        main_table = doc.tables[1]

        # Заполняем помещение (строка 1, ячейка 0)
        room_cell = main_table.cell(1, 0)
        room_text = room_cell.text
        if "Помещение:" in room_text:
            new_text = f"Помещение: {checklist_data.room_name or '______________'}"
            self._set_cell_font(room_cell, new_text, bold=True)

        # === ПОЛУЧАЕМ ПРИОРИТЕТЫ КАТЕГОРИЙ ИЗ БД ===
        session = SessionLocal()
        category_priority = self._get_category_priority(session)

        # Группируем объекты по категориям
        category_items = defaultdict(list)
        category_order = {}

        for item in checklist_data.get_checked_items():
            obj, instructions = find_object_and_instructions(session, item.name, item.category)

            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            if instructions:
                for instr in instructions:
                    category_items[cat_name].append((item.name, instr, item.category))
            else:
                category_items[cat_name].append((item.name, None, item.category))

        session.close()

        # === СОРТИРУЕМ ===
        sorted_categories = sorted(category_items.keys(), key=lambda x: category_order.get(x, 999))
        for cat_name in sorted_categories:
            category_items[cat_name].sort(key=lambda x: x[0])

        # Заполняем таблицу
        start_row = 6  # после 5 объединенных строк + заголовки

        # Удаляем старые строки данных
        while len(main_table.rows) > start_row:
            tbl = main_table._tbl
            tbl.remove(main_table.rows[start_row]._tr)

        current_row = start_row

        # === ВЫВОД ===
        for category_name in sorted_categories:
            items = category_items[category_name]

            # Строка категории
            if current_row >= len(main_table.rows):
                main_table.add_row()
            category_row = main_table.rows[current_row]
            self._merge_cells_horizontal(category_row, 0, 11)
            self._set_cell_font(category_row.cells[0], category_name, bold=True)
            current_row += 1

            # Строки объектов
            for obj_name, instr, _ in items:
                if current_row >= len(main_table.rows):
                    main_table.add_row()
                row = main_table.rows[current_row]

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
                else:
                    for col in range(1, 12):
                        row.cells[col].text = ""

                self._apply_font_to_row(row, font_name='Arial', size_pt=8, bold=False)
                current_row += 1

        # Сохраняем
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()

        doc.save(str(output_file))
        return str(output_file)