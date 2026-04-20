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

    def _get_category_priority(self, session) -> dict:
        categories = session.query(DBCategory).order_by(DBCategory.sort_order).all()
        return {cat.name: cat.sort_order for cat in categories}

    def generate(self, checklist_data: ChecklistData, output_path: str) -> str:
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # Заполняем помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '______________'}"
            if room_cell.paragraphs and room_cell.paragraphs[0].runs:
                room_cell.paragraphs[0].runs[0].font.bold = True

        # === ЗАГРУЗКА ДАННЫХ ===
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

        category_items = defaultdict(list)
        category_order = {}

        for item in checked_items:
            obj = objects_dict.get(item.name)
            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            display_name = obj.display_name if obj else item.name

            if obj and obj.id in instructions_dict:
                for instr in instructions_dict[obj.id]:
                    category_items[cat_name].append((display_name, instr))
            else:
                category_items[cat_name].append((display_name, None))

        session.close()

        # === СОРТИРОВКА ===
        sorted_categories = sorted(category_items.keys(), key=lambda x: category_order.get(x, 999))
        for cat_name in sorted_categories:
            category_items[cat_name].sort(key=lambda x: x[0])

        # === ПОДГОТОВКА ДАННЫХ ===
        rows_data = []
        for category_name in sorted_categories:
            items = category_items[category_name]
            if not items:
                continue
            rows_data.append(('category', category_name, None))
            for obj_name, instr in items:
                rows_data.append(('object', obj_name, instr))

        # === БЫСТРАЯ ОЧИСТКА ТАБЛИЦЫ ===
        start_row = 6
        tbl = main_table._tbl
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tr_elements = tbl.findall('.//w:tr', namespaces=ns)
        while len(tr_elements) > start_row:
            tbl.remove(tr_elements[-1])
            tr_elements = tbl.findall('.//w:tr', namespaces=ns)

        # === МАССОВОЕ ДОБАВЛЕНИЕ СТРОК ===
        for _ in range(len(rows_data)):
            main_table.add_row()

        # === БЫСТРОЕ ЗАПОЛНЕНИЕ (без _apply_font_to_row) ===
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

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()

        doc.save(str(output_file))
        return str(output_file)