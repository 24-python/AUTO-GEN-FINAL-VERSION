# generator/docx_generator.py

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pathlib import Path
from parser.models import ChecklistData
from db.models import Instruction, Category as DBCategory, Object as DBObject, RoomCategory
from db.database import SessionLocal
from collections import defaultdict
import re
import copy


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    PRODUCT_COLORS = {
        # Дезинфицирующие средства (жёлтый)
        "ХИМИТЕК ПОЛИДЕЗ®-СУПЕР": "FFFFCC",
        "ХИМИТЕК ПОЛИДЕЗ®-ЭКСПРЕСС": "FFFFCC",
        "ХИМИТЕК ПОЛИДЕЗ®": "FFFFCC",
        "ПОЛИДЕЗ®": "FFFFCC",
        "ХИМИТЕК УНИВЕРСАЛ-ДЕЗ": "FFFFCC",
        "ХИМИТЕК СВЕЖЕСТЬ-АНТИСЕПТИК": "FFFFCC",

        # Нейтральные моющие (зелёный)
        "ХИМИТЕК УНИВЕРСАЛ-ПД-Н": "99FF99",
        "ХИМИТЕК УНИВЕРСАЛ-ПД": "99FF99",
        "ХИМИТЕК УНИВЕРСАЛ-ПД для дозирующих систем": "99FF99",
        "ХИМИТЕК ИЗУМРУД 100": "99FF99",
        "ХИМИТЕК ИЗУМРУД 110": "99FF99",
        "ХИМИТЕК ИЗУМРУД 300": "99FF99",
        "ХИМИТЕК ИЗУМРУД 310": "99FF99",
        "ХИМИТЕК ИНТЕРЬЕР-ОФИС": "99FF99",
        "ХИМИТЕК ИНТЕРЬЕР-ОФИС-СПРЕЙ": "99FF99",
        "ХИМИТЕК КУХМАСТЕР-ГЕЛЬ": "99FF99",
        "ХИМИТЕК КУХМАСТЕР": "99FF99",
        "ХИМИТЕК КУХМАСТЕР для дозирующих систем": "99FF99",
        "ХИМИТЕК КУХМАСТЕР-ФОРТЕ": "99FF99",
        "ХИМИТЕК КУХМАСТЕР-ФОРТЕ для дозирующих систем": "99FF99",
        "ХИМИТЕК ПЕКАРЬ-АКТИВАТОР": "99FF99",
        "ХИМИТЕК СТИРАЛЬ-02": "99FF99",
        "ХИМИТЕК ЧАРОИТ®": "99FF99",
        "ХИМИТЕК ЧАРОЙТ®-СПРЕЙ": "99FF99",

        # Щелочные и специальные моющие (голубой)
        "ХИМИТЕК ЧУДОДЕЙ®-КОМБИ-ПЕНАКТИВ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-CIP": "99CCFF",
        "ХИМИТЕК ИЗУМРУД 400": "99CCFF",
        "ХИМИТЕК ИЗУМРУД 420": "99CCFF",
        "ХИМИТЕК КЕРАМИК-БЕЛИЗНА": "99CCFF",
        "ХИМИТЕК КЕРАМИК-БЛЕСК": "99CCFF",
        "ХИМИТЕК КЕРАМИК-РЕЛЬЕФ": "99CCFF",
        "ХИМИТЕК КУХМАСТЕР-ПРОФИ": "99CCFF",
        "ХИМИТЕК КУХМАСТЕР-ПРОФИ 12°Ж": "99CCFF",
        "ХИМИТЕК ПЕКАРЬ-АНТИНАГАР": "99CCFF",
        "ХИМИТЕК ПЕНАПОЛ-ПРОФИ": "99CCFF",
        "ХИМИТЕК СПЕЦ-УНИВЕРСАЛ-ЦВМ": "99CCFF",
        "ХИМИТЕК СПЕЦ-УНИВЕРСАЛ": "99CCFF",
        "ХИМИТЕК СПЕЦ-УНИВЕРСАЛ 220002": "99CCFF",
        "ХИМИТЕК СТИРАЛЬ-ПРОФИ": "99CCFF",
        "ХИМИТЕК СТИРАЛЬ-03": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ® модификация 190025": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-АНТИНАГАР": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-АНТИНАГАР-ГЕЛЬ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-АНТИНАГАР-ПЕНАКТИВ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-АНТИНАГАР-ПЕНАКТИВ 210014": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-КОМБИ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-КОМБИ-ЦВМ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-ПОЛИПРОМ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-ПОЛИПРОМ 190022": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-ФОРТЕ": "99CCFF",
        "ХИМИТЕК ЧУДОДЕЙ®-ЭКСПРЕСС": "99CCFF",
        "ХИМИТЕК ШУНГИТ 100": "99CCFF",
        "ХИМИТЕК ШУНГИТ 200": "99CCFF",

        # Кислотные моющие (розовый)
        "ХИМИТЕК ПОЛИКОР®": "FFCCCC",
        "ХИМИТЕК АНТИМИНЕРАЛ-ЛАКТО-ПЕНАКТИВ": "FFCCCC",
        "ХИМИТЕК АНТИМИНЕРАЛ-ЛАКТО": "FFCCCC",
        "ХИМИТЕК АНТИМИНИРАЛ-ФОРТЕ": "FFCCCC",
        "ХИМИТЕК АНТИЗАПАХ-ФОРТЕ": "FFCCCC",
        "ХИМИТЕК АНТИМИНЕРАЛ-CIP": "FFCCCC",
        "ХИМИТЕК ИНТЕРЬЕР": "FFCCCC",
        "ХИМИТЕК ИЗУМРУД 500": "FFCCCC",
        "ХИМИТЕК КУХМАСТЕР-ОПОЛАСКИВАТЕЛЬ": "FFCCCC",
        "ХИМИТЕК ПОЛИКОР®-ГЕЛЬ": "FFCCCC",
        "ХИМИТЕК ПОЛИКОР®-ГЕЛЬ ДДС": "FFCCCC",
        "ХИМИТЕК ПОЛИКОР-ГЕЛЬ® ДДС 180022": "FFCCCC",
        "ХИМИТЕК ПОЛИКОР® ДДС": "FFCCCC",
    }

    # Порядок способов обработки: прочистка -> очистка -> мойка -> ополаскивание -> дезинфекция
    CLEANING_METHOD_ORDER = {
        "прочистка": 1,
        "очистка (обеспыливание поверхностей)": 2,
        "мойка": 3,
        "ополаскивание": 4,
        "дезинфекция": 5,
    }

    # Порядок уровней обработки для вывода
    LEVEL_ORDER = {"основная": 1, "поддерживающая": 2, "генеральная": 3}

    # Объекты для специальных моющих средств
    FLOOR_OBJECTS = {"пол", "трапы"}
    GLASS_OBJECTS = {"зеркала", "окна внешние", "окна внутрицеховые", "монитор"}
    THERMAL_OBJECTS = {
        "плиты индук", "плиты элек", "плиты газ",
        "варочные котлы", "сковороды", "фритюры", "грили",
        "пароконвектоматы", "печи подовые", "печи ротационные",
        "расстоечные шкафы", "вафельницы"
    }

    # Объекты, для которых обязательно добавляется поддерживающая обработка
    SUPPORT_MAINTENANCE_OBJECTS = {
        "прибор кисл теста",
        "бисквиторезки",
        "блендеры",
        "вакуумные роторные шприцы",
        "водяные бани",
        "депозитор волюметрический",
        "дозатор для геля (пульверизатор)",
        "дозаторы для жидкостей",
        "дробилки",
        "измельчители",
        "картофелечистки",
        "машина для резки конд изделий",
        "металлодетектор",
        "миксеры планетарные",
        "минифилы (дозаторы крема)",
        "овощерезки",
        "овощечистки",
        "отсадочные машины",
        "пневматические распылители",
        "прессы для теста",
        "просеиватели мука",
        "просеиватели сахар",
        "протирочные машины",
        "распылители для желе и сиропов",
        "рентгеновские системы контроля",
        "слайсера",
        "солодоварки",
        "тарталетницы",
        "термощупы",
        "тестоделители",
        "тестомесы",
        "тестоокруглители",
        "тестораскатки",
        "ультразвуковые нарезки",
        "ферментаторы",
        "шприц-дозатор начинки",
        "весы напольные",
        "весы настольные",
        "производственные столы д",
        "производственные столы н",
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
        if start_row >= end_row:
            return
        start_cell = table.cell(start_row, col)
        for row in range(start_row + 1, end_row + 1):
            start_cell.merge(table.cell(row, col))

    def _cells_are_equal(self, table, col: int, start_row: int, end_row: int) -> bool:
        first_text = table.cell(start_row, col).text.strip()
        for row in range(start_row + 1, end_row + 1):
            if table.cell(row, col).text.strip() != first_text:
                return False
        return True

    def _merge_adjacent_equal_cells(self, table, col: int, start_row: int, end_row: int):
        if start_row >= end_row:
            return
        range_start = start_row
        range_value = table.cell(range_start, col).text.strip()
        for row in range(start_row + 1, end_row + 1):
            current_value = table.cell(row, col).text.strip()
            if current_value == range_value:
                continue
            else:
                if row - 1 > range_start:
                    for r in range(range_start + 1, row):
                        table.cell(r, col).text = ""
                    self._merge_cells_vertical(table, col, range_start, row - 1)
                range_start = row
                range_value = current_value
        if end_row > range_start:
            for r in range(range_start + 1, end_row + 1):
                table.cell(r, col).text = ""
            self._merge_cells_vertical(table, col, range_start, end_row)

    def _get_category_priority(self, session) -> dict:
        categories = session.query(DBCategory).order_by(DBCategory.sort_order).all()
        return {cat.name: cat.sort_order for cat in categories}

    def _get_room_category_id(self, session, room_category_name: str) -> int:
        """Ищет категорию помещения по точному совпадению (в БД все в нижнем регистре)."""
        if not room_category_name:
            return None
        rc = session.query(RoomCategory).filter(RoomCategory.name == room_category_name).first()
        return rc.id if rc else None

    def _get_instruction_signature(self, instructions: list) -> tuple:
        if not instructions:
            return (("empty",),)
        signatures = []
        sorted_instructions = sorted(
            instructions,
            key=lambda x: self.CLEANING_METHOD_ORDER.get(
                (x.cleaning_method or "").lower().strip(), 99
            )
        )
        for instr in sorted_instructions:
            sig = (
                instr.maintenance_type or "",
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
        groups = {}
        for obj_name, instructions in object_instructions:
            signature = self._get_instruction_signature(instructions)
            if signature not in groups:
                groups[signature] = {"names": [], "instructions": instructions}
            groups[signature]["names"].append(obj_name)
        result = []
        for signature, data in groups.items():
            merged_name = ", ".join(sorted(data["names"]))
            result.append((merged_name, data["instructions"]))
        return result

    def _select_instructions_for_room(self, all_instructions: list, room_category_id: int) -> list:
        if not all_instructions:
            return []
        by_method = defaultdict(list)
        for instr in all_instructions:
            method = instr.cleaning_method or ""
            by_method[method].append(instr)
        selected = []
        for method, instrs in by_method.items():
            best = None
            for maint_level in ["основная", "поддерживающая", "генеральная"]:
                for instr in instrs:
                    if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id == room_category_id:
                        best = instr
                        break
                if best:
                    break
            if not best:
                for maint_level in ["основная", "поддерживающая", "генеральная"]:
                    for instr in instrs:
                        if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id is None:
                            best = instr
                            break
                    if best:
                        break
            if best:
                selected.append(best)
        selected.sort(
            key=lambda x: self.CLEANING_METHOD_ORDER.get(
                (x.cleaning_method or "").lower().strip(), 99
            )
        )
        return selected

    def _clone_row_formatting(self, table, source_row_idx: int, target_row_idx: int):
        """Клонирует форматирование строки (шрифт, размер, границы, жирность для колонок 4 и 9)"""
        source_row = table.rows[source_row_idx]
        target_row = table.rows[target_row_idx]
        for col_idx, source_cell in enumerate(source_row.cells):
            if col_idx >= len(target_row.cells):
                break
            target_cell = target_row.cells[col_idx]
            source_tc_pr = source_cell._tc.find(qn('w:tcPr'))
            if source_tc_pr is not None:
                target_tc_pr = target_cell._tc.find(qn('w:tcPr'))
                if target_tc_pr is not None:
                    target_cell._tc.remove(target_tc_pr)
                target_cell._tc.insert(0, copy.deepcopy(source_tc_pr))
            for p_idx, source_para in enumerate(source_cell.paragraphs):
                if p_idx >= len(target_cell.paragraphs):
                    target_para = target_cell.add_paragraph()
                else:
                    target_para = target_cell.paragraphs[p_idx]
                source_p_pr = source_para._p.find(qn('w:pPr'))
                if source_p_pr is not None:
                    target_p_pr = target_para._p.find(qn('w:pPr'))
                    if target_p_pr is not None:
                        target_para._p.remove(target_p_pr)
                    target_para._p.insert(0, copy.deepcopy(source_p_pr))
                for run in target_para.runs:
                    target_para._p.remove(run._r)
                for source_run in source_para.runs:
                    new_run = copy.deepcopy(source_run._r)
                    r_pr = new_run.find(qn('w:rPr'))
                    if r_pr is None:
                        r_pr = OxmlElement('w:rPr')
                        new_run.insert(0, r_pr)

                    # Сохраняем жирность из исходного run'а
                    source_r_pr = source_run._r.find(qn('w:rPr'))
                    if source_r_pr is not None:
                        source_b = source_r_pr.find(qn('w:b'))
                        if source_b is not None:
                            existing_b = r_pr.find(qn('w:b'))
                            if existing_b is None:
                                r_pr.append(copy.deepcopy(source_b))

                    # Принудительно добавляем bold для колонки 4 (индекс 3)
                    if col_idx == 3:
                        existing_b = r_pr.find(qn('w:b'))
                        if existing_b is None:
                            b = OxmlElement('w:b')
                            r_pr.append(b)
                    # Принудительно добавляем bold для колонки 9 (индекс 8)
                    if col_idx == 8:
                        existing_b = r_pr.find(qn('w:b'))
                        if existing_b is None:
                            b = OxmlElement('w:b')
                            r_pr.append(b)

                    sz = r_pr.find(qn('w:sz'))
                    if sz is None:
                        sz = OxmlElement('w:sz')
                        r_pr.append(sz)
                    sz.set(qn('w:val'), '14')

                    sz_cs = r_pr.find(qn('w:szCs'))
                    if sz_cs is None:
                        sz_cs = OxmlElement('w:szCs')
                        r_pr.append(sz_cs)
                    sz_cs.set(qn('w:val'), '14')

                    r_fonts = r_pr.find(qn('w:rFonts'))
                    if r_fonts is None:
                        r_fonts = OxmlElement('w:rFonts')
                        r_pr.append(r_fonts)
                    r_fonts.set(qn('w:ascii'), 'Arial')
                    r_fonts.set(qn('w:hAnsi'), 'Arial')
                    r_fonts.set(qn('w:cs'), 'Arial')
                    for t in new_run.findall(qn('w:t')):
                        t.text = ''
                    target_para._p.append(new_run)

    def _set_cell_text(self, cell, text: str, bold: bool = False):
        """Устанавливает текст в ячейку. Если bold=True, делает шрифт жирным."""
        if not cell.paragraphs:
            cell.add_paragraph()
        para = cell.paragraphs[0]
        if not para.runs:
            run_elem = OxmlElement('w:r')
            r_pr = OxmlElement('w:rPr')
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '14')
            r_pr.append(sz)
            r_fonts = OxmlElement('w:rFonts')
            r_fonts.set(qn('w:ascii'), 'Arial')
            r_fonts.set(qn('w:hAnsi'), 'Arial')
            r_pr.append(r_fonts)
            if bold:
                b = OxmlElement('w:b')
                r_pr.append(b)
            run_elem.append(r_pr)
            t_elem = OxmlElement('w:t')
            t_elem.set(qn('xml:space'), 'preserve')
            run_elem.append(t_elem)
            para._p.append(run_elem)
        else:
            run = para.runs[0]
            r_pr = run._r.find(qn('w:rPr'))
            if r_pr is None:
                r_pr = OxmlElement('w:rPr')
                run._r.insert(0, r_pr)
            existing_b = r_pr.find(qn('w:b'))
            if existing_b is not None:
                r_pr.remove(existing_b)
            if bold:
                b = OxmlElement('w:b')
                r_pr.append(b)
        run = para.runs[0]
        for t in run._r.findall(qn('w:t')):
            t.text = text
            break

    def _clear_cell_text(self, cell):
        """Очищает текст в ячейке"""
        for para in cell.paragraphs:
            for run in para.runs:
                for t in run._r.findall(qn('w:t')):
                    t.text = ''

    def generate(self, checklist_data: ChecklistData, output_path: str, mode: int = 1,
                 enterprise_products_path: str = None) -> str:
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # Заполняем помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '_____________'}"
            for para in room_cell.paragraphs:
                for run in para.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)
                    run.font.bold = True

        # === ОБРАБОТКА 4-Й СТРОКИ (предупреждение об удалении сырья) ===
        warning_row = main_table.rows[3]
        if checklist_data.room_category not in ["производственное", "складское"]:
            self._clear_cell_text(warning_row.cells[0])

        session = SessionLocal()
        category_priority = self._get_category_priority(session)

        room_category_id = None
        if hasattr(checklist_data, 'room_category') and checklist_data.room_category:
            room_category_id = self._get_room_category_id(session, checklist_data.room_category)
            if room_category_id:
                print(f"🔍 Категория помещения: {checklist_data.room_category} (id={room_category_id})")

        if enterprise_products_path and Path(enterprise_products_path).exists():
            print(f"📦 Парсинг средств предприятия: {Path(enterprise_products_path).name}")

        checked_items = checklist_data.get_checked_items()
        all_names = [item.name for item in checked_items]

        db_objects = session.query(DBObject).filter(DBObject.normalized_name.in_(all_names)).all()
        objects_dict = {obj.normalized_name: obj for obj in db_objects}

        object_ids = [obj.id for obj in db_objects]
        db_instructions = session.query(Instruction).filter(Instruction.object_id.in_(object_ids)).all()

        instructions_dict = defaultdict(list)
        for instr in db_instructions:
            instructions_dict[instr.object_id].append(instr)

        category_object_instructions = defaultdict(list)
        all_object_instructions = []
        category_order = {}
        unmatched_objects = []

        for item in checked_items:
            obj = objects_dict.get(item.name)
            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            display_name = obj.display_name if obj else item.name
            sort_priority = obj.sort_priority if obj else 999
            normalized_name = item.name  # уникальный ключ из чек-листа

            if obj and obj.id in instructions_dict:
                all_instrs = instructions_dict[obj.id]

                if normalized_name in self.SUPPORT_MAINTENANCE_OBJECTS:
                    # Специальный отбор: добавляем все уровни для каждого метода
                    by_method = defaultdict(list)
                    for instr in all_instrs:
                        method = instr.cleaning_method or ""
                        by_method[method].append(instr)
                    instructions = []
                    for method, instrs in by_method.items():
                        for maint_level in ["основная", "поддерживающая"]:
                            best = None
                            for instr in instrs:
                                if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id == room_category_id:
                                    best = instr
                                    break
                            if not best:
                                for instr in instrs:
                                    if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id is None:
                                        best = instr
                                        break
                            if best:
                                instructions.append(best)
                    instructions.sort(
                        key=lambda x: (
                            self.LEVEL_ORDER.get((x.maintenance_type or "").lower(), 99),
                            self.CLEANING_METHOD_ORDER.get((x.cleaning_method or "").lower().strip(), 99)
                        )
                    )
                else:
                    # Стандартный отбор (одна инструкция на метод)
                    instructions = self._select_instructions_for_room(all_instrs, room_category_id)

                if instructions:
                    category_object_instructions[cat_name].append((display_name, instructions, normalized_name))
                    all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                else:
                    unmatched_objects.append(display_name)
            elif obj:
                unmatched_objects.append(display_name)
            else:
                unmatched_objects.append(item.name)

        session.close()

        rows_data = []
        merge_info = []

        if mode == 1:
            sorted_categories = sorted(category_object_instructions.keys(), key=lambda x: category_order.get(x, 999))
            for cat_name in sorted_categories:
                items = category_object_instructions[cat_name]
                if not items:
                    continue
                items.sort(key=lambda x: x[0])  # сортировка по display_name
                rows_data.append(('category', cat_name, None))
                for display_name, instructions, normalized_name in items:
                    group_start_row = len(rows_data)
                    if instructions:
                        for i, instr in enumerate(instructions):
                            cell_text = display_name if i == 0 else ""
                            rows_data.append(('object', cell_text, instr, normalized_name))
                    else:
                        rows_data.append(('object', display_name, None, normalized_name))
                    group_end_row = len(rows_data) - 1
                    if group_end_row > group_start_row:
                        merge_info.append((group_start_row, group_end_row))
        elif mode == 2:
            all_object_instructions.sort(key=lambda x: (x[2], x[0]))
            for display_name, instructions, sort_priority, normalized_name in all_object_instructions:
                group_start_row = len(rows_data)
                if instructions:
                    for i, instr in enumerate(instructions):
                        cell_text = display_name if i == 0 else ""
                        rows_data.append(('object', cell_text, instr, normalized_name))
                else:
                    rows_data.append(('object', display_name, None, normalized_name))
                group_end_row = len(rows_data) - 1
                if group_end_row > group_start_row:
                    merge_info.append((group_start_row, group_end_row))
        elif mode == 3:
            priority_groups = defaultdict(list)
            for display_name, instructions, sort_priority, normalized_name in all_object_instructions:
                priority_groups[sort_priority].append((display_name, instructions, normalized_name))

            for priority in sorted(priority_groups.keys()):
                items = priority_groups[priority]
                # Группируем по инструкциям, сохраняя normalized_name первого объекта в группе
                groups = {}
                for dn, instr, nn in items:
                    signature = self._get_instruction_signature(instr)
                    if signature not in groups:
                        groups[signature] = {"names": [], "instructions": instr, "normalized_name": nn}
                    groups[signature]["names"].append(dn)

                for sig, data in groups.items():
                    merged_name = ", ".join(sorted(data["names"]))
                    group_start_row = len(rows_data)
                    if data["instructions"]:
                        for i, instr in enumerate(data["instructions"]):
                            cell_text = merged_name if i == 0 else ""
                            rows_data.append(('object', cell_text, instr, data["normalized_name"]))
                    else:
                        rows_data.append(('object', merged_name, None, data["normalized_name"]))
                    group_end_row = len(rows_data) - 1
                    if group_end_row > group_start_row:
                        merge_info.append((group_start_row, group_end_row))

        if unmatched_objects:
            rows_data.append(('category', 'Объекты без инструкций (требуют настройки)', None))
            for name in sorted(unmatched_objects):
                rows_data.append(('object', name, None, None))

        # === ОЧИСТКА ТАБЛИЦЫ ===
        start_row = 8
        tbl = main_table._tbl
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tr_elements = tbl.findall('.//w:tr', namespaces=ns)
        while len(tr_elements) > start_row + 1:
            tbl.remove(tr_elements[-1])
            tr_elements = tbl.findall('.//w:tr', namespaces=ns)

        # === ДОБАВЛЕНИЕ СТРОК С КЛОНИРОВАНИЕМ ===
        for _ in range(len(rows_data)):
            main_table.add_row()
            self._clone_row_formatting(main_table, start_row, len(main_table.rows) - 1)

        # Удаляем эталонную строку
        row_to_remove = main_table.rows[start_row]
        tbl.remove(row_to_remove._tr)

        # === ЗАПОЛНЕНИЕ С ПОДСТАНОВКОЙ СРЕДСТВ ИЗ ЧЕК-ЛИСТА ===
        current_row = start_row
        for row_info in rows_data:
            row = main_table.rows[current_row]

            if row_info[0] == 'category':
                self._merge_cells_horizontal(row, 0, 11)
                self._set_cell_text(row.cells[0], row_info[1])
                if row.cells[0].paragraphs and row.cells[0].paragraphs[0].runs:
                    row.cells[0].paragraphs[0].runs[0].font.bold = True

            elif row_info[0] == 'object':
                obj_name, instr = row_info[1], row_info[2]
                normalized_name = row_info[3] if len(row_info) > 3 else None
                self._set_cell_text(row.cells[0], obj_name)

                if instr:
                    cleaning_method = instr.cleaning_method or ""
                    product_name = instr.product_name or ""
                    concentration = instr.concentration or ""
                    cleaning_technique = instr.cleaning_technique or ""
                    extra_method_text = None

                    # === ПЕРЕОПРЕДЕЛЕНИЕ ИЗ ЧЕК-ЛИСТА ===
                    if cleaning_method == "мойка":
                        if normalized_name and normalized_name in self.FLOOR_OBJECTS:
                            if checklist_data.floor_cleaning_product:
                                product_name = checklist_data.floor_cleaning_product
                                concentration = checklist_data.floor_cleaning_concentration
                                extra_method_text = checklist_data.floor_cleaning_method_text
                        elif normalized_name and normalized_name in self.THERMAL_OBJECTS:
                            if checklist_data.thermal_cleaning_product:
                                product_name = checklist_data.thermal_cleaning_product
                                concentration = checklist_data.thermal_cleaning_concentration
                                extra_method_text = checklist_data.thermal_cleaning_method_text
                        elif normalized_name and normalized_name in self.GLASS_OBJECTS:
                            if checklist_data.glass_cleaning_product:
                                product_name = checklist_data.glass_cleaning_product
                                concentration = checklist_data.glass_cleaning_concentration
                                extra_method_text = checklist_data.glass_cleaning_method_text
                        else:
                            # Общее моющее средство
                            if checklist_data.cleaning_product:
                                product_name = checklist_data.cleaning_product
                            if checklist_data.cleaning_concentration:
                                concentration = checklist_data.cleaning_concentration
                            if checklist_data.cleaning_method_text:
                                extra_method_text = checklist_data.cleaning_method_text
                    elif cleaning_method == "дезинфекция":
                        if checklist_data.disinfection_product:
                            product_name = checklist_data.disinfection_product
                        if checklist_data.disinfection_concentration:
                            concentration = checklist_data.disinfection_concentration
                        if checklist_data.disinfection_method_text:
                            extra_method_text = checklist_data.disinfection_method_text

                    self._set_cell_text(row.cells[1], cleaning_method)
                    self._set_cell_text(row.cells[2], instr.instruction_number or "")
                    self._set_cell_text(row.cells[3], product_name if product_name else "_____________", bold=bool(product_name))
                    self._set_cell_text(row.cells[4], cleaning_technique)
                    self._set_cell_text(row.cells[5], concentration if concentration else "_____________")

                    if extra_method_text:
                        para = row.cells[5].paragraphs[0]
                        run_br = para.add_run()
                        br = OxmlElement('w:br')
                        run_br._r.append(br)
                        run_text = para.add_run(extra_method_text)
                        run_text.font.name = 'Arial'
                        run_text.font.size = Pt(7)

                    self._set_cell_text(row.cells[6], instr.temperature if instr.temperature else "_____________")
                    self._set_cell_text(row.cells[7], instr.exposure_time if instr.exposure_time else "_____________")
                    self._set_cell_text(row.cells[8], instr.inventory or "", bold=bool(instr.inventory))
                    self._set_cell_text(row.cells[9], instr.frequency or "")
                    self._set_cell_text(row.cells[10], instr.executor or "")
                    self._set_cell_text(row.cells[11], instr.control_method or "")

                    if product_name:
                        color = self._get_color_for_product(product_name)
                        if color:
                            self._set_cell_background(row.cells[3], color)

            current_row += 1

        # === ВЕРТИКАЛЬНОЕ ОБЪЕДИНЕНИЕ ===
        for group_start, group_end in merge_info:
            actual_start = start_row + group_start
            actual_end = start_row + group_end
            if actual_end > actual_start:
                for row in range(actual_start + 1, actual_end + 1):
                    main_table.cell(row, 0).text = ""
                self._merge_cells_vertical(main_table, 0, actual_start, actual_end)
                for col in [8, 9, 10, 11]:
                    self._merge_adjacent_equal_cells(main_table, col, actual_start, actual_end)

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()
        doc.save(str(output_file))
        return str(output_file)