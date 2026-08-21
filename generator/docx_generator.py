# generator/docx_generator.py

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pathlib import Path
from parser.models import ChecklistData
from db.models import (Instruction, Category as DBCategory, Object as DBObject,
                       RoomCategory, Product, ObjectProperty, ObjectGroup,
                       CleaningMethodOrder, InventoryColor)
from db.database import SessionLocal
from collections import defaultdict
import re
import copy


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    # Порядок уровней обработки для вывода
    LEVEL_ORDER = {"основная": 1, "поддерживающая": 2, "генеральная": 3}

    # Заглушки, которые не должны подставляться из чек-листа
    SKIP_PLACEHOLDERS = {
        "Выберите элемент.", "Не выбрано", "Выберите элемент",
        "", None
    }

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE
        # Кэши конфигурации, загружаемые из БД
        self._config_loaded = False
        # ===== Кэш цветов инвентаря =====
        self._inventory_colors = {}  # {name: hex_color}
        self._load_inventory_colors()

    def _load_inventory_colors(self):
        """Загружает цвета инвентаря из БД в словарь {name: hex_color}."""
        session = SessionLocal()
        try:
            colors = session.query(InventoryColor).all()
            self._inventory_colors = {c.name: c.hex_color for c in colors}
        finally:
            session.close()

    def _load_config(self, session):
        """Загружает конфигурацию из БД (CleaningMethodOrder, ObjectProperty, ObjectGroup)"""
        if self._config_loaded:
            return

        # 1. Порядок методов очистки
        methods = session.query(CleaningMethodOrder).order_by(CleaningMethodOrder.sort_order).all()
        self.cleaning_method_order = {m.method_name: m.sort_order for m in methods}

        # 2. Свойства объектов
        props = session.query(ObjectProperty).join(DBObject).all()
        self.split_objects = set()
        self.multi_method_objects = set()
        self.support_objects = set()
        self.floor_objects = set()
        self.glass_objects = set()
        self.thermal_objects = set()

        for prop in props:
            name = prop.object.normalized_name
            if prop.is_split:
                self.split_objects.add(name)
            if prop.is_multi_method:
                self.multi_method_objects.add(name)
            if prop.has_support_maintenance:
                self.support_objects.add(name)
            if prop.special_product_type == 'floor':
                self.floor_objects.add(name)
            elif prop.special_product_type == 'glass':
                self.glass_objects.add(name)
            elif prop.special_product_type == 'thermal':
                self.thermal_objects.add(name)

        # 3. Группы объектов
        groups = session.query(ObjectGroup).all()
        self.group_headers = {}  # frozenset -> header
        temp_groups = defaultdict(set)
        for g in groups:
            if g.object:
                temp_groups[g.group_name].add(g.object.normalized_name)
        for header, obj_set in temp_groups.items():
            self.group_headers[frozenset(obj_set)] = header

        self._config_loaded = True

    def _load_product_colors(self, session) -> dict:
        """Загружает цвета средств из таблицы products."""
        products = session.query(Product).all()
        return {p.name: p.color for p in products if p.color}

    def _clean_text(self, val: str) -> str:
        """Очищает строку от лишних пробелов и переводов строк (но не XML-элементов)."""
        if not val:
            return ""
        return val.strip().replace('\n', ' ').replace('\r', ' ')

    def _normalize_product_name(self, name: str) -> str:
        if not name:
            return ""
        normalized = re.sub(r'[^\w\s\-]', '', name)
        normalized = normalized.lower()
        normalized = re.sub(r'\s+', ' ', normalized)
        return normalized.strip()

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

    def _remove_extra_paragraphs_in_cell(self, cell):
        """Удаляет все пустые параграфы в ячейке, оставляя только один."""
        tc = cell._tc
        paragraphs = tc.findall(qn('w:p'))
        if not paragraphs:
            return
        first_p = paragraphs[0]
        for p in paragraphs[1:]:
            tc.remove(p)
        if not first_p.findall(qn('w:r')):
            run = OxmlElement('w:r')
            run_pr = OxmlElement('w:rPr')
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '14')
            run_pr.append(sz)
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), 'Arial')
            rFonts.set(qn('w:hAnsi'), 'Arial')
            run_pr.append(rFonts)
            run.append(run_pr)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = ''
            run.append(t)
            first_p.append(run)

    def _merge_cells_vertical(self, table, col: int, start_row: int, end_row: int):
        if start_row >= end_row:
            return
        start_cell = table.cell(start_row, col)
        for row in range(start_row + 1, end_row + 1):
            start_cell.merge(table.cell(row, col))
        self._remove_extra_paragraphs_in_cell(start_cell)

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
            key=lambda x: self.cleaning_method_order.get(
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
                instr.surface_type or "",
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
            key=lambda x: self.cleaning_method_order.get(
                (x.cleaning_method or "").lower().strip(), 99
            )
        )
        return selected

    def _select_all_instructions_for_room(self, all_instructions: list, room_category_id: int) -> list:
        """Выбирает все подходящие инструкции (может быть несколько на один метод).
        Порядок вывода: сначала по уровню обслуживания, внутри уровня – чередование методов.
        """
        if not all_instructions:
            return []

        # 1. Отбираем инструкции с учётом категории помещения и уровня (все уровни)
        by_method = defaultdict(list)
        for instr in all_instructions:
            method = instr.cleaning_method or ""
            by_method[method].append(instr)

        selected = []
        for method, instrs in by_method.items():
            # Сначала пробуем найти инструкции, специфичные для помещения
            specific_instrs = [i for i in instrs if i.room_category_id == room_category_id]
            if specific_instrs:
                selected.extend(specific_instrs)
            else:
                # Если нет специфичных, берём все общие
                common_instrs = [i for i in instrs if i.room_category_id is None]
                selected.extend(common_instrs)

        # 2. Группируем по уровню обслуживания
        level_groups = defaultdict(list)
        for instr in selected:
            level = (instr.maintenance_type or "").lower()
            level_groups[level].append(instr)

        result = []
        for level in ["основная", "поддерживающая", "генеральная"]:
            instrs = level_groups.get(level, [])
            if not instrs:
                continue

            # Группируем инструкции одного уровня по методам
            method_groups = defaultdict(list)
            for instr in instrs:
                method = instr.cleaning_method or ""
                method_groups[method].append(instr)

            # Сортируем методы по CLEANING_METHOD_ORDER
            ordered_methods = sorted(method_groups.keys(),
                                     key=lambda m: self.cleaning_method_order.get(m.lower().strip(), 99))

            # Чередуем инструкции внутри уровня
            while True:
                added = False
                for method in ordered_methods:
                    if method_groups[method]:
                        result.append(method_groups[method].pop(0))
                        added = True
                if not added:
                    break

        return result

    def _select_split_instructions(self, all_instructions: list, room_category_id: int) -> dict:
        """Для split-объектов: группирует инструкции по surface_type с логикой уровня обслуживания."""
        result = {}
        surface_types = ["внешняя", "внутренняя", "очистка от мин. отложений"]
        for sf in surface_types:
            sf_instrs = [i for i in all_instructions if (i.surface_type or "").lower() == sf]
            if not sf_instrs:
                continue
            selected = self._select_instructions_for_room(sf_instrs, room_category_id)
            if selected:
                result[sf] = selected
        return result

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
        text = self._clean_text(text)
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

    def _set_product_cell(self, cell, text: str, bold: bool = False):
        """Устанавливает текст в ячейку, обрабатывая надстрочный знак ®."""
        text = self._clean_text(text)
        if not cell.paragraphs:
            cell.add_paragraph()
        para = cell.paragraphs[0]
        # Очищаем параграф перед вставкой новых runs
        for r in para.runs:
            para._p.remove(r._r)

        # Разбиваем по символу ®
        parts = text.split('®')
        for i, part in enumerate(parts):
            if part:
                run = para.add_run(part)
                run.font.name = 'Arial'
                run.font.size = Pt(7)
                if bold:
                    run.bold = True
            if i < len(parts) - 1:
                # Вставляем надстрочный ®
                run_reg = para.add_run('®')
                run_reg.font.name = 'Arial'
                run_reg.font.size = Pt(7)
                if bold:
                    run_reg.bold = True
                run_reg.font.superscript = True

    def _clear_cell_text(self, cell):
        """Очищает текст в ячейке"""
        for para in cell.paragraphs:
            for run in para.runs:
                for t in run._r.findall(qn('w:t')):
                    t.text = ''

    def _is_valid_product_name(self, name: str) -> bool:
        """Проверяет, похоже ли значение на название средства (не концентрация/метод/заглушка)."""
        if not name:
            return False
        name_lower = name.lower().strip()
        if name_lower in ('выберите элемент.', 'не выбрано', 'выберите элемент', ''):
            return False
        # Паттерны, указывающие на концентрацию или метод
        patterns = [
            r'%', r'дозирующая система', r'ручной', r'система', r'протирание',
            r'щётка', r'ветошь', r'губка', r'погружение', r'замачивание',
            r'орошение', r'распыление', r'протирка', r'обработка', r'раствор'
        ]
        for pattern in patterns:
            if re.search(pattern, name_lower):
                return False
        if re.match(r'^[0-9.,%\-–\s]+$', name_lower):
            return False
        return True

    def generate(self, checklist_data: ChecklistData, output_path: str, mode: int = 1,
                 enterprise_products_path: str = None, progress_callback=None) -> str:
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # Заполняем помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '___________'}"
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
        # Загрузка конфигурации из БД
        self._load_config(session)

        category_priority = self._get_category_priority(session)
        product_colors = self._load_product_colors(session)

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
            normalized_name = item.name

            if obj and obj.id in instructions_dict:
                all_instrs = instructions_dict[obj.id]

                if normalized_name in self.split_objects:
                    instructions = self._select_split_instructions(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('split', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                elif normalized_name in self.support_objects:
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
                            self.cleaning_method_order.get((x.cleaning_method or "").lower().strip(), 99)
                        )
                    )
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                elif normalized_name in self.multi_method_objects:
                    instructions = self._select_all_instructions_for_room(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                else:
                    instructions = self._select_instructions_for_room(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
            elif obj:
                unmatched_objects.append(display_name)
            else:
                unmatched_objects.append(item.name)

        session.close()

        rows_data = []
        merge_info_object = []
        merge_info_columns = []
        surface_merge_info = []

        if mode == 1:
            sorted_categories = sorted(category_object_instructions.keys(), key=lambda x: category_order.get(x, 999))
            for cat_name in sorted_categories:
                items = category_object_instructions[cat_name]
                if not items:
                    continue
                items.sort(key=lambda x: x[1])
                rows_data.append(('category', cat_name, None))
                for typ, display_name, instructions, normalized_name in items:
                    if typ == 'split':
                        surfaces = ["внешняя", "внутренняя", "очистка от мин. отложений"]
                        first_surface = True
                        group_start_row = len(rows_data)
                        for sf in surfaces:
                            if sf not in instructions:
                                continue
                            sf_instrs = instructions[sf]
                            sf_label = {"внешняя": "Внешние поверхности",
                                        "внутренняя": "Внутренние поверхности",
                                        "очистка от мин. отложений": "Очистка от минеральных отложений"}[sf]
                            if first_surface:
                                rows_data.append(('section_header', sf_label, display_name))
                                first_surface = False
                            else:
                                rows_data.append(('section_header', sf_label, None))
                            surface_start = len(rows_data)
                            for instr in sf_instrs:
                                rows_data.append(('object', "", instr, normalized_name, instr.subgroup))
                            surface_end = len(rows_data) - 1
                            if surface_end >= surface_start:
                                surface_merge_info.append((surface_start, surface_end))
                        group_end_row = len(rows_data) - 1
                        merge_info_object.append((group_start_row, group_end_row))
                    else:
                        group_start_row = len(rows_data)
                        if instructions:
                            for i, instr in enumerate(instructions):
                                cell_text = display_name if i == 0 else ""
                                rows_data.append(('object', cell_text, instr, normalized_name, instr.subgroup))
                        else:
                            rows_data.append(('object', display_name, None, normalized_name, None))
                        group_end_row = len(rows_data) - 1
                        if group_end_row > group_start_row:
                            merge_info_object.append((group_start_row, group_end_row))
                            merge_info_columns.append((group_start_row, group_end_row))
        elif mode == 2:
            all_object_instructions.sort(key=lambda x: (x[2], x[0]))
            for display_name, instructions, sort_priority, normalized_name in all_object_instructions:
                if isinstance(instructions, dict):
                    surfaces = ["внешняя", "внутренняя", "очистка от мин. отложений"]
                    first_surface = True
                    group_start_row = len(rows_data)
                    for sf in surfaces:
                        if sf not in instructions:
                            continue
                        sf_instrs = instructions[sf]
                        sf_label = {"внешняя": "Внешние поверхности",
                                    "внутренняя": "Внутренние поверхности",
                                    "очистка от мин. отложений": "Очистка от минеральных отложений"}[sf]
                        if first_surface:
                            rows_data.append(('section_header', sf_label, display_name))
                            first_surface = False
                        else:
                            rows_data.append(('section_header', sf_label, None))
                        surface_start = len(rows_data)
                        for instr in sf_instrs:
                            rows_data.append(('object', "", instr, normalized_name, instr.subgroup))
                        surface_end = len(rows_data) - 1
                        if surface_end >= surface_start:
                            surface_merge_info.append((surface_start, surface_end))
                    group_end_row = len(rows_data) - 1
                    merge_info_object.append((group_start_row, group_end_row))
                else:
                    group_start_row = len(rows_data)
                    if instructions:
                        for i, instr in enumerate(instructions):
                            cell_text = display_name if i == 0 else ""
                            rows_data.append(('object', cell_text, instr, normalized_name, instr.subgroup))
                    else:
                        rows_data.append(('object', display_name, None, normalized_name, None))
                    group_end_row = len(rows_data) - 1
                    if group_end_row > group_start_row:
                        merge_info_object.append((group_start_row, group_end_row))
                        merge_info_columns.append((group_start_row, group_end_row))
        elif mode == 3:
            inserted_headers = set()
            priority_groups = defaultdict(list)
            for display_name, instructions, sort_priority, normalized_name in all_object_instructions:
                priority_groups[sort_priority].append((display_name, instructions, normalized_name))

            for priority in sorted(priority_groups.keys()):
                items = priority_groups[priority]
                normal_items = [(dn, instr, nn) for dn, instr, nn in items if not isinstance(instr, dict)]
                split_items = [(dn, instr, nn) for dn, instr, nn in items if isinstance(instr, dict)]

                groups = {}
                for dn, instr, nn in normal_items:
                    signature = self._get_instruction_signature(instr)
                    if signature not in groups:
                        groups[signature] = {"names": [], "instructions": instr, "normalized_name": nn}
                    groups[signature]["names"].append(dn)

                for sig, data in groups.items():
                    current_nn = data["normalized_name"]
                    for group_set, header_text in self.group_headers.items():
                        if current_nn in group_set and group_set not in inserted_headers:
                            rows_data.append(('group_header', header_text, None))
                            inserted_headers.add(group_set)
                            break
                    merged_name = ", ".join(sorted(data["names"]))
                    group_start_row = len(rows_data)
                    if data["instructions"]:
                        for i, instr in enumerate(data["instructions"]):
                            cell_text = merged_name if i == 0 else ""
                            rows_data.append(('object', cell_text, instr, data["normalized_name"], instr.subgroup))
                    else:
                        rows_data.append(('object', merged_name, None, data["normalized_name"], None))
                    group_end_row = len(rows_data) - 1
                    if group_end_row > group_start_row:
                        merge_info_object.append((group_start_row, group_end_row))
                        merge_info_columns.append((group_start_row, group_end_row))

                for dn, split_instr, nn in split_items:
                    for group_set, header_text in self.group_headers.items():
                        if nn in group_set and group_set not in inserted_headers:
                            rows_data.append(('group_header', header_text, None))
                            inserted_headers.add(group_set)
                            break
                    surfaces = ["внешняя", "внутренняя", "очистка от мин. отложений"]
                    first_surface = True
                    group_start_row = len(rows_data)
                    for sf in surfaces:
                        if sf not in split_instr:
                            continue
                        sf_instrs = split_instr[sf]
                        sf_label = {"внешняя": "Внешние поверхности",
                                    "внутренняя": "Внутренние поверхности",
                                    "очистка от мин. отложений": "Очистка от минеральных отложений"}[sf]
                        if first_surface:
                            rows_data.append(('section_header', sf_label, dn))
                            first_surface = False
                        else:
                            rows_data.append(('section_header', sf_label, None))
                        surface_start = len(rows_data)
                        for instr in sf_instrs:
                            rows_data.append(('object', "", instr, nn, instr.subgroup))
                        surface_end = len(rows_data) - 1
                        if surface_end >= surface_start:
                            surface_merge_info.append((surface_start, surface_end))
                    group_end_row = len(rows_data) - 1
                    merge_info_object.append((group_start_row, group_end_row))

        if unmatched_objects:
            rows_data.append(('category', 'Объекты без инструкций (требуют настройки)', None))
            for name in sorted(unmatched_objects):
                rows_data.append(('object', name, None, None, None))

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

        row_to_remove = main_table.rows[start_row]
        tbl.remove(row_to_remove._tr)

        # === ЗАПОЛНЕНИЕ С НОВОЙ ЛОГИКОЙ ПОДСТАНОВКИ ===
        current_row = start_row
        for row_info in rows_data:
            row = main_table.rows[current_row]

            if row_info[0] == 'category':
                self._merge_cells_horizontal(row, 0, 11)
                self._set_cell_text(row.cells[0], row_info[1])
                if row.cells[0].paragraphs and row.cells[0].paragraphs[0].runs:
                    row.cells[0].paragraphs[0].runs[0].font.bold = True

            elif row_info[0] == 'group_header':
                self._merge_cells_horizontal(row, 0, 11)
                self._set_cell_text(row.cells[0], row_info[1], bold=True)
                self._set_cell_background(row.cells[0], "D2D2D2")
                for para in row.cells[0].paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(7)

            elif row_info[0] == 'section_header':
                if len(row_info) > 2 and row_info[2] is not None:
                    self._set_cell_text(row.cells[0], row_info[2])
                else:
                    self._set_cell_text(row.cells[0], "")
                self._merge_cells_horizontal(row, 1, 11)
                self._set_cell_text(row.cells[1], row_info[1], bold=True)
                self._set_cell_background(row.cells[1], "D2D2D2")
                for para in row.cells[1].paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(7)

            elif row_info[0] == 'object':
                obj_name, instr = row_info[1], row_info[2]
                normalized_name = row_info[3] if len(row_info) > 3 else None
                self._set_cell_text(row.cells[0], obj_name)

                if instr:
                    cleaning_method = self._clean_text(instr.cleaning_method or "")
                    cleaning_technique = self._clean_text(instr.cleaning_technique or "")

                    # ---------- ЛОГИКА ОПРЕДЕЛЕНИЯ product_name, concentration, extra_method ----------
                    db_product = self._clean_text(instr.product_name or "")
                    db_concentration = self._clean_text(instr.concentration or "")
                    db_method = self._clean_text(instr.application_method or "")

                    has_db_product = db_product and self._is_valid_product_name(db_product)

                    if has_db_product:
                        final_product = db_product
                        final_concentration = db_concentration
                        final_extra_method = db_method
                    else:
                        final_product = ""
                        final_concentration = ""
                        final_extra_method = ""
                        if normalized_name and normalized_name not in self.split_objects:
                            checklist_product = None
                            checklist_concentration = None
                            checklist_method = None
                            if cleaning_method == "мойка":
                                if normalized_name in self.floor_objects:
                                    checklist_product = checklist_data.floor_cleaning_product
                                    checklist_concentration = checklist_data.floor_cleaning_concentration
                                    checklist_method = checklist_data.floor_cleaning_method_text
                                elif normalized_name in self.thermal_objects:
                                    checklist_product = checklist_data.thermal_cleaning_product
                                    checklist_concentration = checklist_data.thermal_cleaning_concentration
                                    checklist_method = checklist_data.thermal_cleaning_method_text
                                elif normalized_name in self.glass_objects:
                                    checklist_product = checklist_data.glass_cleaning_product
                                    checklist_concentration = checklist_data.glass_cleaning_concentration
                                    checklist_method = checklist_data.glass_cleaning_method_text
                                else:
                                    checklist_product = checklist_data.cleaning_product
                                    checklist_concentration = checklist_data.cleaning_concentration
                                    checklist_method = checklist_data.cleaning_method_text
                            elif cleaning_method == "дезинфекция":
                                checklist_product = checklist_data.disinfection_product
                                checklist_concentration = checklist_data.disinfection_concentration
                                checklist_method = checklist_data.disinfection_method_text

                            if checklist_product and self._is_valid_product_name(checklist_product):
                                final_product = self._clean_text(checklist_product)
                                final_concentration = self._clean_text(checklist_concentration) if checklist_concentration else ""
                                final_extra_method = self._clean_text(checklist_method) if checklist_method else ""

                    if not final_product:
                        final_product = None

                    self._set_cell_text(row.cells[1], cleaning_method)
                    self._set_cell_text(row.cells[2], self._clean_text(instr.instruction_number or ""))
                    if final_product:
                        self._set_product_cell(row.cells[3], final_product, bold=True)
                    else:
                        self._set_product_cell(row.cells[3], "_____________________", bold=False)

                    self._set_cell_text(row.cells[4], cleaning_technique)

                    # Колонка 6: концентрация + метод разведения (с переносом строки)
                    para = row.cells[5].paragraphs[0]
                    for r in para.runs:
                        para._p.remove(r._r)

                    conc_text = final_concentration if final_concentration else "___________________"
                    run_conc = para.add_run(conc_text)
                    run_conc.font.name = 'Arial'
                    run_conc.font.size = Pt(7)

                    if final_extra_method:
                        run_br = para.add_run()
                        br = OxmlElement('w:br')
                        run_br._r.append(br)
                        run_method = para.add_run(final_extra_method)
                        run_method.font.name = 'Arial'
                        run_method.font.size = Pt(7)

                    self._set_cell_text(row.cells[6], self._clean_text(instr.temperature or "") if instr.temperature else "___________")
                    self._set_cell_text(row.cells[7], self._clean_text(instr.exposure_time or "") if instr.exposure_time else "___________")

                    # ===== КОЛОНКА 8: ИНВЕНТАРЬ =====
                    if normalized_name and normalized_name in self.split_objects:
                        # Для split-объектов инвентарь берётся из инструкции
                        inv_text = self._clean_text(instr.inventory or "")
                        if inv_text and inv_text.lower() in self._inventory_colors:
                            inv_color = self._inventory_colors[inv_text.lower()]
                            self._set_cell_text(row.cells[8], inv_text, bold=True)
                            self._set_cell_background(row.cells[8], inv_color)
                        else:
                            self._set_cell_text(row.cells[8], inv_text if inv_text else "___________", bold=bool(inv_text))
                    else:
                        # Для обычных объектов – цвет из чек-листа
                        if checklist_data.inventory_color:
                            inv_color = self._inventory_colors.get(checklist_data.inventory_color)
                            if inv_color:
                                self._set_cell_text(row.cells[8], checklist_data.inventory_color, bold=True)
                                self._set_cell_background(row.cells[8], inv_color)
                            else:
                                self._set_cell_text(row.cells[8], "промаркированный", bold=True)
                        else:
                            self._set_cell_text(row.cells[8], "промаркированный", bold=True)

                    self._set_cell_text(row.cells[9], self._clean_text(instr.frequency or ""))
                    self._set_cell_text(row.cells[10], self._clean_text(instr.executor or ""))
                    self._set_cell_text(row.cells[11], self._clean_text(instr.control_method or ""))

                    if final_product:
                        color = product_colors.get(final_product)
                        if color:
                            self._set_cell_background(row.cells[3], color)

            current_row += 1

        # === ВЕРТИКАЛЬНОЕ ОБЪЕДИНЕНИЕ ===
        for group_start, group_end in merge_info_object:
            actual_start = start_row + group_start
            actual_end = start_row + group_end
            if actual_end > actual_start:
                for row in range(actual_start + 1, actual_end + 1):
                    main_table.cell(row, 0).text = ""
                self._merge_cells_vertical(main_table, 0, actual_start, actual_end)

        # ===== ОБЪЕДИНЕНИЕ КОЛОНОК 8,9,10,11 (как раньше) =====
        for group_start, group_end in merge_info_columns:
            actual_start = start_row + group_start
            actual_end = start_row + group_end
            if actual_end > actual_start:
                for col in [8, 9, 10, 11]:
                    self._merge_adjacent_equal_cells(main_table, col, actual_start, actual_end)

        # ===== КОЛОНКА 2: ОБЪЕДИНЕНИЕ ПО (maintenance_type, subgroup) =====
        for group_start, group_end in merge_info_columns:
            actual_start = start_row + group_start
            actual_end = start_row + group_end
            if actual_end > actual_start:
                # Собираем для каждой строки ключ группировки (maintenance_type, subgroup)
                keys = []
                for row_idx in range(actual_start, actual_end + 1):
                    row_data_index = row_idx - start_row
                    if row_data_index < len(rows_data):
                        row_data = rows_data[row_data_index]
                        instr = row_data[2] if len(row_data) > 2 and row_data[0] == 'object' else None
                        if instr:
                            maint = (instr.maintenance_type or "").lower()
                            subgroup = (instr.subgroup or "").strip() or "_default_"
                            keys.append((maint, subgroup))
                        else:
                            keys.append(None)
                    else:
                        keys.append(None)

                unique_keys = set([k for k in keys if k is not None])
                if len(unique_keys) <= 1:
                    # Все строки одного ключа – объединяем колонку 2 целиком
                    self._merge_adjacent_equal_cells(main_table, 2, actual_start, actual_end)
                else:
                    # Разбиваем на поддиапазоны по смене ключа
                    current_key = None
                    sub_start = actual_start
                    for row_idx in range(actual_start, actual_end + 1):
                        row_data_index = row_idx - start_row
                        if row_data_index < len(rows_data):
                            row_data = rows_data[row_data_index]
                            instr = row_data[2] if len(row_data) > 2 and row_data[0] == 'object' else None
                            key = None
                            if instr:
                                maint = (instr.maintenance_type or "").lower()
                                subgroup = (instr.subgroup or "").strip() or "_default_"
                                key = (maint, subgroup)
                        else:
                            key = None

                        if key != current_key:
                            if current_key is not None:
                                if sub_start <= row_idx - 1:
                                    self._merge_adjacent_equal_cells(main_table, 2, sub_start, row_idx - 1)
                            current_key = key
                            sub_start = row_idx
                    if current_key is not None:
                        if sub_start <= actual_end:
                            self._merge_adjacent_equal_cells(main_table, 2, sub_start, actual_end)

        # === ОБЪЕДИНЕНИЕ ДЛЯ SPLIT-ОБЪЕКТОВ (не меняется) ===
        for group_start, group_end in surface_merge_info:
            actual_start = start_row + group_start
            actual_end = start_row + group_end
            if actual_end > actual_start:
                for col in [2, 8, 9, 10, 11]:
                    self._merge_adjacent_equal_cells(main_table, col, actual_start, actual_end)

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()
        doc.save(str(output_file))
        return str(output_file)