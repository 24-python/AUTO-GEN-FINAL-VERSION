from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path


def set_russian_language(run):
    """Устанавливает русский язык для проверки орфографии через XML"""
    rPr = run._r.get_or_add_rPr()
    # Удаляем старые языковые настройки
    for lang in rPr.findall(qn('w:lang')):
        rPr.remove(lang)
    # Добавляем русский язык
    lang = OxmlElement('w:lang')
    lang.set(qn('w:val'), 'ru-RU')
    rPr.append(lang)


def set_cell_font(cell, text, font_name='Arial', font_size=Pt(9), bold=False, alignment=None):
    """Устанавливает текст и шрифт в ячейке"""
    cell.text = text
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = font_size
            run.font.bold = bold
            set_russian_language(run)  # устанавливаем русский язык
        if alignment:
            paragraph.alignment = alignment


def set_cell_background(cell, hex_color):
    """Устанавливает цвет фона ячейки"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_width(cell, width_cm):
    """Устанавливает ширину ячейки в сантиметрах"""
    width_emu = int(width_cm * 36000)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_emu))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)


def set_document_language(doc, lang='ru-RU'):
    """Устанавливает язык для всего документа через XML"""
    try:
        # Устанавливаем язык для стилей
        styles = doc.styles
        for style in styles:
            try:
                if style.font:
                    # Через XML для стиля
                    style_element = style._element
                    rPr = style_element.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = OxmlElement('w:rPr')
                        style_element.append(rPr)
                    lang_elem = OxmlElement('w:lang')
                    lang_elem.set(qn('w:val'), lang)
                    # Удаляем старые
                    for old in rPr.findall(qn('w:lang')):
                        rPr.remove(old)
                    rPr.append(lang_elem)
            except:
                pass
    except Exception as e:
        print(f"Предупреждение: не удалось установить язык для стилей: {e}")


def create_template():
    doc = Document()

    # Альбомная ориентация
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    # Устанавливаем русский язык для документа
    set_document_language(doc)

    # ========== ТАБЛИЦА ЛЕГЕНДЫ ==========
    legend_table = doc.add_table(rows=1, cols=10)
    legend_table.style = 'Table Grid'

    # Данные для четных колонок
    legend_texts = [
        "нейтральное средство",
        "кислотное средство",
        "щелочное средство",
        "дезинфектант",
        "средство на безводной основе"
    ]

    colors = ["99FF99", "CCCCFF", "FFCC99", "CCFFFF", "8FBFFA"]

    for col in range(10):
        cell = legend_table.rows[0].cells[col]
        if col % 2 == 0:
            set_cell_width(cell, 1.0)
            set_cell_background(cell, colors[col // 2])
            cell.text = ""
        else:
            set_cell_width(cell, 4.6)
            idx = (col - 1) // 2
            set_cell_font(cell, legend_texts[idx], font_size=Pt(8), bold=False,
                          alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)

    doc.add_paragraph()

    # ========== ОСНОВНАЯ ТАБЛИЦА ==========
    main_table = doc.add_table(rows=6, cols=12)
    main_table.style = 'Table Grid'

    # Объединение первых 5 строк
    for row_idx in range(5):
        start_cell = main_table.cell(row_idx, 0)
        for col_idx in range(1, 12):
            start_cell.merge(main_table.cell(row_idx, col_idx))

    # Заполнение объединенных строк
    texts = [
        "Технологическая карта санитарной обработки № ________",
        "Помещение: _________________________",
        "",
        "Санитарная обработка помещений осуществляется после полного удаления из них сырьевых компонентов и материалов.",
        ""
    ]
    for row_idx, text in enumerate(texts):
        set_cell_font(main_table.cell(row_idx, 0), text, bold=True,
                      alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)

    # Заголовки колонок
    headers = [
        "Объект обработки", "Способ обработки", "№ инструкции",
        "Наименование средства", "Метод уборки",
        "Концентрация средства, %, метод разведения",
        "Температура раствора, °C", "Время выдержки средства, мин",
        "Инвентарь", "Периодичность", "Исполнитель",
        "Метод контроля/периодичность"
    ]
    for col, header in enumerate(headers):
        set_cell_font(main_table.cell(5, col), header, bold=True)

    # Ширина колонок основной таблицы
    main_widths = [3.0, 2.0, 1.5, 2.5, 2.0, 2.5, 1.8, 1.8, 1.8, 2.0, 2.0, 2.5]
    for col, width_cm in enumerate(main_widths):
        for row in range(len(main_table.rows)):
            set_cell_width(main_table.cell(row, col), width_cm)

    # Пустая строка для данных
    row = main_table.add_row()
    for col in range(12):
        set_cell_font(row.cells[col], "", bold=True)
        set_cell_width(row.cells[col], main_widths[col])

    # Сохраняем
    output_path = Path("tech_card_templates/шаблон.docx")
    output_path.parent.mkdir(exist_ok=True)
    doc.save(str(output_path))

    print(f"✅ Шаблон создан: {output_path}")
    print("   - Русский язык установлен для проверки орфографии")
    print("   - Таблица легенды: 10 колонок")
    print("   - Основная таблица: 12 колонок")


if __name__ == "__main__":
    import os

    old = Path("tech_card_templates/шаблон.docx")
    if old.exists():
        os.remove(old)
    create_template()