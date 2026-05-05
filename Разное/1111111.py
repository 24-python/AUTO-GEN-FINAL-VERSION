import zipfile
from lxml import etree

docx_path = '../используемые_средства_новая_таблица.docx'

with zipfile.ZipFile(docx_path, 'r') as z:
    with z.open('word/document.xml') as f:
        root = etree.fromstring(f.read())

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Находим все таблицы
tables = root.findall('.//w:tbl', ns)
print(f'Найдено таблиц: {len(tables)}\n')

for t_idx, table in enumerate(tables):
    rows = table.findall('.//w:tr', ns)
    print(f'=== Таблица {t_idx + 1} ({len(rows)} строк) ===')

    for r_idx, row in enumerate(rows[:4]):  # первые 4 строки
        cells = row.findall('.//w:tc', ns)
        print(f'  Строка {r_idx + 1} ({len(cells)} ячеек):')

        for c_idx, cell in enumerate(cells[:15]):  # первые 15 ячеек
            texts = cell.findall('.//w:t', ns)
            text = ' '.join(t.text or '' for t in texts).strip()[:60]

            # Проверяем чек-боксы
            sdt = cell.findall('.//w:sdt', ns)
            has_checkbox = len(sdt) > 0

            marker = '☒/☐' if has_checkbox else '   '
            if text:
                print(f'    [{c_idx}] {marker} {text}')

        if len(cells) > 15:
            print(f'    ... ещё {len(cells) - 15} ячеек')
    print()