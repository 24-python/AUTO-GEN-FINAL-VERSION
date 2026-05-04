import zipfile
from lxml import etree

docx_path = 'производственных помещений.docx'

with zipfile.ZipFile(docx_path, 'r') as z:
    with z.open('word/document.xml') as f:
        root = etree.fromstring(f.read())

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
tables = root.findall('.//w:tbl', ns)
print(f'Таблиц: {len(tables)}')

for t_idx, table in enumerate(tables):
    rows = table.findall('.//w:tr', ns)
    print(f'\nТаблица {t_idx + 1} - {len(rows)} строк')

    for r_idx, row in enumerate(rows[:8]):
        cells = row.findall('.//w:tc', ns)
        cell_texts = []
        for cell in cells:
            texts = cell.findall('.//w:t', ns)
            text = ' '.join(t.text or '' for t in texts).strip()[:80]
            cell_texts.append(text)
        print(f'  Строка {r_idx}: {len(cells)} ячеек | ' + ' | '.join(cell_texts[:5]))