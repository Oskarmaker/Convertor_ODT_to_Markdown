import docx
from styles import Styles, qn, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT

path = 'C:/Users/vovat/Desktop/АКМС/Тест1.docx' # path to word file
doc = docx.Document(path)
styles = Styles(doc)
styles.run()
image_paragraphs_indexes = []
for i, p in enumerate(doc.paragraphs):
    found_image = False
    for run in p.runs:
        drawing_elements = run._element.findall(qn('w:drawing'))
        if drawing_elements:
            p.style = doc.styles['Рисунок']
            found_image = True
            image_paragraph_index = i
            break
    if found_image:
        next_paragraph = doc.paragraphs[image_paragraph_index + 1]
        next_paragraph.style = doc.styles['Подпись рисунка']
table_paragraphs_indexes = []
for i, block in enumerate(doc.element.body):
    if block.tag.endswith('tbl'):
        table_index = len(table_paragraphs_indexes)
        if table_index < len(doc.tables):
            table = doc.tables[table_index]

            # Применяем стиль таблицы и устанавливаем выравнивание
            table.style = 'Table Grid'
            table.alignment = WD_TABLE_ALIGNMENT.CENTER  # Явно задаём выравнивание по центру

            # Применяем стиль "Текст таблицы" к тексту в ячейках
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # paragraph.style = doc.styles['Текст таблицы']
                        # Дополнительно задаём шрифт для единообразия
                        for run in paragraph.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)

            # Находим предшествующий абзац (если он существует)
            if i > 0:
                prev_element = doc.element.body[i - 1]
                # prev_element.style = doc.styles['Подпись таблицы']
                if prev_element.tag.endswith('p'):
                    paragraph_index = sum(1 for j in range(i - 1) if doc.element.body[j].tag.endswith('p'))
                    if paragraph_index < len(doc.paragraphs):
                        prev_paragraph = doc.paragraphs[
                            paragraph_index]  # Исправлено: prev_paragraph вместо prev_paragraphs
                        prev_paragraph.style = doc.styles['Подпись таблицы']
                        table_paragraphs_indexes.append(paragraph_index)
                        print(f"Applied 'Подпись таблицы' to paragraph before table: {prev_paragraph.text}...")
# print(doc.styles['Подпись таблицы'], list(doc.styles))
doc.save(path)