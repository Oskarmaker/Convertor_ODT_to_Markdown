import docx
from docx.shared import Pt, RGBColor, Inches, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.text import WD_UNDERLINE
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement


class Styles:
    def __init__(self, doc: docx.document.Document):
        self.doc = doc

    def enable_widow_control(self, paragraph):
        """Включает 'Запрет висячих строк' (widow/orphan control) для абзаца"""
        pPr = paragraph._p.get_or_add_pPr()
        widow = OxmlElement('w:widowControl')
        widow.set('w:val', '1')
        pPr.append(widow)


    def heading_1(self):
        style = self.doc.styles['Heading 1'] # Обращение к стилю Заголовок 1
        font = style.font # Обращение к шрифту для стиля Заголовок 1
        font.name = 'Times New Roman' # Изменение имени шрифта
        font.size = Pt(18) # Изменение размера шрифта
        font.color.rgb = RGBColor(0, 0, 0) # Изменение цвета шрифта
        font.bold = True # Полужирный
        font.caps = True # ВСЕ ПРОПИСНЫЕ
        fmt = style.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT # По левому краю
        fmt.left_indent = Mm(1.25) # Левый отступ
        fmt.right_indent = Cm(0) # Правый отступ
        fmt.space_before = Mm(0) # Интервал перед
        fmt.space_after = Mm(10) # Интервал после
        fmt.line_spacing = 1.5 # Межстрочный интервал
        fmt.keep_with_next = True # Не отрывать от следующего
        fmt.page_break_before = True # С новой строки

    def heading_2(self):
        style = self.doc.styles['Heading 2']  # Обращение к стилю Заголовок 1
        font = style.font  # Обращение к шрифту для стиля Заголовок 1
        font.name = 'Times New Roman'  # Изменение имени шрифта
        font.size = Pt(16)  # Изменение размера шрифта
        font.color.rgb = RGBColor(0, 0, 0)  # Изменение цвета шрифта
        font.bold = True  # Полужирный
        fmt = style.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT  # По левому краю
        fmt.left_indent = Mm(1.25)  # Левый отступ
        fmt.right_indent = Cm(0)  # Правый отступ
        fmt.space_before = Mm(15)  # Интервал перед
        fmt.space_after = Mm(10)  # Интервал после
        fmt.line_spacing = 1.5  # Межстрочный интервал
        fmt.keep_with_next = True  # Не отрывать от следующего
        self.enable_widow_control(style) # Запрет висячих строк

    def heading_3(self):
        style = self.doc.styles['Heading 3']  # Обращение к стилю Заголовок 1
        font = style.font  # Обращение к шрифту для стиля Заголовок 1
        font.name = 'Times New Roman'  # Изменение имени шрифта
        font.size = Pt(14)  # Изменение размера шрифта
        font.color.rgb = RGBColor(0, 0, 0)  # Изменение цвета шрифта
        font.bold = True  # Полужирный
        fmt = style.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT  # По левому краю
        fmt.left_indent = Mm(1.25)  # Левый отступ
        fmt.right_indent = Cm(0)  # Правый отступ
        fmt.space_before = Mm(15)  # Интервал перед
        fmt.space_after = Mm(10)  # Интервал после
        fmt.line_spacing = 1.5  # Межстрочный интервал
        fmt.keep_with_next = True  # Не отрывать от следующего
        self.enable_widow_control(style) # Запрет висячих строк

    def normal(self):
        style = self.doc.styles['Normal']  # Обращение к стилю Заголовок 1
        font = style.font  # Обращение к шрифту для стиля Заголовок 1
        font.name = 'Times New Roman'  # Изменение имени шрифта
        font.size = Pt(14)  # Изменение размера шрифта
        font.color.rgb = RGBColor(0, 0, 0)  # Изменение цвета шрифта
        font.bold = False
        fmt = style.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT  # По левому краю
        fmt.left_indent = Mm(1.25)  # Левый отступ
        fmt.right_indent = Cm(0)  # Правый отступ
        fmt.space_before = Mm(15)  # Интервал перед
        fmt.space_after = Mm(10)  # Интервал после
        fmt.line_spacing = 1.5  # Межстрочный интервал
        fmt.keep_with_next = True  # Не отрывать от следующего
        self.enable_widow_control(style)  # Запрет висячих строк