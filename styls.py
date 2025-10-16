import docx
from docx.shared import Pt, RGBColor, Inches, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.text import WD_UNDERLINE
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn  # Для работы с пространствами имён XML


class Styles:
    def __init__(self, doc: docx.document.Document):
        self.doc = doc

    def chenging_style(self, style_name, name, size, color, bold, italic, underline, caps, alignment, left_indent,
                       right_indent, space_before, space_after, line_spacing, first_line_indent, keep_with_next,
                       page_break_before, enable_widow_control):
        # Карта выравнивания
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        style = self.doc.styles[style_name]  # Обращение к стилю
        font = style.font  # Обращение к шрифту для стиля
        font.name = name  # Изменение имени шрифта
        font.size = Pt(size)  # Изменение размера шрифта
        font.color.rgb = RGBColor(*color)  # Изменение цвета шрифта
        font.bold = bold  # Полужирный
        font.italic = italic  # Курсив
        font.underline = underline  # Подчёркивание
        font.all_caps = caps  # ВСЕ ПРОПИСНЫЕ
        fmt = style.paragraph_format  # Обращение к абзацу стиля
        fmt.alignment = alignment_map[alignment]  # Выравнивание
        fmt.left_indent = Mm(left_indent)  # Левый отступ
        fmt.right_indent = Cm(right_indent)  # Правый отступ
        fmt.space_before = Mm(space_before)  # Интервал перед
        fmt.space_after = Mm(space_after)  # Интервал после
        fmt.line_spacing = line_spacing  # Межстрочный интервал
        fmt.first_line_indent = Cm(first_line_indent)  # Отступ для первой строки
        fmt.keep_with_next = keep_with_next  # Не отрывать от следующего
        fmt.page_break_before = page_break_before  # С новой страницы
        if enable_widow_control:
            self.enable_widow_control(style)  # Запрет висячих строк
        else:
            self.disable_widow_control(style)  # Отключение запрета висячих строк

    def enable_widow_control(self, style):
        """Включает 'Запрет висячих строк' (widow/orphan control) для стиля"""
        pPr = style._element.get_or_add_pPr()  # Получаем или создаём элемент pPr для стиля
        # Удаляем существующий элемент w:widowControl, если он есть
        for elem in pPr.findall(qn('w:widowControl')):
            pPr.remove(elem)
        # Добавляем новый элемент w:widowControl с включённым состоянием
        widow = OxmlElement('w:widowControl')  # Создаём элемент
        pPr.append(widow)  # Добавляем в pPr

    def disable_widow_control(self, style):
        """Отключает 'Запрет висячих строк' (widow/orphan control) для стиля"""
        pPr = style._element.get_or_add_pPr()  # Получаем или создаём элемент pPr для стиля
        # Удаляем существующий элемент w:widowControl, если он есть
        for elem in pPr.findall(qn('w:widowControl')):
            pPr.remove(elem)
        # Добавляем w:widowControl с атрибутом w:val="0" для отключения
        widow = OxmlElement('w:widowControl')
        widow.set(qn('w:val'), '0')  # Отключаем контроль висячих строк
        pPr.append(widow)

    def heading_1(self):
        self.chenging_style('Heading 1', 'Times New Roman', 18, (0, 0, 0), True,
                            False, False, True, 'left', 1.25,
                            0, 0, 10, 1.5, 0,
                            True, True, False)

    def heading_2(self):
        self.chenging_style('Heading 2', 'Times New Roman', 16, (0, 0, 0), True,
                            False, False, False, 'left', 1.25,
                            0, 15, 10, 1.5, 0,
                            True, False, True)

    def heading_3(self):
        self.chenging_style('Heading 3', 'Times New Roman', 14, (0, 0, 0), True,
                            False, False, False, 'left', 1.25,
                            0, 15, 10, 1.5, 0,
                            True, False, True)

    def normal(self):
        self.chenging_style('Normal', 'Times New Roman', 14, (0, 0, 0), False,
                            False, False, False, 'justify', 0,
                            0, 0, 0, 1.5, 1.25,
                            False, False, True)

    def caption_of_illustrations(self):
        try:
            self.doc.styles.add_style('Подпись рисунка', WD_STYLE_TYPE.PARAGRAPH)
        except:
            pass
        self.chenging_style('Подпись рисунка', 'Times New Roman', 12, (0, 0, 0), True,
                            False, False, False, 'center', 0,
                            0, 0, 6, 1, 0,
                            False, False, True)

    def caption_of_table(self):
        try:
            self.doc.styles.add_style('Подпись таблицы', WD_STYLE_TYPE.PARAGRAPH)
        except:
            pass
        self.chenging_style('Подпись таблицы', 'Times New Roman', 12, (0, 0, 0), False,
                            True, False, False, 'left', 0,
                            0, 6, 0, 1, 0,
                            False, False, True)

    def text_of_table(self):
        try:
            self.doc.styles.add_style('Текст таблицы', WD_STYLE_TYPE.PARAGRAPH)
        except:
            pass
        self.chenging_style('Текст таблицы', 'Times New Roman', 12, (0, 0, 0), False,
                            True, False, False, 'center', 0,
                            0, 0, 0, 1, 0,
                            False, False, True)

    def image(self):
        try:
            self.doc.styles.add_style('Изображение', WD_STYLE_TYPE.PARAGRAPH)
        except:
            pass
        self.chenging_style('Изображение', 'Times New Roman', 12, (0, 0, 0), False,
                            False, False, False, 'center', 0, 0,
                            0, 0, 1, 0, True,
                            False, True)

    def run(self):
        self.image()
        self.text_of_table()
        self.caption_of_illustrations()
        self.normal()
        self.heading_1()
        self.heading_2()
        self.heading_3()
        self.caption_of_table()