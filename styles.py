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

    def set_explicit_font(self, style, font_name):
        """Явно задаёт шрифт на уровне XML, чтобы переопределить шрифт темы"""
        rPr = style._element.get_or_add_rPr()  # Получаем или создаём элемент rPr (свойства текста)
        # Удаляем существующие элементы шрифта, чтобы избежать конфликтов
        for font_elem in rPr.findall(qn('w:rFonts')):
            rPr.remove(font_elem)
        # Создаём новый элемент w:rFonts
        font = OxmlElement('w:rFonts')
        font.set(qn('w:ascii'), font_name)  # Шрифт для ASCII-символов
        font.set(qn('w:hAnsi'), font_name)  # Шрифт для высоких ANSI-символов
        font.set(qn('w:cs'), font_name)  # Шрифт для сложных скриптов
        font.set(qn('w:eastAsia'), font_name)  # Шрифт для восточноазиатских языков
        rPr.append(font)

    def chenging_style(self, style_name, name, size, color, bold, italic, underline, caps, alignment, left_indent,
                       right_indent, space_before, space_after, line_spacing, first_line_indent, keep_with_next,
                       page_break_before, enable_widow_control):
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        style = self.doc.styles[style_name]
        font = style.font
        font.name = name
        font.size = Pt(size)
        font.color.rgb = RGBColor(*color)
        font.bold = bold
        font.italic = italic
        font.underline = underline
        font.all_caps = caps
        fmt = style.paragraph_format
        fmt.alignment = alignment_map[alignment]
        fmt.left_indent = Cm(left_indent)
        fmt.right_indent = Cm(right_indent)
        fmt.space_before = Mm(space_before)
        fmt.space_after = Mm(space_after)
        fmt.line_spacing = line_spacing
        fmt.first_line_indent = Cm(first_line_indent)
        fmt.keep_with_next = keep_with_next
        fmt.page_break_before = page_break_before
        if enable_widow_control:
            self.enable_widow_control(style)
        else:
            self.disable_widow_control(style)

        # Явно задаём шрифт на уровне XML, чтобы переопределить шрифт темы
        if style_name.startswith('Heading'):
            self.set_explicit_font(style, name)

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
                            False, False, False, 'center', 0,
                            0, 0, 0, 1, 0,
                            False, False, True)

    def image(self):
        try:
            self.doc.styles.add_style('Рисунок', WD_STYLE_TYPE.PARAGRAPH)
        except:
            pass
        self.chenging_style('Рисунок', 'Times New Roman', 12, (0, 0, 0), False,
                            False, False, False, 'center', 0, 0,
                            0, 0, 1, 0, True,
                            False, True)

    def list_bullet(self):
        style = self.doc.styles['List Bullet']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)
        style.font.color.rgb = RGBColor(0, 0, 0)
        fmt = style.paragraph_format
        fmt.left_indent = Cm(2.25)  # Отступ текста 2.25 см
        fmt.first_line_indent = Cm(1.25 - 2.25)  # Висячий отступ для маркера на 1.25 см
        fmt.space_after = Mm(0)
        fmt.line_spacing = 1.5
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        tab_stops = fmt.tab_stops
        tab_stops.add_tab_stop(Cm(1.25))  # Положение маркера 1.25 см
        self.set_explicit_font(style, 'Times New Roman')

    def list_number(self):
        style = self.doc.styles['List Number']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(14)
        style.font.color.rgb = RGBColor(0, 0, 0)
        fmt = style.paragraph_format
        fmt.left_indent = Cm(2.25)  # Отступ текста 2.25 см
        fmt.first_line_indent = Cm(1.25 - 2.25)  # Висячий отступ для маркера на 1.25 см
        fmt.space_after = Mm(0)
        fmt.line_spacing = 1.5
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        tab_stops = fmt.tab_stops
        tab_stops.add_tab_stop(Cm(1.25))  # Положение маркера 1.25 см
        self.set_explicit_font(style, 'Times New Roman')

    def run(self):
        # Настройка полей страницы
        section = self.doc.sections[0]
        section.left_margin = Mm(30)  # Левое поле 30 мм
        section.right_margin = Mm(10)  # Правое поле 10 мм
        section.top_margin = Mm(20)   # Верхнее поле 20 мм
        section.bottom_margin = Mm(20)  # Нижнее поле 20 мм

        # Вызов всех методов стилей
        self.image()
        self.text_of_table()
        self.caption_of_illustrations()
        self.normal()
        self.heading_1()
        self.heading_2()
        self.heading_3()
        self.caption_of_table()
        self.list_bullet()
        self.list_number()