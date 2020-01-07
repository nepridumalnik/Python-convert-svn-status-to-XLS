import xlwt  # Импорт модуля для создания xls файлов


class XLSdoc:
    def __init__(self, sheetname):
        self.sheet = sheetname  # Создание объекта для работы со страницей excel
        self.font = xlwt.Font()  # Создание объекта для указания шрифта
        self.font.name = "Times New Roman"  # Задание шрифта
        self.font.height = 280  # Установка размера шрифта (14х20)

        self.style = xlwt.XFStyle()  # Создание объекта стиля
        self.style.font = self.font  # Применения шрифта к стилю

        self.doc = xlwt.Workbook()  # Создание объекта для записи в книгу excel
        self.sheet = self.doc.add_sheet(sheetname)  # Добавление листа в книгу

    def write(self, row, column, data):  # Метод внесения записи в лист
        self.sheet.write(row, column, data, self.style)  # Запись в лист - строка, столбец, данные, стиль

    def savedoc(self, docname):  # Сохранение страницы под имененем docname
        self.doc.save(docname)
