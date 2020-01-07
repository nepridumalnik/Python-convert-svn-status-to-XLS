import tkinter as tk  # Импорт tkinter для GUI
import tkinter.filedialog as fd  # Импорт модуля для создания окон выбора и сохранения файла
from _xlwt_ import *  # Импорт модуля класса работы с excel


def linen(line):  # Перевод строки в список
    newline = []  # Создание пустого списка

    if line[0] == ' ':  # Приведение краткого символов к понятному виду
        newline.append('Нет изменений')
    elif line[0] == 'A':
        newline.append('Добавление')
    elif line[0] == 'D':
        newline.append('Удаление')
    elif line[0] == 'M':
        newline.append('Изменение')
    elif line[0] == 'R':
        newline.append('Перемещение')
    elif line[0] == 'C':
        newline.append('Конфликт содержимого')
    elif line[0] == 'X':
        newline.append('Внешнее определение предмета')
    elif line[0] == 'I':
        newline.append('Игнорируется')
    elif line[0] == '!':
        newline.append('Отсутствует')
    elif line[0] == '~':
        newline.append('Неправильный тип объекта')

    if line[0] == 'A' and line[3] == '+':
        newline.append('История запланирована')
    else:
        newline.append('Нет запланированной истории')

    if line[6] == 'L':
        newline.append('Заблокировано')
    else:
        newline.append('Незаблокировано')

    if line[9] == ' ':
        newline.append("Дочерний")
    elif line[9] == 'S':
        newline.append('Переключен')

    l = line[13:]  # Создание вспомогательного списка

    while '  ' in l:
        l = l.replace('  ', ' ')  # Удаление всех двойных пробелов

    l = l.split(' ')  # Перевод строки в список

    if l[0] == '':
        l = l[1:]

    if len(l) < 4:
        for i in range(4 - len(l)):
            newline.append(' ')

    for e in l:
        newline.append(e)

    return newline


def readFromFile(filename):  # Чтение из файла svn status
    file = open(filename, "r")  # Открытие файла для чтения

    Lines = []  # Список списков со строками

    for line in file:
        if line != "      >   local edit, incoming delete upon merge\n":  # Удаление строк, не несущих смысловую нагрузку
            Lines.append(linen(line))  # Добавление строк, приведённых в вид исправленных списков

    file.close()  # Завершение работы с файлом
    return Lines  # Возвращение списка со списками


def writeIntoXLS(filename, Lines):  # Запись в книгу excel именем filename списка Lines
    xls = XLSdoc("Status")  # Имя листа "Status"

    xls.write(0, 0, 'Свойство')  # Установка названий столбцаи
    xls.write(0, 1, 'Статус')
    xls.write(0, 2, 'Блокировка')
    xls.write(0, 3, 'Отношения')
    xls.write(0, 4, '№1')
    xls.write(0, 5, '№2')
    xls.write(0, 6, 'Пользователь')
    xls.write(0, 7, 'Имя каталога/файла')

    rownum = 0  # Номер строки
    for line in Lines:
        colnum = 0  # Номер столбца
        for e in line:
            xls.write(rownum + 1, colnum, e)  # Запись в строку, столбец элемента e из списка line
            colnum += 1  # Инкрементация номера столбца
        rownum += 1  # Инкрементация номера строки

    xls.savedoc(filename)  # Сохранение книги


def getFileName(path):
    return path.split("/")[-1]  # Получение названия файла из полного пути


def findFile(self):  # Найти файл для чтения
    global path, Lines
    try:
        path = fd.askopenfilename()  # Полный путь для файла - получается в окне проводника
        if path != '':
            Lines = readFromFile(path)  # Чтение содержимого файла и запись в список
            global l1
            l1["text"] = "Файл: \"" + getFileName(path) + "\""  # Запись в лейбл имени прочтённого файла
    except:
        l1["text"] = "Неверный формат файла"  # Вывод ошибки в случае невозможности прочитать файл


def saveFile(self):  # Запись результата чтения в файл
   try:
       global saveas, Lines, path
       saveas = fd.asksaveasfilename(filetypes=(("XLS files", "*.XLS"), ("All files", "*.*")))  # Открытие окна проводника для выбора пути записи файла
       if saveas != '' and path != '':
           writeIntoXLS(saveas, Lines)  # Запись файла в формате книги excel
           l1["text"] = "Файл сохранён"  # Запись в лейбл сообщения об успешной записи файла
           path = ''  # Очистка пути файла для предотвращения случайной перезаписи
           saveas = ''  # Очистка пути сохранения файла для предотвращения случайной перезаписи
       else:
           l1["text"] = "Не указан файл"  # Вывод ошибки в случае, если не был указан файл для записи
   except:
       l1["text"] = "Похоже файл уже открыт"  # Вывод ошибки в случае если нет доступа к записи в файл с этим именем


Lines = []  # Список для содержимого файла
path = ''  # Путь к файлу
saveas = ''  # Путь к записываемому файлу

root = tk.Tk()  # Создание окна tkinter
root.title("GUI to subject")  # Указание имени окна
root.geometry("250x50+300+300")  # Указание размера окна и его местоположения
root.resizable(False, False)  # Запрет на изменение размеров по двум осям

l1 = tk.Label(text="Файл не выбран", width="40")  # Создание лейбле, указание его надписи и ширина
l1.pack(side=tk.TOP)  # Способ расположения в окне - сверху

frame = tk.Frame()  # Создание фрейма для кнопок

frame.pack(fill=tk.X, side=tk.TOP)  # Расположение фрейма
btn1 = tk.Button(frame, text="Выбрать файл")  # Создание кнопки и надписи
btn1.pack(side=tk.LEFT, padx=1, pady=5, expand=True)  # Расположение кнопки, разрешение на заполнение доступного пространства

btn2 = tk.Button(frame, text="Расшифровать")
btn2.pack(side=tk.LEFT, padx=1, pady=5, expand=True)

btn1.bind('<Button-1>', findFile)  # Привязка функции к нажатию на кнопку левой кнопкой мышки
btn2.bind('<Button-1>', saveFile)

root.mainloop()  # Вывод окна на экран
