"""
1. Импорт библиотек:
        docx - для работы с word;
        openpyxl - для работы с excel;
        os - для перемещения по файловой системе
        re - для вытаскивания серийных номеров с помощью регулярных выражений
2. Определиться откуда запускаем скрипт
3. Рекурсивно проходим по всем папкам
4. Открываем файлы с расширением .docx - открываем и проверяем наличие серийного номера и номера свидетельства
5. Записываем это в отдельный словарь
6. Отдельно читаем файл со списком РЭС, вытаскиваем все значения серийника и номера свидетельства
7. Записываем это в отдельный словарь
8. Сверяем файл из word и excel и записываем результат в файл excel в новый столбец с инвентаризацией
"""
import os
import re

import docx
import openpyxl

start_path = os.getcwd()  #
pathOfExcel = 'C:\\Users\\'
nameOfSheetExcel = ''   # Имя листа в Excel
nameOfColumnExcel = ''

nameOfColumnExcel = 'Название 3'  # Известное имя столбца из учетного файла в Excel
serialNumberRegex = re.compile()  # Объект Regex для серийного номера
certificateNumberRegex = re.compile()  # Объект Regex для номера свидетельства о регистрации РЭС

dictOfWord = {}  # Словарь с
dictOfExcel = {}  # Словарь с

# Рекурсивный проход по всем папкам в текущей директории
for current_dir, dirs, files in os.walk(start_path):
    for file in files:

        # Выбираем все файлы Word
        if file.endswith('.docx'):
            documentWord = docx.Document(file)
            table = documentWord.tables[0]  # Читаем первую таблицу (она там одна)

            # Проходимся по всем строкам и заходим в каждую ячейку
            for row in table.rows:
                for cell in row.cells:

                    # Добавляем в словарь dictOfWord совпадение серийного номера и номера свидетельства
                    dictOfWord[serialNumberRegex.search(cell.text).group()] = serialNumberRegex.search(cell.text).group()

# Открываем файл в директории
documentExcel = openpyxl.load_workbook(pathOfExcel)
sheet1 = documentExcel[nameOfSheetExcel]
for row in sheet1.rows:
    print(str(row[1].value))  # 1 - индекс столбца (начиная с 0)
