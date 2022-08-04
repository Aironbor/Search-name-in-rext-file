# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os
import re
from datetime import datetime

import xlsxwriter

class FindMaterial:
    def print_hi(self, file):
        # Use a breakpoint in the code line below to debug your script.
        profile_list = []
        self.dict = {}
        self.list = []
        with open(file) as f:
            datatext = f.readlines()
            for line in datatext:
                if "PROFILE_NAME" in line:
                    progile_name = re.search('"(.+?)"', line).group(1)
                    profile_list.append(progile_name)
                    if 'I' in progile_name:
                        if 'ДК' in progile_name:
                            self.dict[progile_name] = 'Двутавр Дополнительный колонный'
                            self.list.append(f'Двутавр дополнительный колонный {progile_name}')
                        elif 'ДБ' in progile_name:
                            self.dict[progile_name] = 'Двутавр Дополнительный балочный'
                            self.list.append(f'Двутавр дополнительный балочный {progile_name}')
                        elif 'С' in progile_name:
                            self.dict[progile_name] = 'Двутавр Свайный'
                            self.list.append(f'Двутавр свайный {progile_name}')
                        elif 'K' in progile_name:
                            self.dict[progile_name] = 'Двутавр Колонный'
                            self.list.append(f'Двутавр колонный {progile_name}')
                        elif 'Ш' in progile_name:
                            self.dict[progile_name] = 'Двутавр Широкополочный'
                            self.list.append(f'Двутавр широкополочный {progile_name}')
                        elif 'Б' in progile_name:
                            self.dict[progile_name] = 'Двутавр Нормальный'
                            self.list.append(f'Двутавр нормальный {progile_name}')
                        elif 'М' in progile_name:
                            self.dict[progile_name] = 'Двутавр Специальный горячекатаный прокат'
                            self.list.append(f'Двутавр специальный горячекатаный прокат {progile_name}')
                    elif '[' in progile_name:
                        if 'П' in progile_name:
                            self.dict[progile_name] = 'Швеллер с параллельными полками(П)'
                            self.list.append(f'Швеллер с параллельными полками(П) {progile_name}')
                        elif 'Гн' in progile_name:
                            self.dict[progile_name] = 'Швеллер гнутый'
                            self.list.append(f'Швеллер гнутый {progile_name}')
                        else:
                            self.dict[progile_name] = 'Швеллер с уклоном полок(У)'
                            self.list.append(f'Швеллер с уклоном полок(У) {progile_name}')
                    elif 'L' in progile_name:
                        self.dict[progile_name] = 'Уголок'
                        self.list.append(f'Уголок {progile_name}')

                    elif 'TЭ' in progile_name:
                        self.dict[progile_name] = 'Труба электросварная  ГОСТ 10704-91'
                        self.list.append(f'Труба электросварная  ГОСТ 10704-91 {progile_name}')

                    elif 'ДУ' in progile_name:
                        self.dict[progile_name] = 'Труба круглая  ГОСТ 3262-72'
                        self.list.append(f'Труба электросварная  ГОСТ 3262-72 {progile_name}')

                    elif 'TK' in progile_name:
                        self.dict[progile_name] = 'Труба круглая  ГОСТ 8732-78'
                        self.list.append(f'Труба электросварная  ГОСТ 8732-78 {progile_name}')

                    elif 'Гнз' in progile_name:
                        self.dict[progile_name] = 'Труба прямоугольная'
                        self.list.append(f'Труба прямоугольная {progile_name}')

                    elif '—' in progile_name:
                        self.dict[progile_name] = 'Лист'
                        self.list.append(f'Лист {progile_name}')

                    else:
                        self.dict[progile_name] = ''
                        self.list.append(f'')

    def back_dict(self, file):
        self.print_hi(file)
        return self.dict

    def back_list(self, file):
        self.print_hi(file)
        return self.list


def report_spec_by_material(dictonary, list):
    material_dict = dictonary
    #  ОТЧЕТ ДЛЯ КМД
    workbook = xlsxwriter.Workbook(f'Спецификация материала из Теклы.xlsx')
    # Форматы format()
    name_format = workbook.add_format(
        {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    name_format_main = workbook.add_format(
        {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    special_numb = workbook.add_format(
        {'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
    # Форматы для объединнеых ячеек
    name_merge_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '#0',
    })
    workbook.set_properties({
        'title': f'Выборка металла',
        'subject': 'With document properties',
        'author': 'Ivan Metliaev',
        'manager': '',
        'company': 'ООО Тентовые конструкции',
        'category': 'Цех Металлоконструкций',
        'keywords': 'Раскрой металла, Металл, Материалы',
        'created': datetime.today(),
        'comments': 'Created with Python and Ivan Metliaev program'})
    # Создаваемые листы
    # Наименование листа
    worksheet_0 = workbook.add_worksheet(f'Выборка металла')
    # Размерность колонок
    worksheet_0.set_column(0, 0, 14)
    worksheet_0.set_column(1, 1, 20)
    worksheet_0.set_column(2, 2, 20)
    # Записи
    worksheet_0.merge_range(0, 0, 0, 1, f'Материал из Теклы ',name_merge_format)
    column_name = ['№', 'Cечение', 'Название материала', 'Объединенное']
    # Заголовки первой таблицы
    curnt_numb_row = 3
    num = 0
    worksheet_0.write_row(2, 0, column_name, name_format)
    for material in material_dict:
        # №
        num += 1
        worksheet_0.write(curnt_numb_row, 0, num, special_numb)
        # Материал
        worksheet_0.write(curnt_numb_row, 1, material, name_format_main)
        worksheet_0.write(curnt_numb_row, 2, material_dict[material], name_format_main)
        curnt_numb_row += 1

    curnt_numb_row = 3
    for el in list:
        worksheet_0.write(curnt_numb_row, 3, el, name_format_main)
        curnt_numb_row += 1

    workbook.close()
    os.startfile(f'Спецификация материала из Теклы.xlsx')

    # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.


if __name__ == '__main__':
    file_name = "profiles_red.txt"
    find = FindMaterial()
    report_spec_by_material(find.back_dict(file_name), find.back_list(file_name))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
