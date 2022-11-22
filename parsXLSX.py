from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
from math import isnan
from openpyxl.styles import PatternFill
import time
from collections.abc import Iterable
from file_xlsx import FileXlsx
from os import path
from openpyxl import Workbook
from openpyxl.styles import Alignment

from file_xlsx import FileXlsx

xlsx_file = Path('./Silvery_Port.xlsx')
new_file_path = './test.xlsx'

COLOR_HEADER = 'FFC000'
COLOR_GRAY = 'BDBBB6'
COLOR_PINK = 'E5D1D0'
LIGHT_GREEN = 'B7DEB9'
DEEP_GREEN = '7BC77F'

# Retrieve cell value
listCabin = [57588279, 64130933, 56918501, 56249444, 61893335]
#listCabin = [54864947, 57965113, 61494944, 56662430, 60848967, 58061458]


def fill_row(sheet, row, col, color):
    for i in range(1, col):
        sheet.cell(row, i).fill = PatternFill(patternType='solid', fgColor=color)


""" Create header for new file
    and get the columns number """
def create_header(new_xlsx, names_list):  # TODO
    new_xlsx.set_column_header_count(len(names_list))
    column = new_xlsx.get_column_header_count()
    new_xlsx.row_filling(1, column, names_list)
    new_xlsx.color_row(1, 1, column, COLOR_HEADER)
    new_xlsx.wrap_row(1, 1, column + 1)
    new_xlsx.save()


def get_values(sheet, rw, colons):
    return [sheet.cell(row=rw, column=colon).value for colon in range(1, colons + 1)]


def fill_rows(sheet, row):  # main file
    column = 1
    while sheet.cell(row, column).value:
        sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_PINK)
        column += 1
    """ Add date on the last cell and fill background """
    sheet.cell(row, column).value = time.strftime("%x")
    sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_GRAY)


def get_current_weight_sum(path_to_file):
    """ From main file
        Read the values of the file in the dataframe
        Convert file in dict for keys and values """
    excel_data = pd.read_excel(path_to_file)
    exel_values = excel_data.to_dict('dict')
    cabin_num = []
    weight = excel_data.get("Фактический вес")
    for name, val in exel_values.items():
        if name == 'Номер вагона':
            for key, cabin in val.items():
                if cabin in listCabin:
                    cabin_num.append(key)

    return [weight[elem] for elem in cabin_num]


def get_data_from_main_file(path_to_main_file):
    """
        Get the values from the main table and colorized
        First row and column in header """
    values_list = []
    for row in range(2, rows):
        for column in range(1, columns):
            cell = sheet.cell(row, column)
            if cell.value in listCabin:
                values_list += [get_values(sheet, row, column + 1) + [time.strftime("%x")]]
                fill_rows(sheet, row)

    """ Safe changes in the main file """
    main_workbook.save(path_to_main_file)
    return values_list


def fill_new_xlsx(new_xlsx, values):  # Values from main file
    row, column = new_xlsx.rows_count(), new_xlsx.get_column_header_count()
    if row == 1:
        row += 1
    else:
        new_xlsx.clear_color_row(row, column)
    for value in values:
        new_xlsx.row_filling(row, column, value)
        row += 1
    new_xlsx.save()


def grand_total_handler(file_path, newxls):
    """ Sum of all weight """
    exel_data = pd.read_excel(file_path).to_dict('list')

    weight_list = [weight for weight in exel_data.get('Фактический вес')
                   if not isinstance(weight, Iterable) and not isnan(weight)]

    row, column = newxls.rows_count(), newxls.get_column_header_count()

    print(f"total row: {row}, column {column}")
    newxls.one_cell_filling(row + 1, column - 1, "Итого:")
    newxls.one_cell_filling(row + 1, column, sum(weight_list))
    newxls.color_row(row + 1, 1, column, COLOR_HEADER)


def addition_elements(newxls, weight_sum):  # TODO
    row, column = newxls.rows_count(), newxls.get_column_header_count()

    print(f"row {row}, column {column}")
    """ Count """
    newxls.one_cell_filling(row - 1, column + 1, "Кол-во:")
    newxls.color_one_cell(row - 1, column + 1, DEEP_GREEN)
    newxls.one_cell_filling(row - 1, column + 2, 5)
    """ Sum """
    new_xls.color_row(row, 1, column, LIGHT_GREEN, 'darkGrid')
    newxls.one_cell_filling(row, column + 1, "Сумма:")
    newxls.color_one_cell(row, column + 1, DEEP_GREEN)
    newxls.one_cell_filling(row, column + 2, weight_sum)


if __name__ == "__main__":
    """ Get values from main file """
    main_workbook = load_workbook(xlsx_file)  # path to the Excel file
    sheet = main_workbook.active
    rows = sheet.max_row
    columns = sheet.max_column
    header_names = [cell.value for cell in list(sheet.rows)[0] if cell.value] + ['Дата послупления']
    values_from_main_file = get_data_from_main_file(xlsx_file)
    current_weight_sum = sum(get_current_weight_sum(xlsx_file))

    """ Create new file if doesn't exist """
    new_xls = FileXlsx(new_file_path)

    create_header(new_xls, header_names)
    """ Adding new values """
    fill_new_xlsx(new_xls, values_from_main_file)
    """ Adding sum and count elements """
    """ Adding total sum """
    addition_elements(new_xls, current_weight_sum)
    grand_total_handler(new_file_path, new_xls)

    new_xls.save()
    new_xls.close()
