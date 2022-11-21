import openpyxl
import pandas as pd
from pathlib import Path
from math import isnan
from openpyxl.styles import PatternFill
import time
from collections.abc import Iterable
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


def fill_row(sheet, row, col, color):
    for i in range(1, col):
        sheet.cell(row, i).fill = PatternFill(patternType='solid', fgColor=color)


def create_xlsx_file_header(file_path, names_list):  # new file header
    count = len(names_list)
    new_xlsx = FileXlsx(file_path)
    new_xlsx.row_filling(1, count, names_list)
    new_xlsx.color_row(1, 1, count, COLOR_HEADER)
    new_xlsx.wrap_row(1, 1, count + 1)
    return new_xlsx


def get_values(sheet, rw, colons):
    return [sheet.cell(row=rw, column=colon).value for colon in range(1, colons + 1)]


def fill_rows(sheet, row):  # main file
    column = 1
    while sheet.cell(row, column).value:
        sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_PINK)
        column += 1
    """Add date on the last cell and fill background"""
    sheet.cell(row, column).value = time.strftime("%x")
    sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_GRAY)


def get_weight_sum(path_to_file):  # From main file
    excel_data = pd.read_excel(path_to_file)  # Read the values of the file in the dataframe
    exel_values = excel_data.to_dict('dict')  # Convert file in dict for keys and values
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
        First row and column in header
    """
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


def fill_new_xlsx(new_xlsx, values, weight_sum):  # Values from main file

    row, column = 2, len(values[0])  # Because we already have a header
    for value in values:
        new_xlsx.row_filling(row, len(value), value)
        row += 1


def grand_total_handler(file_path, newxls, column):
    """ Sum of all weight """
    exel_data = pd.read_excel(file_path).to_dict('list')

    weight_list = [weight for weight in exel_data.get('Фактический вес')
                   if not isinstance(weight, Iterable) and not isnan(weight)]
    total_weight = sum(weight_list)
    row = len(weight_list) + 1  # plus header

    newxls.color_row(row, 1, column, LIGHT_GREEN)
    newxls.one_cell_filling(row + 1, column - 1, "Итого:")
    newxls.one_cell_filling(row + 1, column, total_weight)
    newxls.color_row(row + 1, 1, column, COLOR_HEADER)


# def addition_elements():
#     row -= 1
#     """ Sum """
#     new_xlsx.one_cell_filling(row, column + 1, "Сумма:")
#     new_xlsx.color_one_cell(row, column + 1, DEEP_GREEN)
#     new_xlsx.one_cell_filling(row, column + 2, weight_sum)
#     """ Count """
#     new_xlsx.one_cell_filling(row + 1, column + 1, "Кол-во:")
#     new_xlsx.color_one_cell(row + 1, column + 1, DEEP_GREEN)
#     new_xlsx.one_cell_filling(row + 1, column + 2, 5)


if __name__ == "__main__":
    """ Get values from main file """
    main_workbook = openpyxl.load_workbook(xlsx_file)  # path to the Excel file
    sheet = main_workbook.active
    rows = sheet.max_row
    columns = sheet.max_column
    header_names = [cell.value for cell in list(sheet.rows)[0] if cell.value] + ['Дата послупления']
    values_from_main_file = get_data_from_main_file(xlsx_file)
    current_weight_sum = sum(get_weight_sum(xlsx_file))

    """ Create new file if doesn't exist """
    new_xls = create_xlsx_file_header(new_file_path, header_names)
    """ Adding new values """
    fill_new_xlsx(new_xls, values_from_main_file, current_weight_sum)
    """ Adding total sum """
    grand_total_handler(new_file_path, new_xls, len(header_names))
    new_xls.save()
