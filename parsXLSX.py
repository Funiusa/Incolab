import openpyxl
import pandas as pd
from pathlib import Path
from math import isnan
from openpyxl.styles import PatternFill
import time
from openpyxl import Workbook
from openpyxl.styles import Alignment

from file_xlsx import FileXlsx

xlsx_file = Path('./Silvery_Port.xlsx')
new_file_path = './test.xlsx'

COLOR_HEADER = 'FFC000'
COLOR_GRAY = 'BDBBB6'
COLOR_PINK = 'E5D1D0'
LIGHT_GREEN = 'F0F6E8'

# Retrieve cell value
listCabin = [57588279, 64130933, 56918501, 56249444, 61893335]


def fill_row(sheet, row, col, color):
    for i in range(1, col):
        sheet.cell(row, i).fill = PatternFill(patternType='solid', fgColor=color)


def create_xlsx_file_header(path, filename, names_list):
    count = len(names_list)
    newxl = FileXlsx(path, filename)
    newxl.row_filling(1, count, names_list)
    newxl.color_row(1, 1, count, COLOR_HEADER)
    newxl.wrap_row(1, 1, count + 1)
    return newxl


def get_values(sheet, rw, colons):
    return [sheet.cell(row=rw, column=colon).value for colon in range(1, colons + 1)]


def fill_rows(sheet, row):
    column = 1
    while sheet.cell(row, column).value:
        sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_PINK)
        column += 1
    """Add date on the last cell and fill background"""
    sheet.cell(row, column).value = time.strftime("%x")
    sheet.cell(row, column).fill = PatternFill(patternType='solid', fgColor=COLOR_GRAY)


def add_new_elements():
    pass


def get_weight_sum(path_to_file):
    # Read the values of the file in the dataframe
    excel_data = pd.read_excel(path_to_file)
    # Convert file in dict for keys and values
    exel_values = excel_data.to_dict('dict')
    cabin_num = []
    weight = excel_data.get("Фактический вес")

    for name, val in exel_values.items():
        if name == 'Номер вагона':
            for key, cabin in val.items():
                if cabin in listCabin:
                    cabin_num.append(key)

    return [weight[elem] for elem in cabin_num]


def grand_total_handler(file_path, newxls, column):
    """ Sum of all weight """
    exel_data = pd.read_excel(file_path).to_dict('list')

    weight_list = [weight for weight in exel_data.get('Фактический вес') if type(weight) is float and not isnan(weight)]
    total_weight = sum(weight_list)
    row = len(weight_list) + 1  # plus header

    newxls.color_row(row, 1, column, LIGHT_GREEN)
    newxls.one_cell_filling(row + 1, column - 1, "Итого:")
    newxls.one_cell_filling(row + 1, column, total_weight)
    newxls.color_row(row + 1, 1, column, COLOR_HEADER)


def get_dataFromMainXSLX(path_to_file):
    main_workbook = openpyxl.load_workbook(path_to_file)  # path to the Excel file
    sheet = main_workbook.active
    rows = sheet.max_row
    columns = sheet.max_column
    header_names = [cell.value for cell in list(sheet.rows)[0] if cell.value]
    header_names.append('Дата послупления')

    """ Checking the data """
    weight_sum = sum(get_weight_sum(path_to_file))
    print(weight_sum)

    """
        Get the values from the main table
        First row and column is header
    """
    values_list = []
    for row in range(2, rows):
        for column in range(1, columns):
            cell = sheet.cell(row, column)
            if cell.value in listCabin:
                values_list += [get_values(sheet, row, column + 1) + [time.strftime("%x")]]
                fill_rows(sheet, row)

    """ Safe changes in the main file """
    last_column = len(values_list[0])
    count_rows = len(values_list)
    main_workbook.save(path_to_file)

    newxls = create_xlsx_file_header('./', 'test.xlsx', header_names)  # Change file path
    row = 2  # Because we already have a header

    for value in values_list:
        newxls.row_filling(row, len(value), value)
        row += 1
    column = len(values_list[0])

    newxls.one_cell_filling(last_column, last_column + 1, "Сумма:")
    newxls.color_one_cell(last_column, last_column + 1, COLOR_HEADER)
    newxls.one_cell_filling(last_column, last_column + 2, weight_sum)

    # newxls.one_cell_filling(last_column + 1, last_column + 1, "Кол-во:")
    # newxls.color_one_cell(last_column + 1, last_column + 1, COLOR_HEADER)
    # newxls.one_cell_filling(last_column + 1, last_column + 2, count_rows)

    grand_total_handler(new_file_path, newxls, column)

    newxls.save()


if __name__ == "__main__":
    # take_values()
    get_dataFromMainXSLX(xlsx_file)
