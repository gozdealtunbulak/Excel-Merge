import openpyxl
import xlrd
from openpyxl import Workbook, load_workbook
import pandas as pd

name = input('Name of the excel file: ')
extension = input('The extension: ')
path = input('The location of the file: ')
f_name = name + '.' + extension


def check_list(s_num, num):
    """
    :param s_num: the sheet index
    :param num: the row number
    :return: this function returns a list of the desired row of the chosen sheet
    """
    wb = xlrd.open_workbook(r'{}\{}'.format(path, f_name))
    sheet = wb.sheet_by_index(s_num)
    list_check = []
    row_check = sheet.row_values(num)
    list_check.extend(row_check)
    return list_check


def add_new_sheet(sheet_name, index):
    """
    :param index: index of the sheet
    :param sheet_name: name of the sheet to be created
    :return:create the sheet
    """
    wb1 = load_workbook(r'{}\{}'.format(path, f_name))
    wb1.create_sheet(sheet_name, index=index)
    wb1.save(r'{}\{}'.format(path, f_name))


def remove_sheet(s_name):
    """"
    :param s_name of the sheet
    :return: removes the desired sheet
    """
    wb = openpyxl.load_workbook(r'{}\{}'.format(path, f_name))
    sheet = wb.get_sheet_by_name(s_name)
    wb.remove_sheet(sheet)
    wb.save(r'{}\{}'.format(path, f_name))


def read_row_num(s_name):
    """
    :param s_name: name of the sheet
    :return:reads the row number of the sheet desired
    """
    wb = xlrd.open_workbook(r'{}\{}'.format(path, f_name))
    wk = wb.sheet_by_name(s_name)
    existing_row = wk.nrows
    return existing_row


def add_data(s_name, data):
    """
    :param data: data to be added
    :param s_name: name of the sheet
    :return add the desired data to the desired sheet from the last row
    """
    wbk = openpyxl.load_workbook(r'{}\{}'.format(path, f_name))
    wks = wbk[s_name]
    starting_row = 1 + read_row_num(s_name)
    starting_col = 1

    for row, item in enumerate(data):
        for col, val in enumerate(item):
            wks.cell(row=row + starting_row, column=col + starting_col).value = val

    wbk.save(r'{}\{}'.format(path, f_name))
    wbk.close()


def drop_dub(s_index, keep):
    """
    :param s_index: index number of the sheet to be worked on
    :param keep: keep type of the drop_dublicates function
    :return: drops the dublicate rows of the chosen sheet
    """
    path2 = r'{}\{}'.format(path, f_name)
    df = pd.read_excel(r'{}\{}'.format(path, f_name), sheet_name=s_index, index_col=None)
    df.drop_duplicates(keep=keep, inplace=True)

    book = load_workbook(r'{}\{}'.format(path, f_name))
    writer = pd.ExcelWriter(path2, engine='openpyxl')
    writer.book = book
    df.to_excel(writer, sheet_name='final')
    writer.save()
    writer.close()


def compare(s_base, s_check, turn):
    """
    :param turn:number of times compare function is used
    :param s_base: base sheet index
    :param s_check: sheet index for the check sheet
    :return:Given a base sheet, this function tries to get the blank cell information from the check sheet.
    If the check list also does not have the information then it leaves the cell blank
    """
    wb = xlrd.open_workbook(r'{}\{}'.format(path, f_name))
    final_list = []
    base_sheet = wb.sheet_by_index(s_base)
    for i in range(base_sheet.nrows):
        list1 = []
        row_value = base_sheet.row_values(i)
        list1.extend(row_value)
        if "" in list1:
            for x, x in enumerate(list1):
                if x == "":
                    index_none = list1.index("")
                    if check_list(s_check, i)[index_none] != "":
                        list1[index_none] = check_list(s_check, i)[index_none]
                    else:
                        pass
        final_list.append(list1)
    if turn > 1:
        final_list.pop(0)
    return final_list


add_new_sheet('merge', 2)
add_data('merge', compare(0, 1, 1))
add_data('merge', compare(1, 0, 2))
drop_dub(2, 'first')
remove_sheet('merge')