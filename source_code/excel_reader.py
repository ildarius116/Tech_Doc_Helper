import xlrd
import xlwt
import os

font = xlwt.Font()
font.name = 'Times New Roman'
font.colour_index = 2
font.bold = True
style = xlwt.XFStyle()
style.font = font


def xls_reader(path: str, doc_type: str = None) -> list:
    """
    Function read Excel file and return List of this file contents.
    :param path: path of incoming Excel file
    :param doc_type: pattern of file processing
    :return: list of file contents
    """
    workbook = xlrd.open_workbook(path, on_demand=True, formatting_info=True)
    worksheet = workbook.sheet_by_index(0)
    rows_numbers = worksheet.nrows
    col_numbers = worksheet.ncols
    res_list = []
    for i in range(1, rows_numbers):
        tmp_list = []
        tmp_dict = {}
        for j in range(col_numbers):
            value = worksheet.cell_value(i, j)
            tmp_list.append(value)
        if not doc_type:
            tmp_dict["Designator"] = tmp_list[0]
            tmp_dict["Part Number"] = tmp_list[1]
            tmp_dict["Value"] = tmp_list[2]
            tmp_dict["Quantity"] = int(tmp_list[3])
            tmp_dict["Manufacturer"] = tmp_list[4]
            tmp_dict["Comment"] = tmp_list[4]
        if doc_type == "BOM":
            tmp_dict["Designator"] = str(tmp_list[0])
            tmp_dict["Part Number"] = str(tmp_list[1]).lower().split(" ")[0]
            tmp_dict["Value"] = tmp_list[2]
            tmp_dict["Quantity"] = int(tmp_list[3])
            tmp_dict["Manufacturer"] = tmp_list[4]
        elif doc_type == "STORE":
            tmp_dict["Код"] = str(tmp_list[0])
            tmp_dict["Номенклатура"] = str(tmp_list[1]).lower()
            tmp_dict["Количество"] = int(tmp_list[2])
            tmp_dict["Склад"] = tmp_list[4]
            tmp_dict["Адрес хранения"] = tmp_list[5]
        res_list.append(tmp_dict)
    return res_list


def main():
    pass


if __name__ == '__main__':
    main()
