import xlrd
import xlwt
import os

font = xlwt.Font()
font.name = 'Times New Roman'
font.colour_index = 2
font.bold = True
style = xlwt.XFStyle()
style.font = font


def xls_reader(path, raws, columns, doc_type):
    workbook = xlrd.open_workbook(path, on_demand=True, formatting_info=True)
    worksheet = workbook.sheet_by_index(0)
    res_list = []
    for i in range(1, raws):  # количество считываемых строк таблицы
        tmp_list = []
        tmp_dict = {}
        for j in range(0, columns):  # количество считываемых столбцов таблицы
            value = worksheet.cell_value(i, j)
            tmp_list.append(value)
        if doc_type == "BOM":
            tmp_dict["Designator"] = str(tmp_list[0])
            tmp_dict["Part Number"] = str(tmp_list[1]).lower().split(" ")[0]
            tmp_dict["Value"] = tmp_list[2]
            tmp_dict["Quantity"] = tmp_list[3]
            tmp_dict["Description"] = tmp_list[4]
        elif doc_type == "STORE":
            tmp_dict["Код"] = str(tmp_list[0])
            tmp_dict["Номенклатура"] = str(tmp_list[1]).lower()
            tmp_dict["Количество"] = int(tmp_list[2])
            tmp_dict["Склад"] = tmp_list[4]
            tmp_dict["Адрес хранения"] = tmp_list[5]
        res_list.append(tmp_dict)
    return res_list


def write_in_excel(bom, storage, to_path):
    wb = xlwt.Workbook()
    ws1 = wb.add_sheet('Found', cell_overwrite_ok=True)

    found_list = []
    raw = 0
    for bom_item in bom.copy():
        for store_item in storage.copy():
            if raw != 0:
                if bom_item["Part Number"] and bom_item["Part Number"] in store_item['Номенклатура']:
                    write_lst = [bom_item['Part Number'], bom_item['Value'], bom_item['Quantity'], store_item['Код'],
                                 store_item['Номенклатура'], store_item['Количество'], store_item['Адрес хранения']]
                    for col, label in enumerate(write_lst):
                        ws1.write(raw, col, label)
                    raw += 1
                    found_list.append(store_item)
                    storage.remove(store_item)
                    if bom_item in bom:
                        bom.remove(bom_item)
            else:
                write_lst = ["Номенклатура", 'Параметр', 'Количество', 'Код', 'Номенклатура', 'Количество',
                             'Адрес хранения']
                for col, label in enumerate(write_lst):
                    ws1.write(raw, col, label)
                raw += 1

    ws2 = wb.add_sheet('Not Found', cell_overwrite_ok=True)
    raw = 0
    for bom_item in bom.copy():
        if raw != 0:
            write_lst = [bom_item['Designator'], bom_item['Part Number'], bom_item['Value'], bom_item['Quantity']]
        else:
            write_lst = ['Designator', 'Part Number', 'Value', 'Quantity']

        for col, label in enumerate(write_lst):
            ws2.write(raw, col, label)
        raw += 1

    ws3 = wb.add_sheet('STORAGE', cell_overwrite_ok=True)
    raw = 0
    for store_item in storage.copy():
        if raw != 0:
            write_lst = [store_item['Код'], store_item['Номенклатура'], store_item['Количество'],
                         store_item['Адрес хранения']]
        else:
            write_lst = ['Код', 'Номенклатура', 'Количество', 'Адрес хранения']

        for col, label in enumerate(write_lst):
            ws3.write(raw, col, label)
        raw += 1

    wb.save(to_path)


def main():
    bom_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                    "\Bill of Materials-РВ-21.31.701-2.xls")
    storage_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                        "\Высотомер_26 июня.2023.xls")
    to_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                   "\BOM_founded.xls")
    rows = 131
    columns = 5
    bom_list = xls_reader(bom_file_path, rows, columns, 'BOM')
    rows = 595
    columns = 6
    storage_list = xls_reader(storage_file_path, rows, columns, 'STORE')[1:]
    write_in_excel(bom_list, storage_list, to_file_path)


if __name__ == '__main__':
    main()
