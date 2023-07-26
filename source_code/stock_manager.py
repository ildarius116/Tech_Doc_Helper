import xlwt
import os

from excel_reader import xls_reader

font = xlwt.Font()
font.name = 'Times New Roman'
font.colour_index = 2
font.bold = True
style = xlwt.XFStyle()
style.font = font


def write_in_excel(bom: str, storage: str, to_path: str) -> None:
    """
    Function writes to exiting file:
     1. "Not Found" (in 1st sheet) - elements which wasn't found
     2. "Found" (in 2nd sheet) -  elements which was found
     3. "STORAGE" (in 3rd sheet) -  remaining list items in storage
    :param bom: Bill_of_Materials file path
    :param storage: Storage file path
    :param to_path: Exiting file path
    :return: None
    """
    wb = xlwt.Workbook()
    ws1 = wb.add_sheet('Not Found', cell_overwrite_ok=True)
    ws2 = wb.add_sheet('Found', cell_overwrite_ok=True)
    ws3 = wb.add_sheet('STORAGE', cell_overwrite_ok=True)

    found_list = []
    raw = 0
    for bom_item in bom.copy():
        prev_part = ''
        prev_code = ''
        for store_item in storage.copy():
            if raw != 0:
                if bom_item["Part Number"] and bom_item["Part Number"] in store_item['Номенклатура']:
                    if bom_item["Part Number"] != prev_part:
                        write_lst = [bom_item['Part Number'].upper(),
                                     bom_item['Value'],
                                     bom_item['Quantity'],
                                     store_item['Код'],
                                     store_item['Номенклатура'],
                                     store_item['Количество'],
                                     store_item['Адрес хранения'],
                                     store_item['Склад']
                                     ]
                        prev_part = bom_item["Part Number"]
                    else:
                        if store_item['Код'] != prev_code:
                            write_lst = ['', '', '',
                                         store_item['Код'],
                                         store_item['Номенклатура'],
                                         store_item['Количество'],
                                         store_item['Адрес хранения'],
                                         store_item['Склад']
                                         ]
                        else:
                            write_lst = ['', '', '', '',
                                         store_item['Номенклатура'],
                                         store_item['Количество'],
                                         store_item['Адрес хранения'],
                                         store_item['Склад']
                                         ]
                    for col, label in enumerate(write_lst):
                        ws2.write(raw, col, label)
                    raw += 1
                    found_list.append(store_item)
                    storage.remove(store_item)
                    if bom_item in bom:
                        bom.remove(bom_item)
            else:
                write_lst = ["Номенклатура", 'Параметр', 'Количество', 'Код', 'Номенклатура', 'Количество',
                             'Адрес хранения', 'Склад']
                for col, label in enumerate(write_lst):
                    ws2.write(raw, col, label)
                raw += 1

    raw = 0
    for bom_item in bom.copy():
        if raw != 0:
            write_lst = [bom_item['Designator'], bom_item['Part Number'].upper(), bom_item['Value'], bom_item['Quantity'],
                         bom_item['Manufacturer']]
        else:
            write_lst = ['Designator', 'Part Number', 'Value', 'Quantity', 'Manufacturer']

        for col, label in enumerate(write_lst):
            ws1.write(raw, col, label)
        raw += 1

    raw = 0
    for store_item in storage.copy():
        if raw != 0:
            write_lst = [store_item['Код'], store_item['Номенклатура'], store_item['Количество'],
                         store_item['Адрес хранения'], store_item['Склад']]
        else:
            write_lst = ['Код', 'Номенклатура', 'Количество', 'Адрес хранения', 'Склад']

        for col, label in enumerate(write_lst):
            ws3.write(raw, col, label)
        raw += 1

    wb.save(to_path)


def find_in_storage(bom_file: str, storage_file: str, to_file: str) -> None:
    """
    Function performs a sequence of actions.
    :param bom_file: incoming Bill_of_Materials file path
    :param storage_file: incoming Storage file path
    :param to_file: exiting file path
    :return: None
    """
    bom_list = xls_reader(bom_file, 'BOM')
    storage_list = xls_reader(storage_file, 'STORE')[1:]
    write_in_excel(bom_list, storage_list, to_file)


if __name__ == '__main__':
    bom_file_path = os.path.abspath("D:\Проекты\PB-21_tool\PB-21_tool_main\Project Outputs for PB-21_tool_main\BOM"
                                    "\Bill of Materials-PB-21_tool_main.xls")
    storage_file_path = os.path.abspath("D:\Проекты\PB-21_tool\PB-21_tool_main\Project Outputs for PB-21_tool_main\BOM"
                                        "\Высотомер_13 июля.2023.xls")
    to_file_path = os.path.abspath("D:\Проекты\PB-21_tool\PB-21_tool_main\Project Outputs for PB-21_tool_main\BOM"
                                   "\BOM_found_0.xls")
    find_in_storage(bom_file_path, storage_file_path, to_file_path)
