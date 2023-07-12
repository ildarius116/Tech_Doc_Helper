import xlrd
import xlwt
import re
import os

font = xlwt.Font()
font.name = 'Times New Roman'
font.colour_index = 2
font.bold = True
style = xlwt.XFStyle()
style.font = font

types = {'BQ': 'Кварцы',
         'C': 'Конденсаторы',
         'D': 'Микросхемы',
         'DA': 'Микросхемы аналоговые',
         'DD': 'Микросхемы цифровые',
         'FU': 'Предохранители',
         'G': 'Генераторы',
         'HL': 'Устройства индикации',
         'K': 'Реле',
         'L': 'Дроссели, Катушки индуктивности',
         'R': 'Резисторы',
         'SB': 'Переключатели',
         'T': 'Резисторы',
         'U': 'DC/DC преобразователи',
         'VD': 'Диоды',
         'VT': 'Транзисторы',
         'X': 'Соединители',
         'XP': 'Соединители',
         'XS': 'Соединители',
         }


def xls_reader(path, raws, columns):
    workbook = xlrd.open_workbook(path, on_demand=True, formatting_info=True)
    worksheet = workbook.sheet_by_index(0)
    res_list = []
    for i in range(1, raws):  # количество считываемых строк таблицы
        tmp_list = []
        tmp_dict = {}
        for j in range(0, columns):  # количество считываемых столбцов таблицы
            value = worksheet.cell_value(i, j)
            tmp_list.append(value)
        tmp_dict["Designator"] = tmp_list[0]
        tmp_dict["Part Number"] = tmp_list[1]
        tmp_dict["Value"] = tmp_list[2]
        tmp_dict["Quantity"] = int(tmp_list[3])
        tmp_dict["Comment"] = tmp_list[4]
        res_list.append(tmp_dict)
    return res_list


def to_modify_list(array):
    modified_bom_list = []
    tmp_elem = array[0]
    start_designator = tmp_elem['Designator']
    prev_designator = tmp_elem['Designator']
    for bom_item in array[1:].copy():
        bom_item_designator_num = int(re.findall('\d+', bom_item['Designator'])[0])
        tmp_elem_designator_num = int(re.findall('\d+', tmp_elem['Designator'])[0])
        diff_designator = bom_item_designator_num - tmp_elem_designator_num
        diff_elem_q = tmp_elem['Quantity'] - diff_designator

        if tmp_elem['Part Number'] == bom_item['Part Number'] and not diff_elem_q and bom_item['Comment'] != 'NP' and \
                tmp_elem['Comment'] != 'NP':
            prev_designator = bom_item['Designator']
            tmp_elem['Quantity'] += bom_item['Quantity']
        elif tmp_elem['Part Number'] == bom_item['Part Number'] and not diff_elem_q and bom_item['Comment'] == 'NP' and \
                tmp_elem['Comment'] == 'NP':
            prev_designator = bom_item['Designator']
            tmp_elem['Quantity'] += bom_item['Quantity']
        else:
            if start_designator == prev_designator:
                modified_bom_list.append(tmp_elem)
                tmp_elem = bom_item
            else:
                if tmp_elem['Quantity'] == 2:
                    tmp_elem['Designator'] += f",{prev_designator}"
                    start_designator = prev_designator
                elif tmp_elem['Quantity'] > 2:
                    tmp_elem['Designator'] += f"-{prev_designator}"
                    start_designator = prev_designator
                modified_bom_list.append(tmp_elem)
                tmp_elem = bom_item
            prev_designator = bom_item['Designator']
    return modified_bom_list


def write_in_excel(array, path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Bill of Materials', cell_overwrite_ok=True)

    prev_designator = ''
    row = 0
    for item in array:
        cur_designator = re.findall('\D*', item['Designator'])[0]
        if cur_designator == 'DA' or cur_designator == 'DD':
            cur_designator = 'D'
        elif cur_designator == 'XP' or cur_designator == 'XS':
            cur_designator = 'X'
        if cur_designator != prev_designator:
            row += 1
            if cur_designator in types.keys():
                write_lst = f'{types[cur_designator]}'
            else:
                write_lst = 'Компонент'
            ws.write(row, 1, write_lst)
            prev_designator = cur_designator
            row += 1
        if item['Comment'] == 'NP':
            comment = '*'
        else:
            comment = ' '
        if not item['Part Number']:
            comment = '**'
        write_lst = [item['Designator'], item['Part Number'], item['Quantity'], comment]
        for col, label in enumerate(write_lst):
            ws.write(row, col, label)
        row += 1
    row += 1
    note_list = [['*', 'Не паять'], ['**', 'Тестовые точки']]
    for note in note_list:
        for col, label in enumerate(note):
            ws.write(row, col, label)
        row += 1

    wb.save(path)


def main():
    from_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                     "\Copy of Bill of Materials-РВ-21.31.701-2.xls")
    to_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                   "\РВ-21.31.701-2-ПЭ3.xls")
    rows = 837
    columns = 5
    bom_list = xls_reader(from_file_path, rows, columns)
    modified_bom_list = to_modify_list(bom_list)
    write_in_excel(modified_bom_list, to_file_path)


if __name__ == '__main__':
    main()