import xlwt
import re
import os

from excel_reader import xls_reader

font = xlwt.Font()
font.name = 'Times New Roman'
font.colour_index = 2
font.bold = True
style = xlwt.XFStyle()
style.font = font

types = {'C': 'Конденсаторы',
         'D': 'Микросхемы',
         'DA': 'Микросхемы аналоговые',
         'DD': 'Микросхемы цифровые',
         'FU': 'Предохранители',
         'G': 'Генераторы',
         'HL': 'Устройства индикации',
         'K': 'Реле',
         'L': 'Дроссели, Катушки индуктивности',
         'Q': 'Кварцы',
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


def to_modify_list(array: list) -> list:
    """
    Function grouping elements by "Designator" (first) and "Part Number" (second)
    :param array: source list
    :return: modified list
    """
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


def write_in_excel(array: list, path: str) -> None:
    """
    Function writes modified list in Excel file
    :param array: incomming list
    :param path: exiting file path
    :return: None
    """
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
        elif cur_designator == 'BQ' or cur_designator == 'ZQ':
            cur_designator = 'Q'
        if cur_designator != prev_designator:
            row += 1
            write_lst = f'{types.get(cur_designator, "Компонент")}'
            ws.write(row, 1, write_lst)
            prev_designator = cur_designator
            row += 1
        if item['Comment'] == 'NP':
            comment = '*'
        else:
            comment = ' '
        if not item['Part Number']:
            comment = '**'
        write_lst = [item['Designator'], f"{item['Part Number']}, {item['Manufacturer']}", item['Quantity'], comment]
        for col, label in enumerate(write_lst):
            ws.write(row, col, label)
        row += 1
    row += 1
    note_list = [['*', 'Не устанавливать'], ['**', 'Конструктивный элемент']]
    for note in note_list:
        for col, label in enumerate(note):
            ws.write(row, col, label)
        row += 1

    wb.save(path)


def make_bom_gost(from_file: str, to_file: str) -> None:
    """
    Function performs a sequence of actions.
    :param from_file: incoming file path
    :param to_file: exiting file path
    :return: None
    """
    bom_list = xls_reader(from_file)
    modified_bom_list = to_modify_list(bom_list)
    write_in_excel(modified_bom_list, to_file)


if __name__ == '__main__':
    from_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                     "\Bill_of_Materials_no_group_by_pn-РВ-21.31.701-2.xls")
    to_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                   "\РВ-21.31.701-2-ПЭ3.xls")
    make_bom_gost(from_file_path, to_file_path)
