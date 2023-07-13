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

types = {'BQ': ('Кварц', 'Кварцы'),
         'C': ('Конденсатор', 'Конденсаторы'),
         'D': ('Микросхема', 'Микросхемы'),
         'DA': ('Микросхема аналоговая', 'Микросхемы аналоговые'),
         'DD': ('Микросхема цифровая', 'Микросхемы цифровые'),
         'FU': ('Предохранитель', 'Предохранители'),
         'G': ('Генератор', 'Генераторы'),
         'HL': ('Устройство индикации', 'Устройства индикации'),
         'K': ('Реле', 'Реле'),
         'L': ('Дроссель', 'Дроссели'),
         'R': ('Резистор', 'Резисторы'),
         'SB': ('Переключатель', 'Переключатели'),
         'T': ('Трансформатор', 'Трансформаторы'),
         'U': ('DC/DC преобразователь', 'DC/DC преобразователи'),
         'VD': ('Диод', 'Диоды'),
         'VT': ('Транзистор', 'Транзисторы'),
         'X': ('Соединитель', 'Соединители'),
         'XP': ('Соединитель', 'Соединители'),
         'XS': ('Соединитель', 'Соединители'),
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
        tmp_dict["Manufacturer"] = tmp_list[4]
        res_list.append(tmp_dict)
    return res_list


def convert_list_to_dict(array):
    type_dict = {}
    for bom_item in array:
        bom_designator = bom_item['Designator'].split(', ')[0]
        bom_designator_sym = re.findall('\D+', bom_designator)[0]
        designators = re.split('[, ]+', bom_item['Designator'])
        bom_item['Designator'] = []
        tmp_lst = []
        prev_elem = ''
        quantity = 0
        for i, designator in enumerate(designators):
            if not tmp_lst:
                tmp_lst = [designator]
                prev_elem = designator
                quantity = 1
            else:
                if len(tmp_lst[0]) > 5:
                    bom_item['Designator'].append(tmp_lst[0])
                    tmp_lst = [designator]
                    prev_elem = designator
                    quantity = 1
                elif len(designator) > 5:
                    bom_item['Designator'].append(tmp_lst[0])
                    bom_item['Designator'].append(designator)
                    tmp_lst = []
                    prev_elem = designator
                    quantity = 0
                else:
                    bom_item_designator_num = int(re.findall('\d+', designator)[0])
                    tmp_elem_designator_num = int(re.findall('\d+', prev_elem)[0])
                    diff_designator = bom_item_designator_num - tmp_elem_designator_num - 1
                    if not diff_designator:
                        prev_elem = designator
                        quantity += 1
                        if designator == designators[-1]:
                            if quantity > 2:
                                text = f'{tmp_lst[0]}-{prev_elem}'
                            else:
                                text = f'{tmp_lst[0]},{designator}'
                            bom_item['Designator'].append(text)
                            tmp_lst = []
                            prev_elem = designator
                            quantity = 0
                    else:
                        if quantity == 1:
                            bom_item['Designator'].append(tmp_lst[0])
                            bom_item['Designator'].append(designator)
                            tmp_lst = []
                            prev_elem = designator
                            quantity = 0
                        if quantity == 2:
                            bom_item['Designator'].append(tmp_lst[0])
                            bom_item['Designator'].append(prev_elem)
                            tmp_lst = [designator]
                            prev_elem = designator
                            quantity = 1
                        elif quantity > 2:
                            text = f'{tmp_lst[0]}-{prev_elem}'
                            bom_item['Designator'].append(text)
                            tmp_lst = [designator]
                            prev_elem = designator
                            quantity = 1
            if tmp_lst and (i + 1) == len(designators):
                bom_item['Designator'].extend(tmp_lst)
        if bom_designator_sym not in type_dict:
            type_dict[bom_designator_sym] = [bom_item]
        else:
            type_dict[bom_designator_sym].append(bom_item)
    compact_dict(type_dict)
    sort_dict(type_dict)
    modified_dict = modify_dict(type_dict)
    return modified_dict


def compact_dict(dct):
    tmp_lst = []
    for value in dct.values():
        for elem in value:
            designators = elem['Designator'].copy()
            elem['Designator'] = []
            for i, designator in enumerate(designators):
                if len(designators) == 1:
                    elem['Designator'].append(designator)
                elif not tmp_lst:
                    tmp_lst = [designator]
                else:
                    if len(tmp_lst[0]) > 5:
                        elem['Designator'].append(tmp_lst[0])
                        tmp_lst = [designator]
                    elif len(designator) > 5:
                        elem['Designator'].append(tmp_lst[0])
                        elem['Designator'].append(designator)
                        tmp_lst = []
                    else:
                        text = f'{tmp_lst[0]},{designator}'
                        elem['Designator'].append(text)
                        tmp_lst = []
                if tmp_lst and (i + 1) == len(designators):
                    elem['Designator'].extend(tmp_lst)
                    tmp_lst = []


def modify_dict(dct):
    modified_dict = {}
    for key, item_parts in dct.items():
        for item in item_parts:
            manufacturer = item['Manufacturer']
            if key not in modified_dict:
                modified_dict[key] = {manufacturer: [item]}
            else:
                if manufacturer not in modified_dict[key]:
                    modified_dict[key].update({manufacturer: [item]})
                else:
                    modified_dict[key][manufacturer].append(item)
    return modified_dict


def sort_dict(dct):
    for key, item in dct.items():
        # print(item)
        item = sorted(item, key=lambda x: (x['Manufacturer'], x['Part Number']))
        dct[key] = item
        # print(item)
        pass


def write_in_excel(dictionary, pos, path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Спецификация', cell_overwrite_ok=True)

    prev_manufacturer = ''
    write_lst = 'Прочие изделия'
    ws.write(0, 2, write_lst)
    row = 1
    for key_i, value_i in dictionary.items():
        for key_j, value_J in value_i.items():
            if key_j != '-':
                for elem in value_J:
                    cur_manufacturer = key_j
                    cur_designator = elem['Designator'][0]
                    cur_designator_sym = re.findall('\D+', cur_designator)[0]
                    if cur_designator_sym == 'DA' or cur_designator_sym == 'DD':
                        cur_designator_sym = 'D'
                    elif cur_designator_sym == 'XP' or cur_designator_sym == 'XS':
                        cur_designator_sym = 'X'
                    if cur_manufacturer != prev_manufacturer:
                        row += 1
                        if cur_designator_sym in types.keys():
                            if len(value_J) > 1:
                                write_lst = f"{types[cur_designator_sym][1]} {cur_manufacturer}"
                            else:
                                write_lst = f"{types[cur_designator_sym][0]} {elem['Part Number']}"
                        else:
                            write_lst = f"Прочее"
                        ws.write(row, 2, write_lst)
                        prev_manufacturer = cur_manufacturer
                        row += 1
                    designators_list = elem['Designator']
                    for i, designator in enumerate(designators_list):
                        if i == 0:
                            if len(value_J) > 1:
                                write_lst = [pos, '', elem['Part Number'], elem['Quantity'], designator]
                            else:
                                write_lst = [pos, '', cur_manufacturer, elem['Quantity'], designator]
                            pos += 1
                        else:
                            write_lst = ['', '', '', '', designator]
                        for col, label in enumerate(write_lst):
                            ws.write(row, col, label)
                        row += 1

    wb.save(path)


def main():
    from_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                     "\Bill of Materials-РВ-21.31.701-2.xls")
    to_file_path = os.path.abspath("D:\Проекты\RV-21.31.701-2_ver.3\Project Outputs for РВ-21.31.701-2\BOM"
                                   "\РВ-21.31.701-2 - Спецификация.xls")
    rows = 133
    columns = 5
    start_position = 2
    bom_list = xls_reader(from_file_path, rows, columns)
    type_dict = convert_list_to_dict(bom_list)
    write_in_excel(type_dict, start_position, to_file_path)


if __name__ == '__main__':
    main()
