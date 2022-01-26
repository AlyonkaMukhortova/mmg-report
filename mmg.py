import pandas as pd
from pandas.core.frame import DataFrame
from datetime import datetime
from datetime import time

dict_num = {'I ': 1, 'II': 2, 'V': 5, 'IV': 4, 'III': 3}


def find_num(z, ptr):
    z = str(z)
    res = -1
    res1 = -1
    res2 = -1
    ptr1 = -1
    ptr11 = -1
    ptr2 = -1
    for value in dict_num.keys():
        ptr11 = z.upper().find(value)
        print(ptr11)
        if (ptr11 != -1 and ptr11 > ptr):
            res1 = dict_num[value]
            ptr1 = ptr11


    while (not z[ptr].isdigit() and ptr < len(z) - 1):
        ptr += 1
    if (z[ptr].isdigit()):
        res2 = int(z[ptr])
        ptr2 = ptr
    if (ptr2 == -1 and ptr1 == -1):
        res = 'NOT FOUND'
    elif (ptr1 == -1):
        res = res2
    elif (ptr2 == -1):
        res = res1
    elif(ptr1 < ptr2):
        res = res1
    else:
        res = res2
    print('attention!', res1, ptr1, res2, ptr2, '\n', z, '\nend attention\n')
    return res


def func(df):
    file_name = 'СРМЖ итог_report.xlsx'
    df_report = pd.read_excel(file_name, index_col=0, header=0)
    z = df['Заключение']
    z = str(z)
    ptr = z.lower().find('birads')
    if (ptr != -1):
        ptr += 5
        res1 = find_num(z, ptr)
    elif (z.lower().find('ds') != -1):
        ptr = z.lower().find('ds')
        ptr += 2
        res1 = find_num(z, ptr)
    elif (z.lower().find('левая молочная железа') != -1 or z.lower().find('правая молочная железа') != -1):
        if (z.lower().find('левая молочная железа') != -1 and z.lower().find('правая молочная железа') != -1):
            ptr = min(z.lower().find('левая молочная железа'), z.lower().find('правая молочная железа'))
        elif (z.lower().find('левая молочная железа') != -1):
            ptr = z.lower().find('левая молочная железа')
        else:
            ptr = z.lower().find('правая молочная железа')
        ptr += 6
        res1 = find_num(z, ptr)
    else:
        res1 = 'NOT FOUND'
    if (res1 != 'NOT FOUND'):
        z = z[ptr + 4:]
    ptr = z.lower().find('birads')
    if (ptr != -1):
        ptr += 5
        res2 = find_num(z, ptr)
    elif(z.lower().find('ds') != -1):
        ptr = z.lower().find('ds')
        ptr += 2
        res2 = find_num(z, ptr)
    elif (z.lower().find('левая молочная железа') != -1 or z.lower().find('правая молочная железа') != -1):
        if (z.lower().find('левая молочная железа') != -1 and z.lower().find('правая молочная железа') != -1):
            ptr = min(z.lower().find('левая молочная железа'), z.lower().find('правая молочная железа'))
        elif (z.lower().find('левая молочная железа') != -1):
            ptr = z.lower().find('левая молочная железа')
        else:
            ptr = z.lower().find('правая молочная железа')
        ptr += 6
        res2 = find_num(z, ptr)
    else:
        res2 = 'NOT FOUND'
    if (res1 == 0 or res2 == 0):
        if (res1 == 4 or res2 == 4):
            res = 4
        else:
            res = 0
    else:
        if (res1 != 'NOT FOUND' and res2 != 'NOT FOUND'):
            res = max(res1, res2)
        elif (res1 != 'NOT FOUND'):
            res = res1
        elif (res2 != 'NOT FOUND'):
            res = res2
        else:
            res = 'NOT FOUND'
        # print('res = ',res)

    df['BIRADS'] = res
    print(res)
    ptr = -1
    if (not isinstance(df['Описание'], float)):
        ptr = df['Описание'].lower().find('pgmi')
    if (ptr == -1):
        df['PGMI'] = 'NO DATA'
    else:
        df['PGMI'] = df['Описание'][ptr:ptr + 7]
    return df



#чтение файла
file_name = 'ММГ с 01.01.2022 по 09.01.2022.xlsx'
df = pd.read_excel(file_name, index_col=None, header=0)
df['BIRADS'] = ''
df['PGMI'] = ''

df['Дата создания записи'] = pd.to_datetime(df['Дата создания записи'], format='%d.%m.%Y %H:%M:%S')
df['Дата исследования'] = pd.to_datetime(df['Дата исследования'], format='%d.%m.%Y %H:%M:%S')

df.sort_values(by=['Дата создания записи'], inplace=True, ignore_index=True)

df.drop_duplicates(subset='UID исследования', keep='last', inplace=True, ignore_index=True)
counter = []
file_name = 'СРМЖ итог_report.xlsx'
df = df.apply(func, axis=1, result_type='expand')


b0 = len(df.loc[lambda x: x['BIRADS'] == 0].index)
print(b0)



writer = pd.ExcelWriter('ММГ с 01.01.2022 по 09.01.2022_done.xlsx')
sheet_name = 'Обработанная выгрузка'
df.to_excel(writer, sheet_name=sheet_name)
writer.save()
