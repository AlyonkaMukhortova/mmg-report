import pandas as pd
from pandas.core.frame import DataFrame
from datetime import datetime
from datetime import time


def func(df):
    res = df['BIRADS']
    print('yep', df['BIRADS'])
    file_name = 'СРМЖ итог_report.xlsx'
    df_report = pd.read_excel(file_name, index_col=0, header=0)
    if (df_report.loc[lambda x: x['МО'] == df['Организация']].loc[
        lambda x: x['Дата проведения исследований'] == pd.Timestamp(year=df['Дата исследования'].year,
                                                                    month=df['Дата исследования'].month,
                                                                    day=df['Дата исследования'].day)].empty):
        df1 = pd.DataFrame([[df['Организация'], '', df['Дата исследования'].date(), 0, 0, '', 0, '', 0, '', 0, '', 0,
                             '', 0, '', 0, 0, '', 0, '', 0, '']], columns=df_report.columns)
        df_report = df_report.append(df1, ignore_index=True)
    i = df_report.loc[lambda x: x['МО'] == df['Организация']].loc[
        lambda x: x['Дата проведения исследований'] == pd.Timestamp(year=df['Дата исследования'].year,
                                                                    month=df['Дата исследования'].month,
                                                                    day=df['Дата исследования'].day)].index
    df_report.loc[i, 'Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'] += 1
    if (res != "NOT FOUND" and int(res) < 8):
        df_report.loc[i, 'Количество BI-RADS: ' + str(res)] += 1
    if (df['PGMI'] == "PGMI: M" or df['PGMI'] == "PGMI: I"):
        df_report.loc[i, 'Количество M и I степеней по системе PGMI'] += 1
    writer = pd.ExcelWriter('СРМЖ итог_report.xlsx')
    df_report.to_excel(writer)
    writer.save()
    return df


def report(df_report):
    tv = 0
    il = 0
    """ind != 0 and"""
    length = len(df_report.index)
    for ind in range(length):
        if (df_report['Дата проведения исследований'][ind] != df_report['Дата проведения исследований'][ind + 1]):
            tv = 1
            d = df_report['Дата проведения исследований'][ind]
            dt = pd.Timestamp(d.year, d.month, d.day, 0, 0, 1)
            df1 = pd.DataFrame([['', df_report['Дата проведения исследований'][ind], dt, 0, 0, '', 0, '', 0, '', 0, '',
                                 0, '', 0, '', 0, 0, '', 0, '', 0, '']], columns=df_report.columns)
            df_report = df_report.append(df1, ignore_index=True)
            index = df_report.loc[lambda x: x['Итоги (дни)'] == df_report['Дата проведения исследований'][ind]].index
            for s in range(il, ind + 1):
                df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][index] += \
                df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][s]
        for i in range(6):
            if tv == 1:
                for s in range(il, ind + 1):
                    df_report['Количество BI-RADS: ' + str(i)][(df_report.loc[lambda x: x['Итоги (дни)'] == df_report[
                        'Дата проведения исследований'][ind]].index)] += df_report['Количество BI-RADS: ' + str(i)][s]
                df_report['% BI-RADS: ' + str(i) + ' от числа всех СРМЖ'][index] = (
                            df_report['Количество BI-RADS: ' + str(i)][index] /
                            df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][index])
            df_report['% BI-RADS: ' + str(i) + ' от числа всех СРМЖ'][ind] = (
                        df_report['Количество BI-RADS: ' + str(i)][ind] /
                        df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][ind])
        df_report['Количество исследований с указанием BI-RADS: 4-5'][ind] = df_report['Количество BI-RADS: 4'][ind] + \
                                                                             df_report['Количество BI-RADS: 5'][ind]
        df_report['Доля выбранных M и I степеней от числа всех проведенных СРМЖ'][ind] = (
                    df_report['Количество M и I степеней по системе PGMI'][ind] /
                    df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][ind])
        if tv == 1:
            df_report['Количество исследований с указанием BI-RADS: 4-5'][index] = df_report['Количество BI-RADS: 4'][
                                                                                       index] + \
                                                                                   df_report['Количество BI-RADS: 5'][
                                                                                       index]
            df_report['Доля выбранных M и I степеней от числа всех проведенных СРМЖ'][index] = (
                        df_report['Количество M и I степеней по системе PGMI'][index] /
                        df_report['Кол-во ММГ исследований скрининг рака молочной железы ЕРИС'][index])
            il = ind + 1
        tv = 0

    return df_report

#чтение файла
file_name = 'ММГ с 01.01.2022 по 09.01.2022_done.xlsx'
df = pd.read_excel(file_name, index_col=None, header=0)


df_report = DataFrame(columns=(
'МО', 'Итоги (дни)', 'Дата проведения исследований', 'Кол-во ММГ исследований скрининг рака молочной железы ЕРИС',
'Количество BI-RADS: 0', '% BI-RADS: 0 от числа всех СРМЖ',
'Количество BI-RADS: 1', '% BI-RADS: 1 от числа всех СРМЖ', 'Количество BI-RADS: 2', '% BI-RADS: 2 от числа всех СРМЖ',
'Количество BI-RADS: 3', '% BI-RADS: 3 от числа всех СРМЖ', 'Количество BI-RADS: 4', '% BI-RADS: 4 от числа всех СРМЖ',
'Количество BI-RADS: 5', '% BI-RADS: 5 от числа всех СРМЖ', 'Количество исследований с указанием BI-RADS: 4-5',
'Количество BI-RADS: 6', '% BI-RADS: 6 от числа всех СРМЖ',
'Количество BI-RADS: 7', '% BI-RADS: 7 от числа всех СРМЖ',
'Количество M и I степеней по системе PGMI', 'Доля выбранных M и I степеней от числа всех проведенных СРМЖ'))
writer = pd.ExcelWriter('СРМЖ итог_report.xlsx')
df_report.to_excel(writer)
writer.save()
df['Дата создания записи'] = pd.to_datetime(df['Дата создания записи'], format='%d.%m.%Y %H:%M:%S')
df['Дата исследования'] = pd.to_datetime(df['Дата исследования'], format='%d.%m.%Y %H:%M:%S')

df.sort_values(by=['Дата создания записи'], inplace=True, ignore_index=True)

df.drop_duplicates(subset='UID исследования', keep='last', inplace=True, ignore_index=True)
counter = []
file_name = 'СРМЖ итог_report.xlsx'
df = df.apply(func, axis=1, result_type='expand')
df_report = pd.read_excel(file_name, index_col=0, header=0)

df_report.sort_values(by=['Дата проведения исследований'], inplace=True, ignore_index=True)
df_report = report(df_report)
df_report.sort_values(by=['Дата проведения исследований'], inplace=True, ignore_index=True)

b0 = len(df.loc[lambda x: x['BIRADS'] == 0].index)
print(b0)

print(len(df_report.index))

writer = pd.ExcelWriter('2022.01.26_ММГ итог_report_01.01-09.01_v6.xlsx')
sheet_name = 'Отчет'
df_report.to_excel(writer, sheet_name=sheet_name)
sheet_name = 'Обработанная выгрузка'
df.to_excel(writer, sheet_name=sheet_name)
writer.save()
