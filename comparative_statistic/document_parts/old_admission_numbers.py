import os
from datetime import datetime, date
import numpy as np
import pandas as pd
from aux_program_scripts.collect_programs import join_new_old_progs
from aux_program_scripts.write_programs import write_programs


def masks_lvl(path, dirs, CAMPUS):
    masks = [[], [], []]

    for dir_ in dirs:
        files = os.listdir(path + f'/{dir_.strftime('%Y-%m-%d')}')
        mask = [False, False, False]

        for file in files:
            if 'сравнительная' in file.lower():
                if (CAMPUS == 'Москва' and 'сравнительная' == file[:13].lower()) or (CAMPUS != 'Москва' and CAMPUS in file):
                    if 'бакалавриат' in file.lower():
                        mask[0] = True
                    if 'магистратура' in file.lower():
                        mask[1] = True

        for i in range(len(masks)):
            masks[i].append(mask[i])

    return masks


def closest_date_idx(dates, CUR_DATE, end_of_campaign=date(year=2023, month=7, day=25),
                     cur_end_of_campaign=date(year=2024, month=7, day=25)):
    days_till_end = list(map(lambda x: (end_of_campaign - x).days, dates))
    cur_days_till_end = (cur_end_of_campaign - CUR_DATE.date()).days

    idx = np.argmin(abs(np.array(days_till_end) - cur_days_till_end))

    return idx, days_till_end[idx], cur_days_till_end


def fetch_old_df(path, date, lvl, CAMPUS):
    files = os.listdir(path + f'/{date.strftime('%Y-%m-%d')}')

    names = ['', '', '']

    for file in files:
        if '~$' in file:
            continue

        if 'сравнительная' in file.lower():
            if ((CAMPUS == 'Москва' and 'сравнительная' == file[:13].lower()) or (
                    CAMPUS != 'Москва' and CAMPUS in file)):
                if 'бакалавриат' in file.lower():
                    names[0] = file
                if 'магистратура' in file.lower():
                    if 'с доп столбцами' in file.lower():
                        names[2] = file
                    else:
                        names[1] = file

    return pd.read_excel(
        f'/Users/s/Desktop/Admission Numbers 2024/tables_2023/{date.strftime('%Y-%m-%d')}/{names[lvl]}')


def write_old_admission(wb, ws, programs, lvl, CAMPUS, CUR_DATE, info=False, col_correspondence=None, PAID_ONLY=False):
    if CAMPUS == 'Нижний Новгород':
        CAMPUS = 'Нижний Новгород'

    dates = np.array([datetime.strptime(dt, '%Y-%m-%d').date()
                      for dt in os.listdir('/Users/s/Desktop/Admission Numbers 2024/tables_2023') if '2023' in dt])

    masks = masks_lvl('/Users/s/Desktop/Admission Numbers 2024/tables_2023', dates, CAMPUS)

    idx, days_till_end, cur_days_till_end = closest_date_idx(dates[masks[bool(lvl)]], CUR_DATE)
    df_old = fetch_old_df('/Users/s/Desktop/Admission Numbers 2024/tables_2023',
                          dates[masks[bool(lvl)]][idx], bool(lvl), CAMPUS)
    print(df_old)
    if lvl == 0:
        last_stat = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/last_year_bac_{CAMPUS}.xlsx')
        if CAMPUS == 'Москва':
            df_old.iloc[:, 22] = last_stat.iloc[:, 22]
            columns_old = [0, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 24, 22]
        else:
            df_old.iloc[:, 23] = last_stat.iloc[:, 23]
            columns_old = [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 25, 23]

        columns_names = ['name', 'kcp_old', 'work_old', 'special_old', 'separate_old',
                         'kcp_admit_old', 'work_admit_old', 'special_admit_old', 'separate_admit_old',
                         'bvi_admit_old', 'vsosh_admit_old', 'paid_admit_old', 'paid_old']

    elif lvl >= 1:
        last_stat = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/last_year_mag_{CAMPUS}.xlsx')

        if not PAID_ONLY:
            if 'высшая лига' in df_old.iloc[3, 14].lower():
                if CAMPUS == 'Москва':
                    columns_old = [0, 1, 5, 2, 7, 9, 21, 19]
                    df_old.iloc[:, 19] = last_stat.iloc[:, 19]
                else:
                    columns_old = [0, 2, 6, 4, 8, 10, 22, 20]
                    df_old.iloc[:, 20] = last_stat.iloc[:, 20]
            else:
                columns_old = [0, 1, 3, 2, 7, 9, 14, 13]
                df_old.iloc[:, 13] = last_stat.iloc[:, 12]

            columns_names = ['name', 'kcp_old', 'work_old', 'school_old',
                             'kcp_admit_old', 'work_admit_old', 'paid_admit_old', 'paid_old']
        else:
            columns_old = [0, 2, 4]

            columns_names = ['name', 'paid_old', 'paid_admit_old']

    if not PAID_ONLY:
        df_old_total = df_old.iloc[5:, [0, 1]]
        row_total_students = df_old_total.shape[0] - 1

        while pd.isna(df_old_total.iloc[row_total_students, 1]):
            row_total_students -= 1
        if CAMPUS == 'Москва':
            old_tot_inc_students = df_old_total.iloc[row_total_students - 1, 1]
        else:
            old_tot_inc_students = df_old_total.iloc[row_total_students, 1]
    else:
        old_tot_inc_students = 0

    if not PAID_ONLY:
        df_old = df_old.iloc[5:, columns_old].reset_index(drop=True)
    else:
        df_old = df_old.iloc[4:, columns_old].reset_index(drop=True)
    df_old.columns = columns_names

    if lvl == 0:
        for col in ['kcp', 'work', 'special', 'separate', 'paid']:
            print(df_old)
            df_old.loc[:, col+'_admit_old'] = df_old.apply(
                lambda x: 0 if pd.isna(x[col+'_admit_old']) and (not pd.isna(x[col+'_old']) and x[col+'_old'] != '-')
                else x[col+'_admit_old'], axis=1)
    if lvl >= 1 and not PAID_ONLY:
        for col in ['kcp', 'work']:
            df_old.loc[:, col+'_admit_old'] = df_old.apply(
                lambda x: 0 if pd.isna(x[col+'_admit_old']) and (not pd.isna(x[col+'_old']) and x[col+'_old'] != '-')
                else x[col + '_admit_old'], axis=1)

    if not info:
        columns_names = sorted(list(set(columns_names).intersection(col_correspondence.index)),
                               key=lambda x: col_correspondence.loc[x])
        columns = col_correspondence[columns_names[1:]]
        write_programs(wb, ws, join_new_old_progs(programs, df_old, lvl), columns, columns_names[1:])

    return days_till_end, cur_days_till_end, dates[masks[bool(lvl)]][idx], old_tot_inc_students
