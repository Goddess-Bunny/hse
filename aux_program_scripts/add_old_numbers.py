import pandas as pd
import re

MAG = True

if MAG:
    last_year = pd.read_excel(r'/Users/s/Desktop/Admission Numbers 2024/last_year_mag.xlsx')
    programs = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/programs.xlsx', sheet_name='program')
else:
    last_year = pd.read_excel(r'/Users/s/Desktop/Admission Numbers 2024/last_year_bac.xlsx')
    programs = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/programs_bac.xlsx', sheet_name='program')

programs = programs[programs['campus'] == 'Москва']

if MAG:
    last_year = last_year.iloc[7:193, [0, 1, 2, 5, 12]]
    last_year.columns = ['name', 'kcp', 'school', 'work', 'paid']
else:
    last_year = last_year.iloc[7:112, [0, 1, 3, 5, 7, 22]]
    last_year.columns = ['name', 'kcp', 'work', 'special', 'separate', 'paid']

last_year = last_year[~(last_year['name'].str.contains('очно', case=False) |
                        last_year['name'].str.contains(r'\d{2}\.\d{2}\.\d{2}', regex=True))]
print(last_year)
programs['old_name'] = programs.apply(lambda x:
                                      x['rename'] if x['rename'] != '-' else x['name'] if not x['new'] else '-', axis=1)
programs['old_name'] = programs['old_name'].map(lambda x: x[:x.find(' /')] if x.find(' /') != -1 else x)
programs2 = programs[programs['old_name'] != '-']

if MAG:
    joined = programs2.set_index('old_name').join(last_year.set_index('name'), rsuffix='_old')[
        ['kcp_old', 'work_old', 'school_old', 'paid_old', 'id']].reset_index().set_index('id')
else:
    joined = programs2.set_index('old_name').join(last_year.set_index('name'), rsuffix='_old')[
        ['kcp_old', 'work_old', 'school_old', 'special_old', 'separate_old', 'paid_old', 'id']].reset_index().set_index(
        'id')
#joined.fillna(value='-', inplace=True)

programs['kcp_old'] = programs['id'].apply(lambda x: joined.loc[x, 'kcp_old'] if x in joined.reset_index()['id'].unique() else '-')
programs['school_old'] = programs['id'].apply(lambda x: joined.loc[x, 'school_old'] if x in joined.reset_index()['id'].unique() else '-')
programs['work_old'] = programs['id'].apply(lambda x: joined.loc[x, 'work_old'] if x in joined.reset_index()['id'].unique() else '-')
if not MAG:
    programs['special_old'] = programs['id'].apply(lambda x: joined.loc[x, 'special_old'] if x in joined.reset_index()['id'].unique() else '-')
    programs['separate_old'] = programs['id'].apply(lambda x: joined.loc[x, 'separate_old'] if x in joined.reset_index()['id'].unique() else '-')
programs['paid_old'] = programs['id'].apply(lambda x: joined.loc[x, 'paid_old'] if x in joined.reset_index()['id'].unique() else '-')

programs.to_excel('/Users/s/Desktop/Admission Numbers 2024/programs_old.xlsx')