import pandas as pd
import numpy as np
import xlsxwriter
from aux_program_scripts.xlsxtools import process_day
from aux_program_scripts import write_programs, RowWriter, create_format, join_new_old_progs
from green_wave.document_parts import create_header, list_programs, write_aicpk_numbers
from datetime import datetime, date, time
import os

CUR_DATE = datetime.today()
CREATION_TIME = time(hour=16, minute=0)
campuses = ['Москва', 'Санкт-Петербург', 'Пермь', 'Нижний Новгород']
CAMPUS = campuses[3]

programs_adm = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/tables_2024/bac/' + f'{CAMPUS}/' +
                             f'{CUR_DATE.strftime('%Y-%m-%d')}/bac_adm.xlsx')
creation_date = programs_adm.iloc[2, 0]
second_idx = [idx for idx in range(len(creation_date)) if creation_date[idx] == ' ']
creation_date = datetime.strptime(creation_date[second_idx[-2] + 1:], '%d.%m.%Y %H:%M:%S')

try:
    os.mkdir(f"/Users/s/Desktop/GW2024/{CUR_DATE.strftime('%Y-%m-%d')}")
except FileExistsError:
    pass

programs = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/programs_bac.xlsx', sheet_name='program')
programs = programs[(programs['campus'] == CAMPUS) & (~programs['terminated'])]

wb = xlsxwriter.Workbook(f"/Users/s/Desktop/GW2024/{CUR_DATE.strftime('%Y-%m-%d')}/GW.xlsx")
ws = wb.add_worksheet('2024')

# write a header

num_col, col_correspondence = create_header(wb, ws, CAMPUS)

# write index numbers and names

num_row = list_programs(wb, ws, programs)

# drawing green line separating free and paid places
for i in range(5, num_row):
    RowWriter(wb, ws, row=i, col=col_correspondence['green']).insert_green()

# writing a number of places at each program for this year

program_variables = ['kcp', 'work', 'special', 'separate']
cols = col_correspondence[program_variables]

write_programs(wb, ws, programs, cols, program_variables, programs_only=True)

if CAMPUS == campuses[1] or CAMPUS == campuses[2] or CAMPUS == campuses[3]:
    programs.loc[programs['name'] == 'Многопрофильный конкурс "Международный бакалавриат по бизнесу и экономике"', 'name'] = 'Бакалаврская программа «Международный бакалавриат по бизнесу и экономике»'
    if CAMPUS == campuses[2]:
        programs.loc[programs['name'] == 'Юриспруденция', 'name'] = 'Юриспруденция (направление подготовки 40.03.01 Юриспруденция)'
# writing admission numbers

total_inc_students = write_aicpk_numbers(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE)

# reading both spisok.xlsx and spisok_all.xlsx to correct (О Б) и (О Б ЦП) priorities

spisok_all = pd.read_excel('/Users/s/Desktop/GW2024/spisok_all.xlsx')
spisok = pd.read_excel('/Users/s/Desktop/GW2024/spisok.xlsx')

cur_case = None
priors, indices = [[], []], [[], []]
for sp in [spisok, spisok_all]:
    for _, row in sp.iterrows():
        if cur_case != row['Личное дело']:
            cur_case = row['Личное дело']
            newpriors = [np.argsort(prior) for prior in priors]

            for newprior, index in zip(newpriors, indices):
                for newpriority, i in zip(newprior, index):
                    sp.loc[i, 'Приоритет'] = newpriority + 1

            priors, indices = [[], []], [[], []]

        if row['Конкурсная группа'].find('(О Б)') != -1:
            priors[0].append(row['Приоритет'])
            indices[0].append(_)
        elif row['Конкурсная группа'].find('(О Б ЦП)') != -1:
            priors[1].append(row['Приоритет'])
            indices[1].append(_)

spisok[['Абитуриент', 'Приоритет', 'Конкурсная группа']].to_csv('/Users/s/Desktop/GW2024/priority.csv', index=False,
                                                                encoding='utf-8-sig')
spisok_all[['Абитуриент', 'Приоритет', 'Конкурсная группа']].to_csv('/Users/s/Desktop/GW2024/priority_all.csv',
                                                                    index=False, encoding='utf-8-sig')

# computing the number of work applications


if CAMPUS == 'Москва':
    applications_rvr = pd.read_excel('/Users/s/Desktop/GW2024/заявки/список.xlsx')
    applications_rvr = applications_rvr[applications_rvr['Статус заявки'] != 'Отозвана гражданином'].reset_index(
        drop=True)
    all_quotes = pd.read_excel('/Users/s/Desktop/GW2024/заявки/предложения.xlsx', sheet_name='Бак_спец (по квоте)')
    all_quotes.columns = list(all_quotes.iloc[8, :])
    all_quotes = all_quotes.iloc[9:, :].reset_index(drop=True)
    all_quotes = all_quotes[all_quotes['Кампус'] == "Москва"]
    work_apps = spisok
    print(work_apps)
    work_apps = work_apps[(work_apps['Конкурсная группа']).str.contains('О Б ЦП') & (work_apps['Приоритет'] == 1) &
                          (work_apps['Заявление забрано'] == 'Нет')]
    work_apps['Образовательная программа'] = work_apps['Конкурсная группа'].apply(lambda x: x[:x.find(' (')])
    print(work_apps)

    all_quotes['Номер предложения'] = all_quotes['№ Предложения'].astype(np.int64)
    applications_rvr = applications_rvr.set_index('Номер предложения').join(all_quotes.set_index('Номер предложения'))[
        ['Направление/Специальность', 'Образовательная программа', 'Гражданин']].reset_index()

    closed = pd.read_excel('/Users/s/Desktop/GW2024/заявки/закрытые.xlsx')[
        ['ФИО', "Направление/Специальность", 'Образовательная программа']]
    closed_num_of_abitur = closed.groupby(["Направление/Специальность", 'Образовательная программа']).count()[['ФИО']]

    work_apps['abitur'] = work_apps['Абитуриент'].apply(lambda x: x[:x.find(' ') + 2].lower())
    applications_rvr['abitur'] = applications_rvr['Гражданин'].apply(lambda x: x[:-1].lower() if len(x.split(' '))<3 else x[:x.find(' ') + 2].lower())
    work_apps = (work_apps.set_index(['abitur', 'Образовательная программа'])
                 .join(applications_rvr.set_index(['abitur', "Образовательная программа"]),
                       rsuffix='_rvr').reset_index())

    work_apps = work_apps[~pd.isna(work_apps['Гражданин'])]
    work_apps['Номер предложения'] = work_apps['Номер предложения'].astype(np.int64)
    all_quotes['Номер предложения'] = all_quotes['№ Предложения'].astype(np.int64)
    num_of_admits = work_apps.groupby('Номер предложения')[['abitur']].nunique().join(
        all_quotes.set_index('Номер предложения'), how='right', rsuffix='_rvr').reset_index()
    num_of_admits.loc[:, 'abitur'].fillna(0, inplace=True)
    num_of_admits['total_admits'] = np.min(num_of_admits[['abitur', 'Планируемое число договоров']], axis=1)

    programs_work = num_of_admits.groupby(['Направление/Специальность', 'Образовательная программа'])[
        ['Планируемое число договоров', 'abitur', 'total_admits']].sum()
    programs_work = programs_work.join(closed_num_of_abitur, how='outer')[['total_admits', 'ФИО']].fillna(0)

    programs_work = np.sum(programs_work, axis=1).reset_index()[
        ['Направление/Специальность', 'Образовательная программа', 0]]

    for i, row in programs_work.iterrows():
        if 'Реклама' in row['Направление/Специальность'] and 'Реклама' in row['Образовательная программа']:
            programs_work.loc[i, 'Образовательная программа'] = 'Реклама и связи с общественностью (направление подготовки 42.03.01 Реклама и связи с общественностью)'
        if 'Медиакоммуникации' in row['Направление/Специальность'] and 'Реклама' in row['Образовательная программа']:
            programs_work.loc[i, 'Образовательная программа'] = 'Реклама и связи с общественностью (направление подготовки 42.03.05 Медиакоммуникации)'

    programs_work.to_csv('/Users/s/Desktop/GW2024/заявки/сводная.csv', encoding='utf-8-sig')

# reading the rankings
rankings = pd.read_excel(f'/Users/s/Desktop/GW2024/konkurs_{CAMPUS}.xlsx').iloc[4:, :150]

konkurs = {'name': [], 'fio': [], 'score': [], 'original': []}

cur_name = None
idx_fio, idx_score, idx_original, idx_bvi = 0, 0, 0, 0
idx_name = 0
on_konkurs, first_skip = False, False
for idx, row in rankings.iterrows():
    if first_skip:
        first_skip = False
        continue

    if row.iloc[0] == 'Конкурсная группа':
        if idx_name == 0:
            for i in range(1, len(row)):
                if not pd.isna(row.iloc[i]):
                    idx_name = i
                    break

        cur_name = row.iloc[idx_name]
        continue

    if row.iloc[0] == '№ п/п':
        for i in range(len(row)):
            if row.iloc[i] == 'Подлинник документа об обр':
                idx_original = i
            elif row.iloc[i] == 'Абитуриент':
                idx_fio = i
            elif row.iloc[i] == 'Без экзаменов':
                idx_bvi = i
            elif row.iloc[i] == 'Сумма баллов':
                idx_score = i

        on_konkurs = True
        first_skip = True
        continue

    if on_konkurs:
        if pd.isna(row.iloc[0]):
            on_konkurs = False
            continue

        if pd.isna(row.iloc[idx_bvi]):
            konkurs['name'].append(cur_name)
            konkurs['fio'].append(row.iloc[idx_fio])
            konkurs['score'].append(row.iloc[idx_score])
            konkurs['original'].append((lambda x: False if pd.isna(x) or x == 'Нет' else True)(row.iloc[idx_original]))

konkurs = pd.DataFrame(konkurs)


def rename_konkurs(name, CAMPUS):
    for place in ['(О Б)', '(О Б ЦП)', '(О Б Отд)', '(О Б ОП)', '(О К)', '(О Б СК)', '(О Б Отд-СК)']:
        if CAMPUS == 'Москва':
            if 'Реклама и связи с общественностью ' + place + ' (Медиа)' in name:
                return 'Реклама и связи с общественностью (направление подготовки 42.03.05 Медиакоммуникации)'
            if 'Реклама и связи с общественностью ' + place in name:
                return 'Реклама и связи с общественностью (направление подготовки 42.03.01 Реклама и связи с общественностью)'
            if 'Античность ' + place + ' История' in name:
                return 'Античность (направление подготовки 46.03.01 История)'
            if 'Античность ' + place + ' Филология' in name:
                return 'Античность (направление подготовки 45.03.01 Филология)'

        if CAMPUS == campuses[2]:
            if 'Разработка информационных систем для бизнеса (Бизнес-информатика) ' + place in name:
                return 'Разработка информационных систем для бизнеса (направление подготовки 38.03.05 Бизнес-информатика)'
            if 'Разработка информационных систем для бизнеса (Программная инженерия) ' + place in name:
                return 'Разработка информационных систем для бизнеса (направление подготовки 09.03.04 Программная инженерия)'
            if 'Международный бакалавриат по бизнесу и экономике ' + place in name:
                return 'Бакалаврская программа «Международный бакалавриат по бизнесу и экономике»'
            if 'Юриспруденция ' + place in name:
                return 'Юриспруденция (направление подготовки 40.03.01 Юриспруденция)'

        if CAMPUS == campuses[3]:
            if 'Компьютерные науки и технологии БИ ' + place in name:
                return 'Компьютерные науки и технологии (направление подготовки 38.03.05 Бизнес-информатика)'
            if 'Компьютерные науки и технологии ПМИ ' + place in name:
                return 'Компьютерные науки и технологии (направление подготовки 01.03.02 Прикладная математика и информатика)'
            if 'Компьютерные науки и технологии ПИ ' + place in name:
                return 'Компьютерные науки и технологии (направление подготовки 09.03.04 Программная инженерия)'
            if 'Международный бакалавриат по бизнесу и экономике ' + place in name:
                return 'Бакалаврская программа «Международный бакалавриат по бизнесу и экономике»'

    return name[:name.find(' (')]


#spisok_all = pd.read_excel('/Users/s/Desktop/GW2024/spisok_all.xlsx')[['Абитуриент', 'Приоритет', 'Конкурсная группа']]
spisok_all = spisok_all[['Абитуриент', 'Заявление забрано', 'Приоритет', 'Конкурсная группа']]
spisok_all.columns = ['fio', 'withdraw', 'priority', 'name']
spisok_all = spisok_all[spisok_all['withdraw'] == 'Нет']
konkurs = konkurs.set_index(['fio', 'name']).join(spisok_all.set_index(['fio', "name"]), how='inner').reset_index()
konkurs.loc[:, 'name'] = konkurs['name'].apply(rename_konkurs, args=[CAMPUS])

if CAMPUS == 'Москва':
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final.xlsx')
elif CAMPUS == 'Санкт-Петербург':
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_spb.xlsx')
elif CAMPUS == campuses[2]:
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_perm.xlsx')
elif CAMPUS == campuses[3]:
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_nn.xlsx')

if CAMPUS == campuses[0]:
    final = final.iloc[4:72, [1, 3, 4, 5, 6, 20, 21, 22, 25, 26, 29, 30, 31, 32]]
    final.columns = ['name', 'work', 'work_max', 'special', 'separate', 'green_num', 'green_num_old', 'realized_min_score',
                     'yellow_num', 'yellow_num_old', 'realized_min_score_old2022', 'green_num_old2022',
                     'realized_min_score_old2021', 'green_num_old2021']
else:
    final = final.iloc[4:72, [1, 3, 4, 5, 6, 20, 21, 24, 25, 28, 29, 30, 31]]
    final.columns = ['name', 'work', 'work_max', 'special', 'separate', 'green_num', 'green_num_old',
                     'yellow_num', 'yellow_num_old', 'realized_min_score_old2022', 'green_num_old2022',
                     'realized_min_score_old2021', 'green_num_old2021']



konkurs = konkurs.set_index('name').join(final.set_index('name')).reset_index()
konkurs['green_admit'] = konkurs.apply(lambda x: 1 if x['green_num'] != '-' and x['score'] >= x['green_num'] else 0, axis=1)
konkurs['yellow_admit'] = konkurs.apply(lambda x: 1 if x['yellow_num'] != '-' and x['score'] >= x['yellow_num'] else 0, axis=1)

allpriority_konkurs = konkurs.groupby(['fio', 'name'])[['green_admit', 'yellow_admit']].max().reset_index()
allpriority_konkurs = allpriority_konkurs.pivot_table(values=['green_admit', 'yellow_admit'], index=['name'], aggfunc='sum').reset_index()
allpriority_konkurs.columns = ['name', 'ingreen_admit', 'inyellow_admit']
allpriority_konkurs.loc[:, 'inyellow_admit'] = allpriority_konkurs.loc[:, 'inyellow_admit'] - allpriority_konkurs.loc[:, 'ingreen_admit']
priority1_konkurs = (konkurs[(konkurs['priority'] == 1) & (konkurs['original'])].
                     pivot_table(values=['green_admit', 'yellow_admit'], index=['name'], aggfunc='sum')).reset_index()
priority1_konkurs.columns = ['name', 'ingreen_admit_priority', 'inyellow_admit_priority']
priority1_konkurs.loc[:, 'inyellow_admit_priority'] = priority1_konkurs.loc[:, 'inyellow_admit_priority'] - priority1_konkurs.loc[:, 'ingreen_admit_priority']
allpriority_konkurs.loc[:, 'inyellow_admit'] = allpriority_konkurs.apply(lambda x: '-' if final.set_index('name').loc[x['name'], 'yellow_num'] == '-' else x['inyellow_admit'], axis=1)
priority1_konkurs.loc[:, 'inyellow_admit_priority'] = priority1_konkurs.apply(lambda x: '-' if final.set_index('name').loc[x['name'], 'yellow_num'] == '-' else x['inyellow_admit_priority'], axis=1)

pd.DataFrame(konkurs).to_csv(f'/Users/s/Desktop/GW2024/konkurs_{CAMPUS}.csv', index=False, encoding='utf-8-sig')

# preparing for pivots

spisok['Платно'] = spisok['Конкурсная группа'].apply(lambda x: 1 if x.find('(О К)') != -1 else 0)
spisok['Целевое'] = spisok['Конкурсная группа'].apply(lambda x: 1 if x.find('(О Б ЦП)') != -1 else 0)
spisok['Отдельная'] = spisok['Конкурсная группа'].apply(
    lambda x: 1 if x.find('(О Б Отд)') != -1 or x.find('(О Б СК)') != -1 or x.find('(О Б Отд-СК)') != -1 else 0)
got_otdelnaya = spisok.groupby('Личное дело')[['Отдельная']].max()
spisok['Особая'] = spisok.apply(
    lambda x: 1 if x['Конкурсная группа'].find('(О Б ОП)') != -1 and got_otdelnaya.loc[
        x['Личное дело'], 'Отдельная'] == 0 else 0, axis=1)
spisok['Образовательная программа'] = spisok['Конкурсная группа'].apply(rename_konkurs, args=[CAMPUS])
print(spisok['Образовательная программа'].unique())
# creating pivots

origs = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) & (spisok['Заявление забрано'] == 'Нет')]
         .pivot_table(values=['Код'], index=['Образовательная программа'], aggfunc='count')).reset_index()
origs.columns = ['name', 'original_admit']

prior1 = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
                 (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
          .pivot_table(values=['Код'], index=['Образовательная программа'], aggfunc='count')).reset_index()
prior1.columns = ['name', 'original_admit_priority']

bvi = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
              (spisok['Заявление забрано'] == 'Нет') & (spisok['Без ВИ'] == 'Да')]
       .pivot_table(values=['Код'], index=['Образовательная программа'], aggfunc='count')).reset_index()
bvi.columns = ['name', 'bvi_admitted']
print(programs['name'].unique(), bvi.loc[2, 'name'])
if CAMPUS != campuses[0]:
    work = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
                  (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
           .pivot_table(values=['Целевое'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
    work.columns = ['name', 'work_admitted']

    if CAMPUS == campuses[3]:
        work.loc[work['name'] == 'Компьютерные науки и технологии (направление подготовки 09.03.04 Программная инженерия)', 'work_admitted'] += 1
        work.loc[work['name'] == 'Международный бакалавриат по бизнесу и экономике', 'work_admitted'] += 2
    if CAMPUS == campuses[2]:
        work.loc[work['name'] == 'Разработка информационных систем для бизнеса (направление подготовки 09.03.04 Программная инженерия)', 'work_admitted'] += 1
        work.loc[work['name'] == 'Разработка информационных систем для бизнеса (направление подготовки 38.03.05 Бизнес-информатика)', 'work_admitted'] += 2
    if CAMPUS == campuses[1]:
        work.loc[work['name'] == 'Аналитика в экономике', 'work_admitted'] += 1
else:
    work = programs_work[['Образовательная программа', 0]]
    work.columns = ['name', 'work_admitted']
print(work.iloc[10:20], final.loc[10:20, ['work', 'work_max']])
work.loc[:, 'work_admitted'] = work.apply(lambda x: np.nan if final.set_index('name').loc[x['name'], 'work'] == '-' else min(x['work_admitted'], final.set_index('name').loc[x['name'], 'work'], final.set_index('name').loc[x['name'], 'work_max']), axis=1)

special = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
              (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
       .pivot_table(values=['Особая'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
special.columns = ['name', 'special_admitted']
special.loc[:, 'special_admitted'] = special.apply(lambda x: np.nan if final.set_index('name').loc[x['name'], 'special'] == '-' else min(x['special_admitted'], final.set_index('name').loc[x['name'], 'special']), axis=1)

separate = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
              (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
       .pivot_table(values=['Отдельная'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
separate.columns = ['name', 'separate_admitted']
separate.loc[:, 'separate_admitted'] = separate.apply(lambda x: np.nan if final.set_index('name').loc[x['name'], 'separate'] == '-' else min(x['separate_admitted'], final.set_index('name').loc[x['name'], 'separate']), axis=1)

pivots = [origs, prior1, bvi, work, special, separate, allpriority_konkurs, priority1_konkurs, final]
pivot_variables = [['original_admit'], ['original_admit_priority'], ['bvi_admitted'],
                   ['work_admitted'], ['special_admitted'], ['separate_admitted'],
                   ['ingreen_admit', 'inyellow_admit'], ['ingreen_admit_priority', 'inyellow_admit_priority'],
                   list(final.columns)[1:]]

for pivot, prog_var in zip(pivots, pivot_variables):
    cols = col_correspondence[prog_var]
    pivot = programs.set_index('name').join(pivot.set_index('name'), rsuffix='_pivot').reset_index()
    print(pivot)
    pivot.loc[pivot['kcp'] == '-'] = pivot[pivot['kcp'] == '-'].fillna('-')
    pivot.loc[pivot['kcp'] != '-'] = pivot[pivot['kcp'] != '-'].fillna(0)
    write_programs(wb, ws, pivot, cols, prog_var, programs_only=True, old=False)

ws.write(0, 1, f"Предложения по баллам \"зеленой волны\": ПК-2024 24 июля 2024 г. ({CAMPUS})",
         create_format(wb, bold=True, font_size=14, text_wrap=False, halign='left', border=False))
ws.write(1, 1, f"Данные на {creation_date.strftime('%H:%M %d-%m-%Y')}",
         create_format(wb, text_wrap=False, font_size=12, italic=True, halign='left', border=False))

ws.autofilter(4, 0, num_row, num_col)
ws.freeze_panes(5, 1)
ws.set_landscape()
ws.repeat_rows(first_row=3, last_row=4)
ws.set_column(0, 0, 3)
ws.set_column(1, 1, 46.33)
ws.set_column(2, 2, 12.67)
ws.set_column(3, col_correspondence['green'] - 1, 6)
ws.set_column(col_correspondence['green'], col_correspondence['green'], 2)
ws.set_column(col_correspondence['green'] + 1, 40, 6)
ws.set_paper(8)
ws.set_margins(top=0.4, bottom=0.4, left=0.1, right=0.1)
ws.fit_to_pages(width=1, height=0)
ws.set_footer('&RСтраница &P из &N')
ws.set_row(4, 146)
ws.set_row(3, 52)

wb.close()
