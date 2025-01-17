import pandas as pd
import numpy as np
import xlsxwriter
from aux_program_scripts.xlsxtools import process_day
from aux_program_scripts import write_programs, RowWriter, create_format, join_new_old_progs
from green_wave.document_parts import create_header_original, list_programs, write_aicpk_numbers
from datetime import datetime, date, time
import os

CUR_DATE = datetime.today()
CREATION_TIME = time(hour=16, minute=0)
campuses = ['Москва', 'Санкт-Петербург', 'Пермь', 'Нижний Новгород']
CAMPUS = campuses[0]

programsPath = 'Users/s/Desktop/Admission Numbers 2024/'
gwPath = 'Users/s/Desktop/GW2024/'

#
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

gw_stat = pd.read_excel('/Users/s/Desktop/GW2024/gw_stat.xlsx').iloc[7:, :]
gw_stat_df = {'name': [], 'green_score': [], 'fio': [], 'high_prior': [], 'no_high_prior': [], 'original': [], 'green_admit': [],
              'bvi': [], 'score': []}

cur_name, green_score = None, None
for idx, row in gw_stat.iterrows():
    if row.iloc[0] == 'Нижний Новгород':
        break

    if '(О Б)' in row.iloc[0]:
        green_score = row.iloc[4]
        cur_name = rename_konkurs(row.iloc[0], CAMPUS)
        continue

    gw_stat_df['green_score'].append(row.iloc[5])
    gw_stat_df['name'].append(rename_konkurs(row.iloc[4], CAMPUS))
    gw_stat_df['fio'].append(row.iloc[0])
    gw_stat_df['high_prior'].append(True if row.iloc[6] == 'Да' else False)
    gw_stat_df['no_high_prior'].append(True if row.iloc[7] == 'Да' else False)
    gw_stat_df['original'].append(True if row.iloc[8] == 'Да' else False)
    gw_stat_df['green_admit'].append(True if row.iloc[9] == 'Да' else False)
    gw_stat_df['bvi'].append(True if row.iloc[10] == 'Да' else False)
    gw_stat_df['score'].append(row.iloc[11] if not pd.isna(row.iloc[11]) else 0)

gw_stat_df = pd.DataFrame(gw_stat_df)
gw_stat_df.to_csv('/Users/s/Desktop/GW2024/gw_stat.csv', index=False, encoding='utf-8-sig')

admission_score = gw_stat_df[~gw_stat_df['bvi'] & (gw_stat_df['score'] > 0) & gw_stat_df['original'] & gw_stat_df['high_prior']].groupby('name')['score'].min().reset_index()
admission_score.columns = ['name', 'admit_score']
green_score = gw_stat_df.groupby(['name', 'green_score'])['bvi'].max().reset_index().set_index('name')
admission_score['admit_score'] = admission_score.apply(lambda x: x['admit_score'] if not pd.isna(x['admit_score']) else green_score.loc[x['name']], axis=1)

admitted_green = gw_stat_df[gw_stat_df['green_admit'] & gw_stat_df['original'] & gw_stat_df['high_prior']].groupby('name')[['no_high_prior']].count().reset_index()
admitted_other = gw_stat_df[gw_stat_df['original'] & gw_stat_df['no_high_prior'] & ~gw_stat_df['high_prior'] & ~gw_stat_df['bvi'] & gw_stat_df['green_admit']]
admitted_other_bvi = gw_stat_df[gw_stat_df['original'] & gw_stat_df['no_high_prior'] & gw_stat_df['bvi']]

print(admission_score)
programs_adm = pd.read_excel(programsPath + 'tables_2024/bac/' + f'{CAMPUS}/' +
                             f'{CUR_DATE.strftime('%Y-%m-%d')}/bac_adm.xlsx')
creation_date = programs_adm.iloc[2, 0]
second_idx = [idx for idx in range(len(creation_date)) if creation_date[idx] == ' ']
creation_date = datetime.strptime(creation_date[second_idx[-2] + 1:], '%d.%m.%Y %H:%M:%S')

try:
    os.mkdir(f"/Users/s/Desktop/GW2024/{CUR_DATE.strftime('%Y-%m-%d')}")
except FileExistsError:
    pass

programs = pd.read_excel(programsPath + 'programs_bac.xlsx', sheet_name='program')
programs = programs[(programs['campus'] == CAMPUS) & (~programs['terminated'])]

wb = xlsxwriter.Workbook(gwPath + "{CUR_DATE.strftime('%Y-%m-%d')}/original.xlsx")
ws = wb.add_worksheet('2024')

# write a header

num_col, col_correspondence = create_header_original(wb, ws, CAMPUS)

# write index numbers and names

num_row = list_programs(wb, ws, programs)

# writing a number of places at each program for this year

program_variables = ['kcp', 'work', 'special', 'separate']
cols = col_correspondence[program_variables]

write_programs(wb, ws, programs, cols, program_variables, programs_only=True)

if CAMPUS == campuses[1] or CAMPUS == campuses[2] or CAMPUS == campuses[3]:
    programs.loc[programs[
                     'name'] == 'Многопрофильный конкурс "Международный бакалавриат по бизнесу и экономике"', 'name'] = 'Бакалаврская программа «Международный бакалавриат по бизнесу и экономике»'
    if CAMPUS == campuses[2]:
        programs.loc[programs[
                         'name'] == 'Юриспруденция', 'name'] = 'Юриспруденция (направление подготовки 40.03.01 Юриспруденция)'
# writing admission numbers

spisok_all = pd.read_excel(gwPath + 'spisok_all.xlsx', sheet_name='TDSheet')
spisok = pd.read_excel(gwPath + 'spisok.xlsx', sheet_name='TDSheet')

cur_case = None
priors, indices = [[], []], [[], []]
for sp in [spisok, spisok_all]:
    for _, row in sp.iterrows():
        if row['Заявление забрано'] == 'Да':
            continue

        if cur_case != row['Личное дело']:
            cur_case = row['Личное дело']
            newpriors = [np.argsort(prior) for prior in priors]

            for newprior, index in zip(newpriors, indices):
                for i in range(len(index)):
                    for j in range(len(newprior)):
                        if i == newprior[j]:
                            sp.loc[index[i], 'Приоритет'] = j + 1

            priors, indices = [[], []], [[], []]

        if row['Конкурсная группа'].find('(О Б)') != -1:
            priors[0].append(row['Приоритет'])
            indices[0].append(_)
        elif row['Конкурсная группа'].find('(О Б ЦП)') != -1:
            priors[1].append(row['Приоритет'])
            indices[1].append(_)

# # reading the work quota
#
# applications_rvr = pd.read_excel('/Users/s/Desktop/GW2024/заявки/список.xlsx')
# applications_rvr = applications_rvr[applications_rvr['Статус заявки'] != 'Отозвана гражданином'].reset_index(
#     drop=True)
# all_quotes = pd.read_excel('/Users/s/Desktop/GW2024/заявки/предложения.xlsx', sheet_name='Бак_спец (по квоте)')
# all_quotes.columns = list(all_quotes.iloc[8, :])
# all_quotes = all_quotes.iloc[9:, :].reset_index(drop=True)
# all_quotes = all_quotes[all_quotes['Кампус'] == "Москва"]
# work_apps = spisok
#
# work_apps = work_apps[(work_apps['Конкурсная группа']).str.contains('О Б ЦП') & (work_apps['Приоритет'] == 1) &
#                       (work_apps['Заявление забрано'] == 'Нет')]
# work_apps['Образовательная программа'] = work_apps['Конкурсная группа'].apply(lambda x: x[:x.find(' (')])
#
# all_quotes['Номер предложения'] = all_quotes['№ Предложения'].astype(np.int64)
# applications_rvr = applications_rvr.set_index('Номер предложения').join(all_quotes.set_index('Номер предложения'))[
#     ['Направление/Специальность', 'Образовательная программа', 'Гражданин']].reset_index()
#
# closed = pd.read_excel('/Users/s/Desktop/GW2024/заявки/закрытые.xlsx')[
#     ['ФИО', "Направление/Специальность", 'Образовательная программа']]
# closed_num_of_abitur = closed.groupby(["Направление/Специальность", 'Образовательная программа']).count()[['ФИО']]
#
# work_apps['abitur'] = work_apps['Абитуриент'].apply(lambda x: x[:x.find(' ') + 2].lower())
# applications_rvr['abitur'] = applications_rvr['Гражданин'].apply(lambda x: x[:-1].lower() if len(x.split(' '))<3 else x[:x.find(' ') + 2].lower())
# work_apps = (work_apps.set_index(['abitur', 'Образовательная программа'])
#              .join(applications_rvr.set_index(['abitur', "Образовательная программа"]),
#                    rsuffix='_rvr').reset_index())
#
# work_apps = work_apps[~pd.isna(work_apps['Гражданин'])]
# work_apps['Номер предложения'] = work_apps['Номер предложения'].astype(np.int64)
# all_quotes['Номер предложения'] = all_quotes['№ Предложения'].astype(np.int64)
# num_of_admits = work_apps.groupby('Номер предложения')[['abitur']].nunique().join(
#     all_quotes.set_index('Номер предложения'), how='right', rsuffix='_rvr').reset_index()
# num_of_admits.loc[:, 'abitur'].fillna(0, inplace=True)
# num_of_admits['total_admits'] = np.min(num_of_admits[['abitur', 'Планируемое число договоров']], axis=1)
#
# programs_work = num_of_admits.groupby(['Направление/Специальность', 'Образовательная программа'])[
#     ['Планируемое число договоров', 'abitur', 'total_admits']].sum()
# programs_work = programs_work.join(closed_num_of_abitur, how='outer')[['total_admits', 'ФИО']].fillna(0)
#
# programs_work = np.sum(programs_work, axis=1).reset_index()[
#     ['Направление/Специальность', 'Образовательная программа', 0]]
#
# for i, row in programs_work.iterrows():
#     if 'Реклама' in row['Направление/Специальность'] and 'Реклама' in row['Образовательная программа']:
#         programs_work.loc[i, 'Образовательная программа'] = 'Реклама и связи с общественностью (направление подготовки 42.03.01 Реклама и связи с общественностью)'
#     if 'Медиакоммуникации' in row['Направление/Специальность'] and 'Реклама' in row['Образовательная программа']:
#         programs_work.loc[i, 'Образовательная программа'] = 'Реклама и связи с общественностью (направление подготовки 42.03.05 Медиакоммуникации)'

# reading the rankings
rankings = pd.read_excel(gwPath + 'konkurs_{CAMPUS}.xlsx').iloc[4:, :]

konkurs = {'name': [], 'fio': [], 'score': [], 'original': [], 'bvi': []}

cur_name = None
idx_fio, idx_score, idx_original, idx_bvi = 0, 0, 0, 0
idx_name = 0
idx_exact_scores = {}
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
        found_score = False
        for i in range(len(row)):
            if row.iloc[i] == 'Подлинник документа об обр':
                idx_original = i
            elif row.iloc[i] == 'Абитуриент':
                idx_fio = i
            elif row.iloc[i] == 'Без экзаменов':
                idx_bvi = i
            elif row.iloc[i] == 'Сумма баллов':
                idx_score = i
                found_score = True
            elif found_score and not pd.isna(row.iloc[i]) and len(row.iloc[i]) > 0:
                idx_exact_scores[row.iloc[i]] = i

        on_konkurs = True
        first_skip = True
        continue

    if on_konkurs:
        if pd.isna(row.iloc[0]):
            on_konkurs = False
            idx_scores = {}
            continue

        konkurs['name'].append(cur_name)
        konkurs['fio'].append(row.iloc[idx_fio])
        konkurs['score'].append(row.iloc[idx_score])
        konkurs['original'].append((lambda x: False if pd.isna(x) or x == 'Нет' else True)(row.iloc[idx_original]))

        if pd.isna(row.iloc[idx_bvi]):
            konkurs['bvi'].append(0)

            for exam in idx_exact_scores.keys():
                if not exam in konkurs.keys():
                    konkurs[exam] = [200] * (len(konkurs['name']) - 1) + [row.iloc[idx_exact_scores[exam]]]
                else:
                    konkurs[exam].append(row.iloc[idx_exact_scores[exam]])

            cur_len = len(konkurs['name'])
            for exam in konkurs.keys():
                if exam not in ['name', 'fio', 'score', 'original', 'bvi']:
                    if len(konkurs[exam]) < cur_len:
                        konkurs[exam].append(200)
        else:
            konkurs['bvi'].append(1)
            cur_len = len(konkurs['name'])
            for exam in konkurs.keys():
                if exam not in ['name', 'fio', 'score', 'original', 'bvi']:
                    if len(konkurs[exam]) < cur_len:
                        konkurs[exam].append(200)

for key in konkurs.keys():
    print(key)
konkurs = pd.DataFrame(konkurs).fillna(0)



# # reading all admits
# admitted = pd.read_excel('/Users/s/Desktop/GW2024/admitted.xlsx')
# admitted = admitted[pd.isna(admitted['Приказ об исключении из приказа'])].groupby(['Регистрационный номер', 'СНИЛС'])[
#     ['№ п/п']].sum().reset_index()
# admitted = admitted.set_index('Регистрационный номер').join(spisok_all.set_index("Личное дело"))[['СНИЛС', 'Код', '№ п/п']].reset_index()
# admitted = admitted.groupby(['Регистрационный номер', 'СНИЛС', 'Код'])[['№ п/п']].sum().reset_index()
# print(admitted)
# real_admitted = pd.concat([pd.read_excel('/Users/s/Desktop/GW2024/real_admitted.xlsx', sheet_name=sheet)
#                            for sheet in ['БВИ fin', 'Целевая квота', 'Особая квота']])
# admitted1 = admitted.set_index('СНИЛС').join(real_admitted.set_index('СНИЛС'), how='inner', rsuffix='_a').reset_index()
# real_admitted = pd.read_excel('/Users/s/Desktop/GW2024/real_admitted.xlsx', sheet_name='Отдельная квота')
# admitted2 = admitted.set_index('Код').join(real_admitted.set_index('Уникальный идентификатор'), how='inner', rsuffix='_a').reset_index()
#
# admitted = pd.concat([admitted1, admitted2])
# print(admitted)
#
# spisok_all['Образовательная программа'] = spisok_all['Конкурсная группа'].apply(rename_konkurs, args=[CAMPUS])
# spisok['Образовательная программа'] = spisok['Конкурсная группа'].apply(rename_konkurs, args=[CAMPUS])
# spisok_all['admitted'] = ~pd.isna(
#     spisok_all.set_index(['Личное дело', "Специальность", "Образовательная программа"]).join(
#         admitted.set_index(['Регистрационный номер', 'Направление', "Образовательная программа"]),
#         rsuffix='_adm').reset_index()['№ п/п'])
# spisok['admitted'] = ~pd.isna(
#     spisok.set_index(['Личное дело', "Специальность", "Образовательная программа"]).join(
#         admitted.set_index(['Регистрационный номер', 'Направление', "Образовательная программа"]),
#         rsuffix='_adm').reset_index()['№ п/п'])


#spisok_all['Поступил'] = spisok_all['Статус ССПВО'].apply(lambda x: 1 if x == 'Включен в приказ о зачислении' else 0)
spisok_all['Платно'] = spisok_all['Конкурсная группа'].apply(lambda x: 1 if x.find('(О К)') != -1 else 0)
spisok_all['Целевое'] = spisok_all['Конкурсная группа'].apply(lambda x: 1 if x.find('(О Б ЦП)') != -1 else 0)
spisok_all['Отдельная'] = spisok_all['Конкурсная группа'].apply(
    lambda x: 1 if x.find('(О Б Отд)') != -1 or x.find('(О Б СК)') != -1 or x.find('(О Б Отд-СК)') != -1 else 0)
got_otdelnaya = spisok_all.groupby('Личное дело')[['Отдельная']].max()
spisok_all['Особая'] = spisok_all.apply(
    lambda x: 1 if x['Конкурсная группа'].find('(О Б ОП)') != -1 and got_otdelnaya.loc[
        x['Личное дело'], 'Отдельная'] == 0 else 0, axis=1)

spisok_all['original'] = ~pd.isna(spisok_all.set_index(['Личное дело', 'Конкурсная группа']).join(
    spisok.set_index(['Личное дело', 'Конкурсная группа']), rsuffix='_orig').reset_index()['Абитуриент_orig'])
print(spisok_all['original'].sum())
#spisok_all = pd.read_excel('/Users/s/Desktop/GW2024/spisok_all.xlsx')[['Абитуриент', 'Приоритет', 'Конкурсная группа']]
spisok_all_short = spisok_all[
    ['Абитуриент', 'Личное дело', 'Заявление забрано', 'Приоритет', 'Конкурсная группа', 'Целевое',
     'Платно', 'Особая', 'Отдельная', 'Без ВИ', 'Филиал', 'original']]
spisok_all_short.columns = ['fio', 'id', 'withdraw', 'priority', 'name', 'work', 'paid', 'special',
                            'separate', 'bvi', 'campus', 'original']
spisok_all_short = spisok_all_short[(spisok_all_short['withdraw'] == 'Нет') & (spisok_all_short['campus'] == CAMPUS)]

konkurs_short = konkurs.copy()
konkurs = konkurs.set_index(['fio', 'name']).join(spisok_all_short.set_index(['fio', "name"]), how='inner',
                                                  rsuffix='_s').reset_index()
konkurs.loc[:, 'name'] = konkurs['name'].apply(rename_konkurs, args=[CAMPUS])

konkurs.to_csv(f'/Users/s/Desktop/GW2024/konkurs_{CAMPUS}.csv', index=False, encoding='utf-8-sig')

#spisok['Поступил'] = spisok['Статус ССПВО'].apply(lambda x: 1 if x == 'Включен в приказ о зачислении' else 0)
spisok['Платно'] = spisok['Конкурсная группа'].apply(lambda x: 1 if x.find('(О К)') != -1 else 0)
spisok['Целевое'] = spisok['Конкурсная группа'].apply(lambda x: 1 if x.find('(О Б ЦП)') != -1 else 0)
spisok['Отдельная'] = spisok['Конкурсная группа'].apply(
    lambda x: 1 if x.find('(О Б Отд)') != -1 or x.find('(О Б СК)') != -1 or x.find('(О Б Отд-СК)') != -1 else 0)
got_otdelnaya = spisok.groupby('Личное дело')[['Отдельная']].max()
spisok['Особая'] = spisok.apply(
    lambda x: 1 if x['Конкурсная группа'].find('(О Б ОП)') != -1 and got_otdelnaya.loc[
        x['Личное дело'], 'Отдельная'] == 0 else 0, axis=1)

spisok_short = spisok[
    ['Абитуриент', 'Личное дело', 'Заявление забрано', 'Приоритет', 'Конкурсная группа', 'Целевое',
     'Платно', 'Особая', 'Отдельная', 'Без ВИ', 'Филиал']]
spisok_short.columns = ['fio', 'id', 'withdraw', 'priority', 'name', 'work', 'paid', 'special', 'separate',
                        'bvi', 'campus']
spisok_short = spisok_short[(spisok_short['withdraw'] == 'Нет') & (spisok_short['campus'] == CAMPUS)]

konkurs_short = konkurs_short.set_index(['fio', 'name']).join(spisok_short.set_index(['fio', "name"]), how='inner',
                                                              rsuffix='_s').reset_index()
konkurs_short.loc[:, 'name'] = konkurs_short['name'].apply(rename_konkurs, args=[CAMPUS])

if CAMPUS == 'Москва':
    final = pd.read_excel(gwPath + 'final_original.xlsx')
elif CAMPUS == 'Санкт-Петербург':
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_spb.xlsx')
elif CAMPUS == campuses[2]:
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_perm.xlsx')
elif CAMPUS == campuses[3]:
    final = pd.read_excel(f'/Users/s/Desktop/GW2024/final_nn.xlsx')

exam = ['Русский язык', 'История', 'Английский язык', 'Испанский язык', 'Китайский язык', 'Немецкий язык',
        'Французский язык', 'Литература', 'Информатика и ИКТ', 'Математика', 'География', 'Обществознание',
        'Творческое испытание (дизайн)',
        'Дополнительное вступительное испытание творческой направленности (журналистика)', 'Физика',
        'Биология', 'Химия']
if CAMPUS == campuses[0]:
    final = final.iloc[4:72, [1, 3, 4, 5, 6, 20, 21, 22, 25, 26, 29, 30, 31, 32] + list(range(39, 56))]
    final.columns = ['name', 'work', 'work_max', 'special', 'separate', 'green_num', 'green_num_old',
                     'realized_min_score',
                     'yellow_num', 'yellow_num_old', 'realized_min_score_old2022', 'green_num_old2022',
                     'realized_min_score_old2021', 'green_num_old2021'] + exam

    final.loc[final['name'] == 'Актёр', 'name'] = 'Актер'
else:
    final = final.iloc[4:72, [1, 3, 4, 5, 6, 20, 21, 24, 25, 28, 29, 30, 31]]
    final.columns = ['name', 'work', 'work_max', 'special', 'separate', 'green_num', 'green_num_old',
                     'yellow_num', 'yellow_num_old', 'realized_min_score_old2022', 'green_num_old2022',
                     'realized_min_score_old2021', 'green_num_old2021']
final = final.fillna(0)

konkurs = konkurs.set_index('name').join(final.set_index('name'), rsuffix='_final').reset_index()
konkurs['green_admit'] = konkurs.apply(lambda x: 1 if x['green_num'] != '-' and x['score'] >= x['green_num'] else 0,
                                       axis=1)
konkurs['yellow_admit'] = konkurs.apply(lambda x: 1 if x['yellow_num'] != '-' and x['score'] >= x['yellow_num'] else 0,
                                        axis=1)

for ex in exam:
    konkurs[f'score_diff_{ex}'] = konkurs[ex] - konkurs[ex + '_final']
konkurs['min_admit'] = konkurs[[f'score_diff_{ex}' for ex in exam]].min(axis=1) >= 0

konkurs_short = konkurs_short.set_index('name').join(final.set_index('name'), rsuffix='_final').reset_index()
konkurs_short['green_admit'] = konkurs_short.apply(
    lambda x: 1 if x['green_num'] != '-' and x['score'] >= x['green_num'] else 0, axis=1)
konkurs_short['yellow_admit'] = konkurs_short.apply(
    lambda x: 1 if x['yellow_num'] != '-' and x['score'] >= x['yellow_num'] else 0, axis=1)

for ex in exam:
    konkurs_short[f'score_diff_{ex}'] = konkurs_short[ex] - konkurs_short[ex + '_final']
konkurs_short['min_admit'] = konkurs_short[[f'score_diff_{ex}' for ex in exam]].min(axis=1) >= 0

konkurs_short.to_csv(gwPath + f'konkurs_{CAMPUS}_orig.csv', index=False, encoding='utf-8-sig')

allpriority_konkurs = konkurs[
    (konkurs['bvi'] == 0) & (konkurs['bvi_s'] == 'Нет') & (konkurs['work'] == 0) & (konkurs['special'] == 0) &
    (konkurs['separate'] == 0) & (konkurs['paid'] == 0)].groupby(['fio', 'name'])[
    ['green_admit', 'yellow_admit']].max().reset_index()
allpriority_konkurs = allpriority_konkurs.pivot_table(values=['green_admit', 'yellow_admit'], index=['name'],
                                                      aggfunc='sum').reset_index()
allpriority_konkurs.columns = ['name', 'ingreen_admit', 'inyellow_admit']
allpriority_konkurs.loc[:, 'inyellow_admit'] = allpriority_konkurs.loc[:, 'inyellow_admit'] - allpriority_konkurs.loc[:,
                                                                                              'ingreen_admit']
allpriority_konkurs.loc[:, 'inyellow_admit'] = allpriority_konkurs.apply(
    lambda x: '-' if final.set_index('name').loc[x['name'], 'yellow_num'] == '-' else x['inyellow_admit'], axis=1)

minkcp_konkurs = konkurs[
    (konkurs['bvi'] == 0) & (konkurs['bvi_s'] == 'Нет') & (konkurs['work'] == 0) & (konkurs['special'] == 0) & (
                konkurs['separate'] == 0) & (konkurs['paid'] == 0)].groupby(['fio', 'name'])[
    ['min_admit']].max().reset_index()
minkcp_konkurs = minkcp_konkurs.pivot_table(values=['min_admit'], index=['name'], aggfunc='sum').reset_index()
minkcp_konkurs.columns = ['name', 'kcp_leftover_admit']

minbvi_konkurs = \
konkurs[(konkurs['bvi'] == 1) & (konkurs['bvi_s'] == 'Да') & (konkurs['paid'] == 0)].groupby(['fio', 'name'])[
    ['min_admit']].max().reset_index()
minbvi_konkurs = minbvi_konkurs.pivot_table(values=['min_admit'], index=['name'], aggfunc='sum').reset_index()
minbvi_konkurs.columns = ['name', 'bvi_admit']

minwork_konkurs = konkurs[(konkurs['work'] == 1)].groupby(['fio', 'name'])[['min_admit']].max().reset_index()
minwork_konkurs = minwork_konkurs.pivot_table(values=['min_admit'], index=['name'], aggfunc='sum').reset_index()
minwork_konkurs.columns = ['name', 'work_admit']

minspecial_konkurs = konkurs[(konkurs['special'] == 1)].groupby(['fio', 'name'])[['min_admit']].max().reset_index()
minspecial_konkurs = minspecial_konkurs.pivot_table(values=['min_admit'], index=['name'], aggfunc='sum').reset_index()
minspecial_konkurs.columns = ['name', 'special_admit']

minseparate_konkurs = konkurs[(konkurs['separate'] == 1)].groupby(['fio', 'name'])[['min_admit']].max().reset_index()
minseparate_konkurs = minseparate_konkurs.pivot_table(values=['min_admit'], index=['name'], aggfunc='sum').reset_index()
minseparate_konkurs.columns = ['name', 'separate_admit']

priority1_konkurs = (konkurs_short[(konkurs_short['priority'] == 1) & (konkurs_short['bvi'] == 0) &
                                   (konkurs_short['work'] == 0) & (konkurs_short['separate'] == 0) &
                                   (konkurs_short['special'] == 0) & (konkurs_short['paid'] == 0)].
                     pivot_table(values=['green_admit', 'yellow_admit'], index=['name'], aggfunc='sum')).reset_index()
priority1_konkurs.columns = ['name', 'ingreen_admit_priority1', 'inyellow_admit_priority1']
priority1_konkurs.loc[:, 'inyellow_admit_priority1'] = priority1_konkurs.loc[:,
                                                       'inyellow_admit_priority1'] - priority1_konkurs.loc[:,
                                                                                     'ingreen_admit_priority1']
priority1_konkurs.loc[:, 'inyellow_admit_priority1'] = priority1_konkurs.apply(
    lambda x: '-' if final.set_index('name').loc[x['name'], 'yellow_num'] == '-' else x['inyellow_admit_priority1'],
    axis=1)

priority2_konkurs = (konkurs_short[(konkurs_short['priority'] == 2) & (konkurs_short['bvi'] == 0) &
                                   (konkurs_short['work'] == 0) & (konkurs_short['separate'] == 0) &
                                   (konkurs_short['special'] == 0) & (konkurs_short['paid'] == 0)].
                     pivot_table(values=['green_admit'], index=['name'], aggfunc='sum')).reset_index()
priority2_konkurs.columns = ['name', 'ingreen_admit_priority2']
priority3_konkurs = (konkurs_short[(konkurs_short['priority'] == 3) & (konkurs_short['bvi'] == 0) &
                                   (konkurs_short['work'] == 0) & (konkurs_short['separate'] == 0) &
                                   (konkurs_short['special'] == 0) & (konkurs_short['paid'] == 0)].
                     pivot_table(values=['green_admit'], index=['name'], aggfunc='sum')).reset_index()
priority3_konkurs.columns = ['name', 'ingreen_admit_priority3']

spisok['Образовательная программа'] = spisok['Конкурсная группа'].apply(rename_konkurs, args=[CAMPUS])

# creating pivots

bvi = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
              (spisok['Заявление забрано'] == 'Нет') & (spisok['Без ВИ'] == 'Да')]
       .pivot_table(values=['Код'], index=['Образовательная программа'], aggfunc='count')).reset_index()
bvi.columns = ['name', 'bvi_admitted']

# if CAMPUS != campuses[0]:
#     work = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
#                   (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
#            .pivot_table(values=['Целевое'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
#     work.columns = ['name', 'work_admitted']
#
#     if CAMPUS == campuses[3]:
#         work.loc[work['name'] == 'Компьютерные науки и технологии (направление подготовки 09.03.04 Программная инженерия)', 'work_admitted'] += 1
#         work.loc[work['name'] == 'Международный бакалавриат по бизнесу и экономике', 'work_admitted'] += 2
#     if CAMPUS == campuses[2]:
#         work.loc[work['name'] == 'Разработка информационных систем для бизнеса (направление подготовки 09.03.04 Программная инженерия)', 'work_admitted'] += 1
#         work.loc[work['name'] == 'Разработка информационных систем для бизнеса (направление подготовки 38.03.05 Бизнес-информатика)', 'work_admitted'] += 2
#     if CAMPUS == campuses[1]:
#         work.loc[work['name'] == 'Аналитика в экономике', 'work_admitted'] += 1
# else:
#     work = programs_work[['Образовательная программа', 0]]
#     work.columns = ['name', 'work_admitted']
# print(work.iloc[10:20], final.loc[10:20, ['work', 'work_max']])
# work.loc[:, 'work_admitted'] = work.apply(lambda x: np.nan if final.set_index('name').loc[x['name'], 'work'] == '-' else min(x['work_admitted'], final.set_index('name').loc[x['name'], 'work'], final.set_index('name').loc[x['name'], 'work_max']), axis=1)

special = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
                  (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
           .pivot_table(values=['Особая'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
special.columns = ['name', 'special_admitted']
special.loc[:, 'special_admitted'] = special.apply(
    lambda x: np.nan if final.set_index('name').loc[x['name'], 'special'] == '-' else min(x['special_admitted'],
                                                                                          final.set_index('name').loc[
                                                                                              x['name'], 'special']),
    axis=1)

separate = (spisok[(spisok['Платно'] == 0) & (spisok['Филиал'] == CAMPUS) &
                   (spisok['Заявление забрано'] == 'Нет') & (spisok['Приоритет'] == 1)]
            .pivot_table(values=['Отдельная'], index=['Образовательная программа'], aggfunc='sum')).reset_index()
separate.columns = ['name', 'separate_admitted']
separate.loc[:, 'separate_admitted'] = separate.apply(
    lambda x: np.nan if final.set_index('name').loc[x['name'], 'separate'] == '-' else min(x['separate_admitted'],
                                                                                           final.set_index('name').loc[
                                                                                               x['name'], 'separate']),
    axis=1)

admitted_other = (admitted_other.set_index(['name', 'fio']).
                  join(konkurs_short[['name', 'fio', 'priority']].set_index(['name', 'fio'])).reset_index())
admitted_other['min_priority'] = admitted_other.set_index('fio').join(
    admitted_other.groupby('fio')[['priority']].min(), rsuffix='_j').reset_index()['priority_j']

admitted_other = admitted_other[admitted_other['priority'] == admitted_other['min_priority']]
admitted_other = admitted_other.groupby('name')[['fio']].count().reset_index()
admitted_other.columns = ['name', 'possible_admit']

admitted_other_bvi = admitted_other_bvi.groupby('name')[['fio']].count().reset_index()
admitted_other_bvi.columns = ['name', 'possible_bvi']

pivots = [minkcp_konkurs, minbvi_konkurs, minwork_konkurs, minspecial_konkurs, minseparate_konkurs,
          bvi, special, separate, allpriority_konkurs, priority1_konkurs, priority2_konkurs,
          priority3_konkurs, admission_score, admitted_other, admitted_other_bvi]
pivot_variables = [['kcp_leftover_admit'], ['bvi_admit'], ['work_admit'], ['special_admit'], ['separate_admit'],
                   ['bvi_admitted'], ['special_admitted'], ['separate_admitted'],
                   ['ingreen_admit', 'inyellow_admit'], ['ingreen_admit_priority1', 'inyellow_admit_priority1'],
                   ['ingreen_admit_priority2'], ['ingreen_admit_priority3'], ['admit_score'], ['possible_admit'], ['possible_bvi']]

for pivot, prog_var in zip(pivots, pivot_variables):
    cols = col_correspondence[prog_var]
    pivot = programs.set_index('name').join(pivot.set_index('name'), rsuffix='_pivot').reset_index()
    pivot.loc[pivot['kcp'] == '-'] = pivot[pivot['kcp'] == '-'].fillna('-')
    pivot.loc[pivot['kcp'] != '-'] = pivot[pivot['kcp'] != '-'].fillna(0)
    write_programs(wb, ws, pivot, cols, prog_var, programs_only=True, old=False)



ws.autofilter(4, 0, num_row, num_col)
ws.freeze_panes(5, 1)
ws.set_landscape()
ws.repeat_rows(first_row=3, last_row=4)
ws.set_column(0, 0, 3)
ws.set_column(1, 1, 46.33)
ws.set_column(2, 2, 12.67)
ws.set_paper(8)
ws.set_margins(top=0.4, bottom=0.4, left=0.1, right=0.1)
ws.fit_to_pages(width=1, height=0)
ws.set_footer('&RСтраница &P из &N')
ws.set_row(4, 146)
ws.set_row(3, 52)

wb.close()

for campus in campuses:
    print(campus, spisok[(spisok['Филиал'] == campus) & (spisok['Заявление забрано'] == 'Нет')][
        ['Личное дело', 'Отдельная']].groupby('Личное дело').sum()['Отдельная'].shape)
print(spisok[(spisok['Заявление забрано'] == 'Нет') & (spisok['Платно'] == 0)][
          ['Конкурсная группа', 'Личное дело']].groupby('Личное дело').sum().shape)
