import pandas as pd
from aux_program_scripts import write_programs, parse_list
from datetime import datetime
import numpy as np


def process_number(x):
    x = x.split(' ')
    ans = 0

    for pow_, triple in enumerate(x):
        ans += int(triple) * 10 ** (3 * (len(x) - pow_ - 1))

    return ans


def write_asav_numbers(wb, ws, programs, col_correspondence, LVL, CAMPUS, CUR_DATE, PAID_ONLY, pathOfInputs):
    programs_adm = pd.read_excel(pathOfInputs + r'\tables_2024\mag\\' + fr'{CAMPUS}\{CUR_DATE.strftime('%Y-%m-%d')}\mag_adm.xlsx')

    programs_adm = programs_adm.iloc[5:199, :7]
    total_inc_students = programs_adm.iloc[-1, 1]

    programs_adm = programs_adm.iloc[:-5, :]
    programs_adm.columns = ['name', 'kcp', 'work', 'paid', 'kcp_admit', 'work_admit', 'paid_admit']
    programs_adm.iloc[:, 1:] = programs_adm.iloc[:, 1:].map(lambda x: x if x == '-' or type(x) is not str else int(x))

    for i in range(3):
        for j in range(programs_adm.shape[0]):
            if pd.isna(programs_adm.iloc[j, 4 + i]):
                if not pd.isna(programs_adm.iloc[j, 1 + i]) and not programs_adm.iloc[j, 1 + i] == '-':
                    if programs_adm.iloc[j, 1 + i] > 0:
                        programs_adm.iloc[j, 4 + i] = 0

    programs_adm = programs_adm.iloc[:, [0] + list(range(3, 7))]

    adm_df = parse_list(programs_adm, LVL)

    if CAMPUS == 'Москва':
        adm_df.loc[adm_df['name'].str.contains('ЦПМ'), 'name'] = \
            'Совместная магистратура НИУ ВШЭ и Центра педагогического мастерства'

        adm_df.loc[adm_df['name'].str.contains('Системный анализ'), 'major'] = \
            '01.04.02 Прикладная математика и информатика, 01.04.04 Прикладная математика'
    if CAMPUS == 'Санкт-Петербург':
        adm_df.loc[adm_df['name'].str.contains('бизнеса и экономики'), 'name'] = \
            adm_df.loc[adm_df['name'].str.contains('бизнеса и экономики'), 'name'].apply(
                lambda x: x[:x.rfind(' ')]
        )

        adm_df = adm_df.groupby('name').sum(min_count=1).reset_index()
        adm_df.loc[adm_df['name'].str.contains('бизнеса и экономики'), 'major'] = \
            '38.04.01 Экономика, 38.04.02 Менеджмент'

        adm_df.loc[adm_df['name'].str.contains('государства и общества'), 'name'] = \
        adm_df.loc[adm_df['name'].str.contains('государства и общества'), 'name'].apply(
            lambda x: x[:x.rfind(' ')]
        )
    if CAMPUS == 'Нижний Новгород':
        adm_df.loc[adm_df['name'].str.contains('экономике и менеджменте'), 'name'] = \
            adm_df.loc[adm_df['name'].str.contains('экономике и менеджменте'), 'name'].apply(
                lambda x: x[:x.rfind(' ')]
            )
        adm_df = adm_df.map(lambda x: np.nan if x == '-' else x)
        print(adm_df['kcp_admit'])

        adm_df = adm_df.groupby('name').sum(min_count=1).reset_index()
        adm_df.loc[adm_df['name'].str.contains('экономике и менеджменте'), 'major'] = \
            '38.04.01 Экономика, 38.04.02 Менеджмент'
    #if CAMPUS == 'Пермь':


    if programs.loc[programs['school'] != '-', 'school'].shape[0] > 0:
        adm_df.loc[adm_df['name'].isin(
            programs.loc[programs['school'] != '-', 'name']), 'paid'] -= \
            programs.loc[programs['school'] != '-', 'school'].iloc[0]

    if not PAID_ONLY:
        program_variables = ['kcp_admit', 'work_admit', 'paid', 'paid_admit']
    else:
        program_variables = ['paid', 'paid_admit']

    join_programs = (programs.set_index(['major', 'name'])
                     .join(adm_df.set_index(['major', 'name']), rsuffix='2').reset_index())
    join_programs.loc[join_programs['kcp'] != '-'] = (join_programs.loc[join_programs['kcp'] != '-']
                                                      .infer_objects(copy=False).fillna(0))
    join_programs.loc[join_programs['kcp'] == '-'] = join_programs.loc[join_programs['kcp'] == '-'].fillna('-')

    cols = col_correspondence[program_variables]
    write_programs(wb, ws, join_programs, cols, program_variables, old=False)

    return total_inc_students


def write_aicpk_numbers(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE):
    programs_adm = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/tables_2024/bac/' + f'{CAMPUS}/' +
                                 f'{CUR_DATE.strftime('%Y-%m-%d')}/bac_adm.xlsx')

    programs_adm = programs_adm.iloc[7:78, [0, 1] + list(range(5, 17))].reset_index(drop=True)
    total = programs_adm.iloc[-1, 0]
    total = total[total.find(' ') + 1:]
    total_inc_students = process_number(total[total.find(' ')+1:])

    programs_adm = programs_adm.iloc[:-2, 1:]
    programs_adm.columns = ['name', 'major', 'kcp', 'work', 'special', 'separate', 'kcp_admit', 'work_admit',
                            'special_admit', 'separate_admit', 'bvi_admit', 'paid', 'paid_admit']
    programs_adm.iloc[:, 2:] = programs_adm.iloc[:, 2:].map(
        lambda x: x if type(x) is not str else
        process_number(x) if x.find('+') == -1 else process_number(x[:x.find(' ')]))

    programs_adm['format'] = 'Очное обучение'

    if CAMPUS == 'Пермь':
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Юриспруденция'),
        'format'] = 'Очно-заочное обучение'
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Управление бизнесом'),
        'format'] = 'Очно-заочное обучение'

        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Юриспруденция') &
                         (programs_adm['format'] == 'Очно-заочное обучение'), 'name'] = 'Юриспруденция (очно-заочное)'
    if CAMPUS == 'Нижний Новгород':
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Программная инженерия'),
        'format'] = 'Очно-заочное обучение'
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Экономика и бизнес'),
        'format'] = 'Очно-заочное обучение'

    if CAMPUS == 'Москва':
        programs_adm.loc[programs_adm['name'].str.contains("Актер"), 'name'] = 'Актёр'
        programs_adm.loc[programs_adm['name'].str.contains(
            'Экономика и политика Азии'), 'name'] = ('Программа двух дипломов НИУ ВШЭ' +
                                                     ' и Университета Кёнхи \"Экономика и политика в Азии\"')
        programs_adm.loc[
            (programs_adm['kcp'] == 0) &
            (programs_adm['name'] == 'Юриспруденция: правовое регулирование бизнеса'), 'format'] =\
            'Очно-заочное обучение'
    else:
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), 'name'] = \
            'Многопрофильный конкурс "Международный бакалавриат по бизнесу и экономике"'
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), 'major'] = \
            '38.03.01 Экономика, 38.03.02 Менеджмент'
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), :] = \
            (programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), :].
             infer_objects(copy=False).fillna(0))

        programs_adm = programs_adm.groupby('name').sum(min_count=1).reset_index()
        programs_adm = programs_adm.map(lambda x: np.nan if x is None else x)
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), 'major'] = \
            '38.03.01 Экономика, 38.03.02 Менеджмент'
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), 'format'] = \
            'Очное обучение'

    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(lambda x: x if x.find('(') == -1 else x[:x.find(' (')])
    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(
        lambda x: x if x.find('«') == -1 else x.replace('«', '\"').replace('»', '\"'))
    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(lambda x: ' '.join(x.split()))

    for i in range(4):
        for j in range(programs_adm.shape[0]):
            if pd.isna(programs_adm.iloc[j, 6 + i]):
                if not pd.isna(programs_adm.iloc[j, 2 + i]):
                    if programs_adm.iloc[j, 2 + i] > 0:
                        programs_adm.iloc[j, 6 + i] = 0

    print(programs_adm)

    for j in range(programs_adm.shape[0]):
        if pd.isna(programs_adm.iloc[j, 12]):
            programs_adm.iloc[j, 12] = 0

    join_programs = (programs.set_index(['format', 'major', 'name']).
                     join(programs_adm.set_index(['format', 'major', 'name']), rsuffix='2').reset_index())
    print(join_programs.loc[:, ['name', 'major', 'work_admit']])
    join_programs.loc[join_programs['kcp'] != '-'] = (join_programs.loc[join_programs['kcp'] != '-'].
                                                      infer_objects(copy=False).fillna(0))
    join_programs.loc[join_programs['kcp'] == '-'] = join_programs.loc[join_programs['kcp'] == '-'].fillna('-')

    program_variables = ['kcp_admit', 'work_admit', 'special_admit', 'separate_admit', 'bvi_admit', 'paid',
                         'paid_admit']

    cols = col_correspondence[program_variables]
    write_programs(wb, ws, join_programs, cols, program_variables, old=False)

    return total_inc_students
