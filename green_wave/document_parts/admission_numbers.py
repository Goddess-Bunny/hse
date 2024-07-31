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


def write_aicpk_numbers(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE, nopaid=False):
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

    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(lambda x: x if x.find('(') == -1 else x[:x.find(' (')])
    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(
        lambda x: x if x.find('«') == -1 else x.replace('«', '\"').replace('»', '\"'))
    programs_adm.loc[:, 'name'] = programs_adm['name'].apply(lambda x: ' '.join(x.split()))

    if CAMPUS == 'Пермь':
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Юриспруденция'),
        'format'] = 'Очно-заочное обучение'
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Управление бизнесом'),
        'format'] = 'Очно-заочное обучение'

        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Юриспруденция') &
                         (programs_adm['format'] == 'Очно-заочное обучение'), 'name'] = 'Юриспруденция (очно-заочное)'
        programs_adm.loc[programs_adm['name'].str.contains(
            'Разработка информационных систем для бизнеса') & programs_adm['major'].str.contains(
            'Программная инженерия'), 'name'] = (
            'Разработка информационных систем для бизнеса (направление подготовки 09.03.04 Программная инженерия)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Разработка информационных систем для бизнеса') & programs_adm['major'].str.contains(
            'Бизнес-информатика'), 'name'] = (
            'Разработка информационных систем для бизнеса (направление подготовки 38.03.05 Бизнес-информатика)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Юриспруденция') & programs_adm['major'].str.contains(
            'Юриспруденция') & programs_adm['format'].str.contains('Очное'), 'name'] = (
            'Юриспруденция (направление подготовки 40.03.01 Юриспруденция)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Юриспруденция') & programs_adm['major'].str.contains(
            'Юриспруденция') & programs_adm['format'].str.contains('Очно-заочное'), 'name'] = (
            'Юриспруденция (направление подготовки 40.03.01 Юриспруденция) (Очно-заочное)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Управление бизнесом'), 'name'] = (
            'Управление бизнесом (Очно-заочное)')
    if CAMPUS == 'Нижний Новгород':
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Программная инженерия'),
        'format'] = 'Очно-заочное обучение'
        programs_adm.loc[(programs_adm['kcp'] == 0) & (programs_adm['name'] == 'Экономика и бизнес'),
        'format'] = 'Очно-заочное обучение'
        programs_adm.loc[programs_adm['name'] == 'Программная инженерия', 'name'] = (
            'Программная инженерия (Очно-заочное)')
        programs_adm.loc[programs_adm['name'] == 'Экономика и бизнес', 'name'] = (
            'Экономика и бизнес (Очно-заочное)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Компьютерные науки и технологии') & programs_adm['major'].str.contains(
            'Прикладная математика и информатика'), 'name'] = (
            'Компьютерные науки и технологии (направление подготовки 01.03.02 Прикладная математика и информатика)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Компьютерные науки и технологии') & programs_adm['major'].str.contains(
            'Программная инженерия'), 'name'] = (
            'Компьютерные науки и технологии (направление подготовки 09.03.04 Программная инженерия)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Компьютерные науки и технологии') & programs_adm['major'].str.contains(
            'Бизнес-информатика'), 'name'] = (
            'Компьютерные науки и технологии (направление подготовки 38.03.05 Бизнес-информатика)')

    if CAMPUS == 'Москва':
        programs_adm.loc[programs_adm['name'].str.contains("Актер"), 'name'] = 'Актёр'
        programs_adm.loc[programs_adm['name'].str.contains(
            'Экономика и политика Азии') & programs_adm['major'].str.contains('Зарубежное регионоведение'), 'name'] = ('Программа двух дипломов НИУ ВШЭ и Университета Кёнхи "Экономика и политика в Азии" (направление подготовки 41.03.01 Зарубежное регионоведение)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Экономика и политика Азии') & programs_adm['major'].str.contains('Публичная политика и социальные науки'), 'name'] = (
            'Программа двух дипломов НИУ ВШЭ и Университета Кёнхи "Экономика и политика в Азии" (направление подготовки 41.03.06 Публичная политика и социальные науки)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Международные отношения и глобальные исследования') & programs_adm['major'].str.contains(
            'Международные отношения'), 'name'] = (
            'Международная программа "Международные отношения и глобальные исследования" (направление подготовки 41.03.05 Международные отношения)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Международные отношения и глобальные исследования') & programs_adm['major'].str.contains(
            'Публичная политика'), 'name'] = (
            'Международная программа "Международные отношения и глобальные исследования" (направление подготовки 41.03.06 Публичная политика и социальные науки)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Глобальные цифровые коммуникации') & programs_adm['major'].str.contains(
            'Реклама и связи'), 'name'] = (
            'Глобальные цифровые коммуникации (направление подготовки 42.03.01 Реклама и связи с общественностью)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Глобальные цифровые коммуникации') & programs_adm['major'].str.contains(
            'Медиакоммуникации'), 'name'] = (
            'Глобальные цифровые коммуникации (направление подготовки 42.03.05 Медиакоммуникации)')
        programs_adm.loc[(programs_adm['name'] ==
             'Реклама и связи с общественностью') & programs_adm['major'].str.contains(
             'Реклама и связи с общественностью'), 'name'] = (
             'Реклама и связи с общественностью (направление подготовки 42.03.01 Реклама и связи с общественностью)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Реклама и связи с общественностью') & programs_adm['major'].str.contains(
            'Медиакоммуникации'), 'name'] = (
            'Реклама и связи с общественностью (направление подготовки 42.03.05 Медиакоммуникации)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Античность') & programs_adm['major'].str.contains(
            'Филология'), 'name'] = (
            'Античность (направление подготовки 45.03.01 Филология)')
        programs_adm.loc[programs_adm['name'].str.contains(
            'Античность') & programs_adm['major'].str.contains(
            'История'), 'name'] = (
            'Античность (направление подготовки 46.03.01 История)')
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
    if CAMPUS != 'Москва':
        programs_adm.loc[programs_adm['name'].str.contains("бакалавриат по бизнесу и экономике"), 'name'] = \
            'Бакалаврская программа «Международный бакалавриат по бизнесу и экономике»'

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
    print(join_programs.loc[40:50, ['name', 'major', 'kcp_admit']])
    print(programs.loc[40:50, ['name', 'major']])
    print(programs_adm.loc[40:50, ['name', 'major']])
    join_programs.loc[join_programs['kcp'] != '-'] = (join_programs.loc[join_programs['kcp'] != '-'].
                                                      infer_objects(copy=False).fillna(0))
    join_programs.loc[join_programs['kcp'] == '-'] = join_programs.loc[join_programs['kcp'] == '-'].fillna('-')

    program_variables = ['kcp_admit', 'work_admit', 'special_admit', 'separate_admit', 'bvi_admit']
    if not nopaid:
        program_variables += ['paid']
    cols = col_correspondence[program_variables]
    write_programs(wb, ws, join_programs, cols, program_variables, old=False, programs_only=True)

    return total_inc_students
