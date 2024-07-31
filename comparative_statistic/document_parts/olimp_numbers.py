import pandas as pd
from aux_program_scripts import parse_list, write_programs
from datetime import datetime


def write_yaprofi_potential(wb, ws, programs, col_correspondence, CAMPUS):
    MAG_AUX = 2

    programs_yaprofi = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/yaprofi/{CAMPUS}/yaprofi.xlsx',
                                     header=None)

    programs_yaprofi = programs_yaprofi.iloc[:, [0, 1, 2, 3, 5, 6, 7]]
    programs_yaprofi.columns = ['name', 'year', 'olimp', 'profile', 'medal', 'winner', 'prize']

    # pre-processing of programs with yaprofi info
    yaprofi_df = parse_list(programs_yaprofi, MAG_AUX)
    yaprofi_df['olimp'] = yaprofi_df['olimp'].str.split(';')
    yaprofi_df = yaprofi_df.explode("olimp").reset_index(drop=True)
    yaprofi_df.loc[:, 'olimp'] = yaprofi_df['olimp'].apply(lambda x: x.strip())

    # pre-processing of winners tables
    yaprofi_winners = pd.read_excel(
        '/Users/s/Desktop/Admission Numbers 2024/yaprofi/yaprofi_info.xlsx', sheet_name='Список')
    yaprofi_winners['level'] = yaprofi_winners['Итоговый статус'].map({'Призер': 3, 'Победитель': 2,
                                                                       'Золотой медалист': 1, 'Бронзовый медалист': 1,
                                                                       'Серебряный медалист': 1, 'Медалист': 1})

    # leaving only 2023/2024 preferences, modifying the name of the olimp track to include the name of profile
    yaprofi_df = yaprofi_df[yaprofi_df['year'] == '2023/2024']
    yaprofi_df['level'] = yaprofi_df[['medal', 'winner', 'prize']].apply(lambda x: int(sum(x)), axis=1)
    yaprofi_df.loc[yaprofi_df['profile'] != '-', 'olimp'] = \
        (yaprofi_df.loc[yaprofi_df['profile'] != '-', 'olimp'] +
         ' "' + yaprofi_df.loc[yaprofi_df['profile'] != '-', 'profile'] + '"')

    # creating a copy of winners who participated in a track with a profile
    yaprofi_winners_profile = yaprofi_winners[yaprofi_winners['Профиль участия'] != '-'].copy()
    yaprofi_winners_profile.loc[:, 'Направление'] = (yaprofi_winners_profile['Направление'] + ' "' +
                                                     yaprofi_winners_profile['Профиль участия'] + '"')
    # adding their copies to the winners list
    yaprofi_winners = pd.concat([yaprofi_winners, yaprofi_winners_profile], ignore_index=True)

    # for each winner specifying which program he can apply to
    join_yp = yaprofi_winners.set_index('Направление').join(yaprofi_df.set_index('olimp'),
                                                            rsuffix='_program').reset_index()
    # rename a key column which will be printed on the final table
    join_yp['yaprofi_potential'] = join_yp['UID']

    a = join_yp[join_yp['level'] <= join_yp['level_program']]
    print(a[a['major'] == '38.04.02 Менеджмент']['name'].unique())

    # filter those who really have prefs on a given program, and count them by the name of a program/major/format/total
    yaprofi_nums = [join_yp[join_yp['level'] <= join_yp['level_program']].
                    groupby(column)[['yaprofi_potential']].nunique()
                    for column in ['name', ['format', 'major'], 'format', 'yaprofi_potential']]
    yaprofi_nums[3] = yaprofi_nums[3].sum()
    yaprofi_nums[0] = programs.set_index('name').join(yaprofi_nums[0]).reset_index()

    yaprofi_admissible = ~pd.isna(yaprofi_nums[0]['yaprofi_potential'])

    yaprofi_nums[0].loc[~pd.isna(yaprofi_nums[0]['yaprofi_potential'])] = \
        (yaprofi_nums[0].loc[~pd.isna(yaprofi_nums[0]['yaprofi_potential'])]
         .infer_objects(copy=False).fillna(0))
    yaprofi_nums[0].loc[pd.isna(yaprofi_nums[0]['yaprofi_potential'])] = \
        yaprofi_nums[0].loc[pd.isna(yaprofi_nums[0]['yaprofi_potential'])].fillna('-')

    program_variables = ['yaprofi_potential']
    cols = col_correspondence[program_variables]

    write_programs(wb, ws, yaprofi_nums[0], cols, program_variables,
                   old=False, total=yaprofi_nums[3], format_df=yaprofi_nums[2], major_df=yaprofi_nums[1])

    return yaprofi_admissible


def write_vl_potential(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE):
    def process_status_diploma(status):
        if 'дипломанты' in status:
            degrees = status[status.find('дипломанты') + len("дипломанты ") + 1:status.find(' степени')]

            if degrees.count(' ') == 0:
                return 1
            else:
                return degrees.count(' ')
        else:
            return -1

    join_admits = get_olimp_admits(programs, CAMPUS, CUR_DATE)

    programs_vl = (pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/vl/{CAMPUS}/vl.xlsx').dropna(how='all').
                   reset_index(drop=True))
    programs_vl = programs_vl.map(lambda x: x.strip() if type(x) is str else x)
    programs_vl['required_degree'] = programs_vl['status'].apply(process_status_diploma)
    programs_vl['required_medal'] = programs_vl['status'].apply(lambda x: 1 if 'медалисты' in x.lower() else 0)

    programs_vl_major = programs_vl[(pd.isna(programs_vl['track'])) | (programs_vl['track'] == 'Все треки')]
    programs_vl_sometrack = programs_vl[(programs_vl['track'] != 'Все треки') & (~pd.isna(programs_vl['track']))]
    if programs_vl_sometrack.shape[0] > 0:
        programs_vl_sometrack['major.track'] = programs_vl_sometrack.apply(
            lambda x: x['major'] + '"' + x['track'] + '"',
            axis=1)
    vl_winners = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/vl/vl_info.xlsx', sheet_name='Список')

    vl_winners = vl_winners.map(lambda x: x.strip() if type(x) is str else x)
    vl_winners['medal'] = vl_winners['Результат'].apply(lambda x: 1 if x == 'Медалист' else 0)
    vl_winners['diploma'] = vl_winners['Результат'].map(
        {'Диплом I степени': 1, 'Диплом II степени': 2, 'Диплом III степени': 3, 'Медалист': 10})
    vl_winners['major.track'] = vl_winners.apply(
        lambda x: x['Направление'] + '"' + x['Трек'] + '"' if not pd.isna(x['Трек']) else ' ', axis=1)

    vl_winners_major = (vl_winners.set_index('Направление').
                        join(programs_vl_major.set_index('major'), rsuffix='_program', how='inner').reset_index())
    if programs_vl_sometrack.shape[0] > 0:
        vl_winners_sometrack = (vl_winners[~pd.isna(vl_winners['Трек'])].set_index('major.track').
                                join(programs_vl_sometrack.set_index('major.track'),
                                     rsuffix='_program', how='inner').reset_index())
        name_major_column = 'major_program'
    else:
        vl_winners_sometrack = pd.DataFrame()
        name_major_column = 'major'

    vl_winners = pd.concat([vl_winners_major, vl_winners_sometrack], ignore_index=True)

    vl_winners = vl_winners.set_index('name').join(programs.set_index('name'), rsuffix='_program').reset_index()
    vl_winners['vl_potential'] = vl_winners['Рег номер']

    vl_nums = [vl_winners[(vl_winners['diploma'] <= vl_winners['required_degree']) |
                          (vl_winners['medal'] >= vl_winners['required_medal'])].
               groupby(column)[['vl_potential']].nunique()
               for column in ['name', ['format', name_major_column], 'format', 'vl_potential']]
    vl_nums[3] = vl_nums[3].sum()

    program_variables = ['vl_potential']
    cols = col_correspondence[program_variables]

    t = programs.set_index('name').join(
        vl_nums[0]).reset_index()

    vl_admissible = ~pd.isna(t['vl_potential'])

    t.loc[~pd.isna(t['vl_potential'])] = (t.loc[~pd.isna(t['vl_potential'])].infer_objects(copy=False).fillna(0))
    t.loc[pd.isna(t['vl_potential'])] = t.loc[pd.isna(t['vl_potential'])].fillna('-')

    for idx, row in t.iterrows():
        prog_join_vl = join_admits.loc[(join_admits['format']==row['format']) & (join_admits['major']==row['major']) &
                                    (join_admits['name']==row['name']), 'vl_admit'].iloc[0]
        print(prog_join_vl, row['vl_potential'])
        if not pd.isna(prog_join_vl) and row['vl_potential'] != '-' and row['vl_potential'] < prog_join_vl:
            diff = prog_join_vl - row['vl_potential']
            t.loc[idx, 'vl_potential'] += diff
            vl_nums[1].loc[(row['format'], row['major']), 'vl_potential'] += diff
            vl_nums[2].loc[row['format'], 'vl_potential'] += diff
            vl_nums[3] += diff

    write_programs(wb, ws, t, cols, program_variables,
                   old=False, total=vl_nums[3], format_df=vl_nums[2], major_df=vl_nums[1])

    return vl_admissible


def write_rp_potential(wb, ws, programs, col_correspondence, CAMPUS):
    rp_winners = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/rp/rp.xlsx').iloc[1:, [1, 29, 30]]
    rp_winners.columns = ['early_potential', 'campus', 'name']
    rp_winners.loc[:, 'campus'] = rp_winners['campus'].apply(
        lambda x: x[x.find('-') + 2:] if x.find('-') != -1 else x)

    rp_winners = rp_winners.loc[rp_winners['campus'] == CAMPUS, ['early_potential', 'name']]

    joined_rp = rp_winners.set_index('name').join(programs.set_index('name'), lsuffix='_winner').reset_index()

    rp_nums = [joined_rp.groupby(column)[['early_potential']].nunique()
               for column in ['name', ['format', 'major'], 'format', 'early_potential']]
    rp_nums[3] = rp_nums[3].sum()

    program_variables = ['early_potential']
    cols = col_correspondence[program_variables]

    t = programs.set_index('name').join(
        rp_nums[0]).reset_index()

    early_admissible = ~pd.isna(t['early_potential'])

    t.loc[~pd.isna(t['early_potential'])] = (t.loc[~pd.isna(t['early_potential'])].infer_objects(copy=False).fillna(0))
    t.loc[pd.isna(t['early_potential'])] = t.loc[pd.isna(t['early_potential'])].fillna('-')

    write_programs(wb, ws, programs.set_index('name').join(
        rp_nums[0]).reset_index().fillna('-'), cols, program_variables,
                   old=False, total=rp_nums[3], format_df=rp_nums[2], major_df=rp_nums[1])

    return early_admissible


def get_olimp_admits(programs, CAMPUS, CUR_DATE):
    def olimp_names(name):
        if 'Я-профессионал' in name:
            return 'Я-профессионал'
        elif 'Высшая лига' in name:
            return 'Высшая лига'
        elif 'Раннее приглашение' in name:
            return 'Раннее приглашение'
        else:
            return name

    olimp_admits = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/tables_2024/mag/' + f'{CAMPUS}/' +
                                 f'{CUR_DATE.strftime('%Y-%m-%d')}/mag_all_adm.xlsx')

    # 0: id, 2: olimp is active, 3: name of olimp, 25: name, 26: campus
    olimp_admits = olimp_admits.iloc[1:, [0, 2, 3, 18, 25, 26]]
    olimp_admits.columns = ['id', 'olimp_status', 'olimp_name', 'origin', 'name', 'campus']

    olimp_admits.loc[:, 'olimp_name'] = olimp_admits['olimp_name'].apply(olimp_names)

    olimp_admits.loc[:, 'campus'] = olimp_admits['campus'].apply(
        lambda x: x[x.find('-') + 2:] if x.find('-') != -1 else x)

    olimp_admits = olimp_admits.loc[(olimp_admits['olimp_status'] == 'Да') &
                                    (~pd.isna(olimp_admits['olimp_name'])) &
                                    (olimp_admits['campus'] == CAMPUS) & pd.isna(olimp_admits['origin']), :]

    if CAMPUS == 'Москва':
        olimp_admits.loc[olimp_admits['name'] == 'Совместная магистратура НИУ ВШЭ и ЦПМ', 'name'] =\
            'Совместная магистратура НИУ ВШЭ и Центра педагогического мастерства'

    olimp_admits = olimp_admits.groupby(['id', 'olimp_name', 'name']).nunique().reset_index()
    olimp_admits.loc[:, 'yaprofi_admit'] = olimp_admits['olimp_name'].apply(
        lambda x: 1 if 'Я-профессионал' in x else 0)
    olimp_admits.loc[:, 'vl_admit'] = olimp_admits.apply(
        lambda x: 1 if ('Высшая лига' in x['olimp_name']) else 0, axis=1)
    olimp_admits.loc[:, 'early_admit'] = olimp_admits.apply(
        lambda x: 1 if ('Раннее приглашение' in x['olimp_name']) else 0, axis=1)
    olimp_admits.loc[:, 'other_admit'] = olimp_admits.apply(
        lambda x: 0 if x['early_admit'] + x['yaprofi_admit'] + x['vl_admit'] >= 1 else 1, axis=1)

    olimp_admits = (olimp_admits.groupby(['name', 'id'])[['early_admit', 'yaprofi_admit', 'vl_admit', 'other_admit']]
                    .sum().reset_index())

    olimp_admits.loc[:, 'vl_admit'] = olimp_admits.apply(
        lambda x: 1 if x['vl_admit'] > 0 and x['yaprofi_admit'] < 1 else 0, axis=1)
    olimp_admits.loc[:, 'early_admit'] = olimp_admits.apply(
        lambda x: 1 if (x['early_admit'] > 0 and x['yaprofi_admit'] + x['vl_admit'] < 1) else 0,
        axis=1)
    olimp_admits.loc[:, 'other_admit'] = olimp_admits.apply(
        lambda x: 1 if x['early_admit'] + x['yaprofi_admit'] + x['vl_admit'] < 1 and x['other_admit'] > 0 else 0,
        axis=1)

    olimp_admits = olimp_admits.groupby('name')[['early_admit', 'yaprofi_admit', 'vl_admit', 'other_admit']].sum()

    join_admits = olimp_admits.join(programs.set_index('name'), how='right').reset_index()

    return join_admits


def write_olimp_admits(wb, ws, programs, col_correspondence, admissible, CAMPUS, CUR_DATE):
    join_admits = get_olimp_admits(programs, CAMPUS, CUR_DATE)

    join_admits.loc[admissible['yaprofi'], 'yaprofi_admit'] = \
        join_admits.loc[admissible['yaprofi'], 'yaprofi_admit'].infer_objects(copy=False).fillna(0)
    join_admits.loc[~admissible['yaprofi'], 'yaprofi_admit'] = '-'
    join_admits.loc[admissible['vl'], 'vl_admit'] = \
        join_admits.loc[admissible['vl'], 'vl_admit'].infer_objects(copy=False).fillna(0)
    join_admits.loc[~admissible['vl'], 'vl_admit'] = '-'
    join_admits.loc[admissible['early'], 'early_admit'] = \
        join_admits.loc[admissible['early'], 'early_admit'].infer_objects(copy=False).fillna(0)
    join_admits.loc[~admissible['early'], 'early_admit'] = '-'

    join_admits = join_admits.fillna(0).map(
        lambda x: int(x) if type(x) is float else x).reset_index()

    program_variables = ['early_admit', 'yaprofi_admit', 'vl_admit', 'other_admit']
    cols = col_correspondence[program_variables]

    write_programs(wb, ws, join_admits, cols, program_variables, old=False)


def write_olimp_vsosh(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE):
    comp_group = pd.read_excel(f'/Users/s/Desktop/Admission Numbers 2024/tables_2024/bac/' + f'{CAMPUS}/' +
                               f'{CUR_DATE.strftime('%Y-%m-%d')}/bvi.xlsx')
    # 7 - id, 8 - group name, 10 - bvi, 11 - bvi_doc, 16 - campus, 20 - major
    comp_group = comp_group.iloc[:, [2, 5, 6, 8, 10, 11, 16, 20]]
    comp_group.columns = ['closed', 'id', 'status', 'name', 'bvi', 'bvi_doc', 'campus', 'major_nonums']
    comp_group = comp_group[comp_group['name'].str.contains('(О Б)')]
    comp_group.loc[:, 'name'] = comp_group['name'].apply(lambda x: x[:x.find(' (')])

    abitur_bvi = comp_group[
        (comp_group['bvi'] == 'Да') &
        (comp_group['campus'] == CAMPUS) & (comp_group['closed'] == 'Нет')]
    abitur_bvi.loc[:, 'bvi_doc'] = abitur_bvi['bvi_doc'].apply(lambda x: x[x.find(':') + 2:x.find(',')])
    abitur_bvi = abitur_bvi[abitur_bvi['bvi_doc'] == 'Всероссийская олимпиада школьников']
    abitur_bvi = abitur_bvi.groupby(['id', 'name', 'major_nonums']).nunique().reset_index()

    abitur_bvi = abitur_bvi.groupby(['name', 'major_nonums']).count()


    programs['major_nonums'] = programs['major'].apply(lambda x: x[x.find(' ') + 1:])
    joined_bvi = programs.set_index(['name', 'major_nonums']).join(abitur_bvi, rsuffix='_bvi').reset_index()
    joined_bvi['vsosh_admit'] = joined_bvi['id_bvi']

    program_variables = ['vsosh_admit']
    cols = col_correspondence[program_variables]

    joined_bvi.loc[joined_bvi['kcp'] == '-', 'vsosh_admit'] = '-'
    joined_bvi.loc[joined_bvi['kcp'] != '-', 'vsosh_admit'] = \
        joined_bvi.loc[joined_bvi['kcp'] != '-', 'vsosh_admit'].fillna(0)

    write_programs(wb, ws, joined_bvi, cols, program_variables, old=False)
