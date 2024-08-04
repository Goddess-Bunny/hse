import pandas as pd
import numpy as np
import re


def nan_process(x):
    if pd.isna(x):
        return '-'
    else:
        return x


# in order to parse, pass the data_frame with column 'name', which contains names of programs, along with
# formats and majors. All other columns must have some name
def parse_list(df, lvl, verbose=False):
    names = list(df.columns)[1:]

    programs = {'campus': [], 'lvl': [], 'group': [], 'major': [], 'name': [], 'format': []}

    if verbose:
        programs |= {'new': [], 'rename': [], 'terminated': [], 'online': [], 'paid_only': []}

    programs |= {name: [] for name in names}

    _format, group, major = '', '', ''

    if lvl == 0:
        lvl_main = 'Бакалавриат'
    else:
        lvl_main = 'Магистратура'

    for index, row in df.iterrows():
        if pd.isna(row['name']):
            continue

        if 'обучен' in row['name'].lower() and 'очн' in row['name'].lower():
            _format = 'Очное обучение'
            if "заочн" in row['name'].lower():
                if '-' in row['name'].lower():
                    _format = "Очно-заочное обучение"
                else:
                    _format = "Заочное обучение"

            continue
        elif bool(re.search(r'\d{2}\.00\.00', row['name'])):
            group = row['name']
            continue

        elif bool(re.search(r'\d{2}\.\d{2}\.\d{2}', row['name'])):
            lvl_towrite = lvl_main
            if 'направлени' in row['name'].lower():
                major = row['name'][row['name'].find(' ') + 1:].strip('\"')
                if 'подготов' in row['name'].lower():
                    major = major[major.find(' ') + 1:].strip('\"')
            elif 'специальность' in row['name'].lower():
                major = row['name'][row['name'].find(' ') + 1:].strip('\"')
                lvl_towrite = 'Специалитет'
            else:
                major = row['name']

            continue

        new_prog, rename, terminated, campus, online = False, '-', False, 'Москва', False

        name = row['name'].strip()

        if 'онлайн' in name.lower() and '(онлайн)' not in name.lower():
            idx = name.lower().find('онлайн')
            name = name.replace(name[idx:idx+6], "").strip()
            online = True

        if name.rfind(')') != -1:
            # here we separate bracketed words from the whole program name;
            # brackets contain the campus name, whether a program is online, whether it is "twinkling".
            # here we assume that meta info of the program comes from the right side, up until
            # brackets are part of a program's name.
            # res variable contains substring after brackets - those contain meta info of the program
            close_bracket = name.rfind(')')

            while name.rfind(')') != -1:
                word = name[name.rfind('(') + 1:name.rfind(')')]
                end_of_digits = re.search(r'\d{2}\.\d{2}\.\d{2}', major).span()[1]
                major_digitless = major[end_of_digits+1:]

                if 'пермь' in word.lower():
                    campus = 'Пермь'
                elif 'санкт' in word.lower():
                    campus = 'Санкт-Петербург'
                elif 'москва' in word.lower():
                    campus = 'Москва'
                elif 'нижний' in word.lower():
                    campus = 'Нижний Новгород'
                elif 'онлайн' in word.lower():
                    online = True
                elif 'мерцающая' in word.lower():
                    new_prog = True
                elif 'направление' in word.lower():
                    pass
                elif major_digitless.lower() == word.lower():
                    pass
                else:
                    break

                name = name[:name.rfind('(') - 1]

            if len(row['name'][close_bracket + 1:]) > 0:
                res = row['name'][close_bracket + 2:]

                if 'не планируется' in res.lower():
                    terminated = True

                if 'новая' in res.lower():
                    new_prog = True
                elif 'переименование' in res.lower():
                    if '\"' in res:
                        start = res.find('\"') + 1
                        rename = res[start:-1]
                    else:
                        start = res.find('переименование программы ') + len('переименование программы ')
                        rename = res[start:]

                if rename.rfind('(') != -1:
                    rename = rename[:rename.rfind('(') - 1]

        if verbose:
            if 'online' not in names:
                programs['online'].append(online)
            programs['new'].append(new_prog)
            programs['rename'].append(rename)
            programs['terminated'].append(terminated)

        programs['lvl'].append(lvl_towrite.replace('  ', ' ').strip())
        programs['format'].append(_format.replace('  ', ' ').strip())
        programs['group'].append(group.replace('  ', ' ').strip())
        programs['major'].append(major.replace('  ', ' ').strip())
        programs['name'].append(name.replace('  ', ' ').strip())
        programs['campus'].append(campus.replace('  ', ' ').strip())

        for name in names:
            programs[name].append(nan_process(row[name]))

    return pd.DataFrame(programs)


def join_new_old_progs(new, old, LVL, program_nans='-'):
    old = parse_list(old, LVL)
    old = old[~(old['format'] == 'Заочное обучение')]
    old['name_old'] = old['name'].apply(lambda x: x.strip('*'))
    old = old.map(lambda x: x.strip().replace('ё', 'е') if type(x) is str else x)

    old.columns = [column + '_old' if column not in new.columns and '_old' not in column
                   else column for column in old.columns]

    pre_joined = new.map(lambda x: x.strip().replace('ё', 'е') if type(x) is str else x)

    joined = (pre_joined.set_index(['format', 'major', 'name_old']).join(
        old.set_index(['format', 'major', 'name_old']), rsuffix='_old').
              reset_index(level=[0, 1]).reset_index(drop=True).fillna(program_nans))

    #print(joined['name'][joined['kcp'] == '-'])

    return joined
