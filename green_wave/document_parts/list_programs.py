from aux_program_scripts.xlsxtools import create_format


def list_programs(wb, ws, programs, row=5):
    formats = programs['format'].unique()
    majors = programs['major'].unique()

    num_of_programs = programs.groupby(['name']).nunique()['id']
    id = 1

    for _format in formats:
        for major in sorted(majors, key=lambda x: int(x[:2]) * 10e4 + int(x[3:5]) * 10e2 + int(x[6:8])):
            names = programs[(programs['format'] == _format) &
                             (programs['major'] == major)]['name']

            if len(names) == 0:
                continue

            for name in names:
                color, italic, num_stars = None, False, 0

                program = programs[
                    (programs['format'] == _format) & (programs['major'] == major) & (programs['name'] == name)]

                if program['paid_only'].iloc[0]:
                    italic = True

                if name.find('/') != -1:
                    name = name[:name.find('/')].strip()

                if num_of_programs.loc[name] > 1:
                    programs.loc[
                        (programs['format'] == _format) & (programs['major'] == major) &
                        (programs['name'] == name), 'name'] = name + ' (направление подготовки ' + major + ')'
                    name += ' (направление подготовки ' + major + ')'

                if _format == 'Очно-заочное обучение':
                    programs.loc[
                        (programs['format'] == _format) & (programs['major'] == major) &
                        (programs['name'] == name), 'name'] = name + ' (Очно-заочное)'
                    name += ' (Очно-заочное)'


                ws.write(row, 0, id,
                         create_format(wb, halign='left', font_color=color, italic=italic, indent=0.1))
                ws.write(row, 1, name,
                         create_format(wb, halign='left', font_color=color, italic=italic, indent=0.1))
                row += 1
                id += 1
    return row
