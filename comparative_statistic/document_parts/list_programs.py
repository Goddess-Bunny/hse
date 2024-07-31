from aux_program_scripts.xlsxtools import create_format


def list_programs(wb, ws, programs, stars=[], stars_start=1, row=4):
    formats = programs['format'].unique()
    majors = programs['major'].unique()

    formats_rows, majors_rows, names_rows = {}, {}, {}

    ws.write(row, 0, 'Всего', create_format(wb, font_size=14, bold=True))
    row += 1

    for _format in formats:
        formats_rows[_format] = row
        ws.write(row, 0, _format,
                 create_format(wb, bg_color='#999999', font_size=14, bold=True))
        row += 1

        for major in sorted(majors, key=lambda x: int(x[:2]) * 10e4 + int(x[3:5]) * 10e2 + int(x[6:8])):
            names = programs[(programs['format'] == _format) &
                             (programs['major'] == major)]['name']

            if len(names) == 0:
                continue

            lvl = programs[(programs['format'] == _format) &
                           (programs['major'] == major)]['lvl'].iloc[0]

            major_type = 'Направление подготовки '
            if 'специалитет' in lvl.lower():
                major_type = 'Специальность '

            majors_rows[major] = row
            ws.write(row, 0, major_type + major,
                     create_format(wb, bg_color='#CCCCCC', bold=True, halign='left', indent=1))
            row += 1

            for name in names:
                color, italic, num_stars = None, False, 0

                program = programs[
                    (programs['format'] == _format) & (programs['major'] == major) & (programs['name'] == name)]

                for num, star in enumerate(stars):
                    if name in star:
                        num_stars = num + stars_start

                if program['new'].iloc[0]:
                    color = '#660099'
                elif program['terminated'].iloc[0]:
                    color = '#0066CC'

                if program['paid_only'].iloc[0]:
                    italic = True

                if name.find('/') != -1:
                    name = name[:name.find('/')].strip()

                name += '*'*num_stars

                if program['online'].iloc[0]:
                    name += ' (онлайн)'

                names_rows[name] = row
                ws.write(row, 0, name,
                         create_format(wb, halign='left', font_color=color, italic=italic, indent=1))
                row += 1

    row_indices = {'formats': formats_rows, 'majors': majors_rows, 'names': names_rows}
    return row
