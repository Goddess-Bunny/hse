from aux_program_scripts.xlsxtools import RowWriter, create_format


def write_programs(wb, ws, programs, cols, program_variables, start_row=5, old=True,
                   total=None, format_df=None, major_df=None, programs_only=False):
    elem = [create_format(wb), create_format(wb, bold=True, bg_color='#CCCCCC'),
            create_format(wb, bold=True, font_size=14, bg_color='#999999'), create_format(wb, bold=True, font_size=14)]
    elem_old = [create_format(wb, font_color='#0066CC'),
                create_format(wb, font_color='#0066CC', bold=True, bg_color='#CCCCCC'),
                create_format(wb, font_color='#0066CC', bold=True, font_size=14, bg_color='#999999'),
                create_format(wb, font_color='#0066CC', bold=True, font_size=14)]

    all_variables = program_variables + ['format', 'major', 'name', 'id']
    if old:
        all_variables += ['id_old']
    progs = programs[all_variables].map(lambda x: 0 if x == '-' else x)

    style = [[e_old if 'old' in var else e for var in program_variables] for e, e_old in zip(elem, elem_old)]

    formats = programs['format'].unique()
    majors = programs['major'].unique()

    if not programs_only:
        if total is None:
            total = progs[program_variables].sum()
        if format_df is None:
            format_df = progs.groupby(['format']).sum()
        if major_df is None:
            major_df = progs.groupby(['format', 'major']).sum()

    # writing totals
    for i, col, var in zip(range(len(cols)), cols, program_variables):
        row = start_row

        if not programs_only:
            ws.write(row, col, total.loc[var], style[3][i])
            row += 1

        for _format in formats:
            if not programs_only:
                try:
                    ws.write(row, col,  format_df.loc[_format, var], style[2][i])
                except KeyError:
                    print(f'No {_format}')
                    ws.write(row, col, 0, style[2][i])

                row += 1

            for major in sorted(majors, key=lambda x: int(x[:2]) * 10e4 + int(x[3:5]) * 10e2 + int(x[6:8])):
                ids = progs[(programs['format'] == _format) &
                              (programs['major'] == major)]['id']

                if len(ids) == 0:
                    continue

                if not programs_only:
                    try:
                        ws.write(row, col, major_df.loc[(_format, major), var], style[1][i])
                    except KeyError:
                        print(f'No {_format}, {major}')
                        ws.write(row, col, 0, style[1][i])
                    row += 1

                processed_ids = []

                for id_ in ids:
                    if id_ in processed_ids:
                        continue

                    program = programs[
                        (programs['format'] == _format) & (programs['major'] == major) & (programs['id'] == id_)]

                    if programs[programs['id_old'] == program['id_old'].iloc[0]].shape[0] > 1 and old:
                        unified_programs = programs[programs['id_old'] == program['id_old'].iloc[0]]

                        processed_ids.extend(unified_programs['id'].unique())
                        if 'old' in var:
                            RowWriter(wb, ws, row=row, col=col, height=unified_programs.shape[0]).add_cols(
                                list(unified_programs[var]), [elem_old[0]], [1])
                        else:
                            for j in range(unified_programs.shape[0]):
                                RowWriter(wb, ws, row=row+j, col=col).add_cols(
                                    [unified_programs[var].iloc[j]], [elem[0]], [1])

                        row += unified_programs.shape[0]
                    else:
                        ws.write(row, col, program[var].iloc[0], style[0][i])
                        row += 1
