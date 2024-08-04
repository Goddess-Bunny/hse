from aux_program_scripts.xlsxtools import RowWriter, create_format, process_day
from datetime import datetime
import pandas as pd


def create_header(wb, ws, CAMPUS):
    header_elem_shortline = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10)
    header_elem_longline = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10, rotation=90)
    header_elem_longline_green = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10, rotation=90,
                                         font_color='#00BE00')
    header_elem_longline_yellow = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10, rotation=90,
                                         font_color='#FFA500')

    columns = {}
    head = RowWriter(wb, ws, height=2)

    head.add_cols(['№', 'Конкурс'], [header_elem_shortline]*2, [1]*2)
    names = ['id', 'name']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    # adding the informational part (number of places)

    info = ["КЦП", "в т.ч. целевая квота", 'целевая квота с учетом предложений от заказчиков',
            "в т.ч. особая квота", "в т.ч. отдельная квота"]
    names = ['kcp', 'work', 'work_max', 'special', 'separate']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_shortline] + [header_elem_longline]*(len(names) - 1)
    info_width = [1]*len(names)

    head.add_cols(info, info_style, info_width)

    # adding the number of admits

    admits = [["Количество заявлений"],
              ["БВИ", "целевая квота", "особая квота", "отдельная квота", "на места в рамках общего конкурса"]]
    names = ['bvi_admit', 'work_admit', 'special_admit', 'separate_admit', 'kcp_admit']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_shortline], [header_elem_longline]*len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(admits, info_style, info_width)

    # adding the number of admits who submitted original documents

    original_admits = ['Количество абитуриентов с оригиналами',
                       'Количество абитуриентов с оригиналами с первым приоритетом']
    names = ['original_admit', 'original_admit_priority']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_longline] * len(names)
    info_width = [1] * len(names)

    head.add_cols(original_admits, info_style, info_width)

    # adding the number of admitted special

    admitted = [["в т.ч. количество имеющихся оригиналов с первым приоритетом"],
              ["БВИ", "целевая квота", "особая квота", "отдельная квота"]]
    names = ['bvi_admitted', 'work_admitted', 'special_admitted', 'separate_admitted']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_shortline], [header_elem_longline]*len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(admitted, info_style, info_width)

    # adding the number of leftover kcp places

    leftover = ['Кол-во мест на общий конкурс', '25% от КЦП']
    names = ['leftover_admit', '25']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_longline]*2
    info_width = [1]*2

    head.add_cols(leftover, info_style, info_width)

    # adding waves information

    waves = ['Предложение пр-мы 24.07 по баллу "зеленой волны" 2024',
             'Балл "зеленой волны" в 2023, утвержденный на заседании',
             'Фактический проходной балл в 2023 году',
             'Кол-во аб-тов в зоне "зеленой волны", поступающих на конкурсные места',
             'Количество оригиналов с первым приоритетом в зоне "зеленой волны"',
             'Предложение пр-мы 24.07 по баллу "листа ожидания" 2024',
             'Балл "желтой волны" в 2023, утвержденный на заседании',
             'Кол-во аб-тов в зоне "листа ожидания", поступающих на конкурсные места',
             'Количество оригиналов с первым приоритетом в зоне "желтой волны"']
    names = ['green_num', 'green_num_old', 'realized_min_score', 'ingreen_admit', 'ingreen_admit_priority',
             'yellow_num', 'yellow_num_old', 'inyellow_admit', 'inyellow_admit_priority']
    if CAMPUS != 'Москва':
        waves = waves[:2] + waves[3:]
        names = names[:2] + names[3:]
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = ([header_elem_longline_green]*3 + [header_elem_longline]*2 + [header_elem_longline_yellow]*2 +
                  [header_elem_longline]*2)
    info_width = [1]*len(names)

    head.add_cols(waves, info_style, info_width)

    # adding old min_score/wave information

    waves_old = [['2022', '2021'], ['Фактический балл зачисления в 2022 году', 'Балл "ЗВ" 2022 первоначальный',
                                    'Фактический балл зачисления в 2021 году', 'Балл "ЗВ" 2021 первоначальный']]
    names = ['realized_min_score_old2022', 'green_num_old2022', 'realized_min_score_old2021', 'green_num_old2021']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_shortline]*2, [header_elem_longline]*len(names)]
    info_width = [[2, 2], [1] * len(names)]

    head.add_cols(waves_old, info_style, info_width)

    # adding the decided wave scores

    # waves_final = [['Решения по итогам совещания 28.07.2024'], ['Балл "зеленой волны"', 'Балл "желтой волны"']]
    # names = ['green_final', 'yellow_final']
    # columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}
    #
    # info_style = [[header_elem_shortline], [header_elem_longline_green, header_elem_longline_yellow]]
    # info_width = [[2], [1] * len(names)]
    #
    # head.add_cols(waves_final, info_style, info_width)

    # green line
    columns |= {'green': head.cur_col}
    head.insert_green()

    # adding paid places

    paid = ['Плановое количество платных мест 2023', 'Количество заключенных договоров*',
            'Количество оплаченных договоров*', 'Критерии заключения договора']
    names = ['paid', 'paid_signed', 'paid_paid', 'paid_min_scores']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_longline]*len(names)
    info_width = [1] * len(names)

    head.add_cols(paid, info_style, info_width)

    return head.cur_col - 1, pd.Series(columns)


def create_header_original(wb, ws, CAMPUS):
    header_elem_shortline = create_format(wb, font_name='Times New Roman', font_size=11)
    header_elem_longline = create_format(wb, bold=True, font_name='Times New Roman', font_size=11)
    header_elem_longline_green = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10, rotation=90,
                                               font_color='#00BE00')
    header_elem_longline_yellow = create_format(wb, bold=True, font_name='Arial Cyr', font_size=10, rotation=90,
                                                font_color='#FFA500')

    columns = {}
    head = RowWriter(wb, ws, height=2)

    head.add_cols(['№', 'Конкурс'], [header_elem_shortline] * 2, [1] * 2)
    names = ['id', 'name']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    # adding the informational part (number of places)

    info = [['Количество мест в рамках контрольных цифр приема'], ["Всего мест КЦП", "в т.ч. целевая квота",
                                                                   "в т.ч. особая квота", "в т.ч. отдельная квота"]]
    names = ['kcp', 'work', 'special', 'separate']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_longline], [header_elem_shortline] * len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(info, info_style, info_width)

    # adding the number of admits

    admits = [[
                  "Количество заявлений от абитуриентов, которые прошли минимальный порог и попали в конкурсные списки на места в рамках КЦП"],
              ["Всего заявлений на места КЦП", "в т.ч. БВИ", "в т.ч. целевая квота", "в т.ч. особая квота",
               "в т.ч. отдельная квота ", 'в т.ч. на места в рамках общего конкурса']]
    names = ['kcp_admit', 'bvi_admit', 'work_admit', 'special_admit', 'separate_admit', 'kcp_leftover_admit']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_longline], [header_elem_shortline] * len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(admits, info_style, info_width)

    # adding the number of admits who submitted original documents and is surely admitted

    admits = [["Количество представленных к зачислению в приоритетный этап"],
              ["БВИ", "целевая квота", "особая квота", "отдельная квота"]]
    names = ['bvi_admitted', 'work_admitted', 'special_admitted', 'separate_admitted']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_longline], [header_elem_shortline] * len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(admits, info_style, info_width)

    # adding the number of leftover kcp places

    leftover = ['Кол-во мест на общий конкурс', '25% от КЦП']
    names = ['leftover_admit', '25']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_longline] * 2
    info_width = [1] * 2

    head.add_cols(leftover, info_style, info_width)

    # adding green wave information

    waves = [['«зеленая волна»'], ['Балл «зеленой волны»', 'Количество абитуриентов, вошедших в список «зеленой волны»',
                                   'Количество оригиналов с первым приоритетом в списке «зеленой волны»',
                                   'Количество оригиналов со вторым приоритетом в списке «зеленой волны»',
                                   'Количество оригиналов с третьим приоритетом в списке «зеленой волны»']]
    names = ['green_num', 'ingreen_admit', 'ingreen_admit_priority1',
             'ingreen_admit_priority2', 'ingreen_admit_priority3']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_longline], [header_elem_shortline] * len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(waves, info_style, info_width)

    # adding the number of priority leftover kcp places

    leftover = ['Осталось мест КЦП для заполнения (учитываются только оригиналы первого приоритета)',
                'Кол-во абитуриентов сверх мест КЦП (перебор)']
    names = ['leftover_admit_priority', 'overdraw']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_longline] * 2
    info_width = [1] * 2

    head.add_cols(leftover, info_style, info_width)

    # adding yellow wave information

    waves = [['«желтая волна»'], ['Балл «желтой волны»', 'Количество абитуриентов, вошедших в список «желтой волны»',
                                  'Количество оригиналов с первым приоритетом в списке «желтой волны»']]
    names = ['yellow_num', 'inyellow_admit', 'inyellow_admit_priority1']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [[header_elem_longline], [header_elem_shortline] * len(names)]
    info_width = [[len(names)], [1] * len(names)]

    head.add_cols(waves, info_style, info_width)

    # adding

    waves = ['Балл фактического закрытия бюджетных мест', 'Количество зачисленных на основные конкурсные места',
             'Кол-во абитуриентов, которые могли бы претендовать на зачисление по зеленой волне за счет НИУ ВШЭ',
             'Кол-во абитуриентов БВИ сверх мест КЦП - места за счет НИУ ВШЭ']
    names = ['admit_score', 'admit_number', 'possible_admit', 'possible_bvi']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    info_style = [header_elem_shortline] * len(names)
    info_width = [1] * len(names)

    head.add_cols(waves, info_style, info_width)


    return head.cur_col - 1, pd.Series(columns)
