from aux_program_scripts.xlsxtools import RowWriter, create_format, process_day
from datetime import datetime
import pandas as pd


def create_header(wb, ws, lvl, CUR_DATE, num_days, num_days_old, from_start=False, PAID_ONLY=False):
    header_elem = create_format(wb, bold=True)
    title_elem = create_format(wb, bold=True, font_size=16, border=False)
    header_elem_old = create_format(wb, font_color='#0066CC', rotation=90)

    columns = {}
    head = RowWriter(wb, ws, height=2)

    # adding the 'Образовательная программа' column
    names = ['name']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    head.add_cols(['Образовательная программа'], [header_elem], [1])
    info_date_part = 'Прием документов завершился'
    if not PAID_ONLY:
        # adding the informational part (number of places today vs last year)
        info = [["Количество мест в рамках контрольных цифр приема"],
                ["КЦП в 2024 году", "Места за счет средств НИУ ВШЭ в 2024 году",
                 "Справочно: КЦП в 2023 году", "Справочно: места за счет средств НИУ ВШЭ в 2023 году",
                 "в т.ч. места по целевой квоте в 2024 году", "Справочно: целевая квота в 2023 году"]]

        names = ['kcp', 'school', 'kcp_old', 'school_old', 'work', 'work_old']

        if lvl == 0:
            info = [["Количество мест в рамках контрольных цифр приема"],
                    ["КЦП в 2024 году*", "Справочно: КЦП в 2023 году*",
                     "в т.ч. места по целевой квоте в 2024 году", "Справочно: целевая квота в 2023 году",
                     "в т.ч. места по особой квоте в 2024 году", "Справочно: особая квота в 2023 году",
                     "в т.ч. места по отдельной квоте в 2024 году", "Справочно: отдельная квота в 2023 году"]]
            names = ['kcp', 'kcp_old', 'work', 'work_old', 'special', 'special_old', 'separate', 'separate_old']
        else:
            info = [["Количество мест в рамках контрольных цифр приема"],
                    ["КЦП в 2024 году", "Справочно: КЦП в 2023 году",
                     "Места за счет средств НИУ ВШЭ в 2024 году", "Справочно: места за счет средств НИУ ВШЭ в 2023 году",
                     "в т.ч. места по целевой квоте в 2024 году", "Справочно: целевая квота в 2023 году"]]
            names = ['kcp', 'kcp_old', 'school', 'school_old', 'work', 'work_old']

        info_style = [[header_elem], [header_elem, header_elem_old]*(len(info[1])//2)]
        columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}
        info_width = [[len(info[1])], [1]*len(info[1])]

        head.add_cols(info, info_style, info_width)

        # adding number of admits information
        #info_date_part = f'за {num_days_old}' + process_day(num_days_old) + ' с начала приема'
        #if not from_start:
        #    info_date_part = f'за {num_days_old} ' + process_day(num_days_old) + ' до конца приема'
        info_date_part = 'Прием документов завершился'

        nums = ['Количество заявлений на места в рамках КЦП 2024',
                f'Справочно: заявлений на места\nв рамках КЦП 2023 ({info_date_part})',
                'в т.ч. количество заявлений на места по целевой квоте 2024',
                f'Справочно: заявлений по целевой\nквоте 2023 ({info_date_part})'
                ]
        names = ['kcp_admit', 'kcp_admit_old', 'work_admit', 'work_admit_old']

        if lvl == 0:
            nums += ['в т.ч. количество заявлений на места по особой квоте 2024',
                     f'Справочно: заявлений по особой квоте 2023\n({info_date_part})',
                     'в т.ч. количество заявлений на места по отдельной квоте 2024',
                     f'Справочно: заявлений по отдельной квоте 2023\n({info_date_part})',
                     'в т.ч. количество заявлений от дипломантов олимпиад, имеющих право на прием БВИ 2024',
                     f'Справочно: заявлений на прием БВИ 2023\n({info_date_part})',
                     'в т.ч. БВИ по ВсОШ 2024',
                     f'Справочно: заявлений БВИ по ВсОШ 2023\n({info_date_part})'
                     ]
            names += ['special_admit', 'special_admit_old', 'separate_admit', 'separate_admit_old',
                      'bvi_admit', 'bvi_admit_old', 'vsosh_admit', 'vsosh_admit_old']

        nums_style = [header_elem if i % 2 == 0 else header_elem_old for i in range(len(nums))]
        nums_width = [1 for i in range(len(nums))]

        columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}
        head.add_cols(nums, nums_style, nums_width)

        # adding olimps information
        if lvl == 2:
            olimps = [["Олимпиады и \"Раннее приглашение\""], [
                "Потенциальное кол-во абитуриентов из числа включенных в приказ на раннее приглашение в московском кампусе",
                "Кол-во заявлений по раннему приглашению",
                "Потенциальное кол-во абитуриентов из числа дипломантов 2024 года олимпиады \"Высшая лига\"*",
                'Кол-во заявлений от абитуриентов, заявивших льготу по олимпиаде \"Высшая лига\"',
                'Потенциальное кол-во абитуриентов из числа дипломантов 2024 года олимпиады\"Я - профессионал\"*',
                'Кол-во заявлений от абитуриентов, заявивших льготу по олимпиаде \"Я - профессионал\"',
                'Кол-во заявлений от абитуриентов, заявивших другие олимпиады и конкурсы в качестве льготы'
            ]]
            names = ['early_potential', 'early_admit', 'vl_potential', 'vl_admit', 'yaprofi_potential', 'yaprofi_admit',
                     'other_admit']
            columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

            olimps_style = [[header_elem], [header_elem]*len(olimps[1])]
            olimps_width = [[len(olimps[1])], [1]*len(olimps[1])]

            head.add_cols(olimps, olimps_style, olimps_width)

        # inserting a green line
        columns |= {'green': head.cur_col}
        head.insert_green()

    # adding information on paid places
    paid = ['Количество мест с оплатой стоимости обучения на договорной основе 2024',
            'Справочно: кол-во платных мест 2023\n(на конец приема документов)',
            'Количество заявлений на места с оплатой стоимости обучения на договорной основе 2024',
            f'Справочно: заявлений на платные места 2023\n({info_date_part})']
    names = ['paid', 'paid_old', 'paid_admit', 'paid_admit_old']
    columns |= {col_name: head.cur_col + i for i, col_name in enumerate(names)}

    paid_style = [header_elem if i % 2 == 0 else header_elem_old for i in range(len(paid))]
    paid_width = [1 for i in range(len(paid))]

    head.add_cols(paid, paid_style, paid_width)

    # writing a title for the file
    title = RowWriter(wb, ws, height=2, row=0)

    if lvl == 0:
        admission_lvl = 'бакалавриат'
    else:
        admission_lvl = 'магистратуру'

    #cur_info_date_part = f'за {num_days}' + process_day(num_days) + ' с начала приема'
    #if not from_start:
    #    cur_info_date_part = f'за {num_days} ' + process_day(num_days) + ' до конца приема'
    cur_info_date_part = 'прием документов завершился'

    if not PAID_ONLY:
        title_width = columns['green']

        title.add_cols([['Сравнительная статистика количества поданных заявлений в ' + admission_lvl],
                        [f'на {CUR_DATE.strftime('%d.%m.%Y %H:%M')} ({cur_info_date_part})']],
                       [[title_elem], [title_elem]], [[title_width], [title_width]])

    return head.cur_col - 1, pd.Series(columns)
