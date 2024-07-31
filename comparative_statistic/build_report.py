import pandas as pd
import xlsxwriter
from comparative_statistic.document_parts import (create_header, list_programs, write_old_admission,
                                                  write_yaprofi_potential,
                                                  write_vl_potential, write_asav_numbers, write_aicpk_numbers,
                                                  write_rp_potential,
                                                  write_olimp_admits, write_olimp_vsosh)
from aux_program_scripts.xlsxtools import process_day
from aux_program_scripts import write_programs, RowWriter, create_format, join_new_old_progs
from datetime import datetime, date, timedelta
import os
BAC, MAG = 0, 2
LVL = MAG

campuses = ['Москва', 'Санкт-Петербург', 'Пермь', 'Нижний Новгород']
CAMPUS = campuses[0]
CUR_DATE = datetime.today()

# ПОСТАВИТЬ СВОИ ПУТИ СЮДА
pathToOutputFolder = "/Users/s/Desktop/Admission2024/"
pathOfInputsFolder = '/Users/s/Desktop/Admission Numbers 2024/'

pathStat = pathToOutputFolder + f"{CAMPUS}/{CUR_DATE.strftime('%Y-%m-%d')}/"

try:
    os.mkdir(pathStat)
except FileExistsError:
    pass

if LVL >= MAG:
    programs = pd.read_excel(pathOfInputsFolder + 'programs.xlsx', sheet_name='program')
else:
    programs = pd.read_excel('/Users/s/Desktop/Admission Numbers 2024/programs_bac.xlsx', sheet_name='program')

wb = xlsxwriter.Workbook(pathStat + "workbook.xlsx")
ws = wb.add_worksheet('Сравнение')

programs = programs[programs['campus'] == CAMPUS]

# WRITE CODE HERE
# --------------

# getting old admission info (dates and total number of students)
days_till_end, cur_days_till_end, prev_date, old_tot_inc_students = write_old_admission(wb, ws, programs, LVL, CAMPUS,
                                                                                        CUR_DATE, info=True)

# creating header
num_col, col_correspondence = create_header(wb, ws, LVL, CUR_DATE, num_days=cur_days_till_end,
                                            num_days_old=days_till_end)
# writing down all programs

starred_programs, stars_start = [], 1

if LVL == BAC and CAMPUS == 'Москва':
    starred_programs.extend([['Актёр', 'Кинопроизводство'], ['Журналистика'], ['Дизайн', "Мода"]])
    stars_end_dates = [date(year=2024, month=7, day=8), date(year=2024, month=7, day=15),
                       date(year=2024, month=7, day=18)]
    stars_end_dates_prev = [date(year=2023, month=7, day=8), date(year=2023, month=7, day=8),
                            date(year=2023, month=7, day=18)]
    stars_start = 2

num_row = list_programs(wb, ws, programs, stars=starred_programs, stars_start=stars_start, row=5)

# writing old admission numbers
write_old_admission(wb, ws, programs, LVL, CAMPUS, CUR_DATE, col_correspondence=col_correspondence)

# drawing green line separating free and paid places
for i in range(5, num_row):
    RowWriter(wb, ws, row=i, col=col_correspondence['green']).insert_green()

# writing the number of places at each program for this year
if LVL == BAC:
    program_variables = ['kcp', 'work', 'special', 'separate']
if LVL >= MAG:
    program_variables = ['kcp', 'school', 'work']

cols = col_correspondence[program_variables]

write_programs(wb, ws, programs, cols, program_variables)

# reading the asav/aic-pk file
if LVL >= MAG:
    total_inc_students = write_asav_numbers(wb, ws, programs, col_correspondence, LVL, CAMPUS, CUR_DATE)
else:
    total_inc_students = write_aicpk_numbers(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE)

# processing the preferences/olympiad information
if LVL == MAG:
    admissible = {}
    admissible['yaprofi'] = write_yaprofi_potential(wb, ws, programs, col_correspondence, CAMPUS)
    admissible['vl'] = write_vl_potential(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE)
    admissible['early'] = write_rp_potential(wb, ws, programs, col_correspondence, CAMPUS)

    write_olimp_admits(wb, ws, programs, col_correspondence, admissible, CAMPUS, CUR_DATE)
if LVL == BAC:
    write_olimp_vsosh(wb, ws, programs, col_correspondence, CAMPUS, CUR_DATE)

# writing the total number of prospective students
RowWriter(wb, ws, row=num_row).add_cols(
    [f"Количество абитуриентов на {CUR_DATE.strftime('%d.%m.%Y')}" +
     #f"\n(за {cur_days_till_end} {process_day(cur_days_till_end)} до конца приема):", total_inc_students],
    "\n(Прием документов завершился)", total_inc_students],
    [create_format(wb, bold=True, halign='left', indent=1)] * 2, [1] * 2)
RowWriter(wb, ws, row=num_row + 1).add_cols(
    [f"Количество абитуриентов на {prev_date.strftime('%d.%m.%Y')}:" +
     #f"\n(за {days_till_end} {process_day(days_till_end)} до конца приема):", old_tot_inc_students],
     "\n(Прием документов завершился)", old_tot_inc_students],
    [create_format(wb, bold=True, halign='left', font_color='#0066CC', indent=1)] * 2, [1] * 2)

if LVL == MAG:
    ws.write(num_row + 3, 0, '* общие суммы и суммы по направлениям подготовки\n' +
             'считались как кол-во уникальных абитуриентов. Учитываются только дипломанты-выпускники 2024 года',
             create_format(wb, italic=True, indent=1, border=False, halign='left', text_wrap=False))
    num_row += 2

# writing comments
if LVL == BAC and CAMPUS == 'Москва':
    ws.write(num_row + 3, 0, "* в 2024 году выделено 5 мест за счет средств НИУ ВШЭ для общего конкурса по программе" +
             " \"Актёр\";в 2023 году не было выделено мест" +
             " за счет средств НИУ ВШЭ для общего конкурса",
             create_format(wb, italic=True, border=False, halign='left', indent=1, text_wrap=False))

    star_text = ["** для ОП «Кинопроизводство» и «Актер» сравнение с ", "*** для ОП «Журналистика» сравнение с ",
                 "**** для ОП «Дизайн» и «Мода» сравнение с "]

    for i, text in enumerate(star_text):
        if prev_date < stars_end_dates_prev[i]:
            text += (f"{prev_date.strftime('%d.%m.%Y')}" + f" (за {(stars_end_dates_prev[i] - prev_date).days}" +
                     f" {process_day((stars_end_dates_prev[i] - prev_date).days)} до" +
                     f" окончания приема документов")
        else:
            text += f"{stars_end_dates_prev[i].strftime('%d.%m.%Y')}" + " (в 2023 году прием документов на эту дату закончился"

        if CUR_DATE.date() < stars_end_dates[i]:
            text += (f", на текущий день осталось {(stars_end_dates[i] - CUR_DATE.date()).days}"
                     f" {process_day((stars_end_dates[i] - CUR_DATE.date()).days)} до конца приема)")
        else:
            text += (f", на текущий день прием документов завершился)")

        ws.write(num_row + 4 + i, 0, text,
                 create_format(wb, italic=True, border=False, halign='left', indent=1, text_wrap=False))

    num_row += len(star_text) + 2

ws.write(num_row + 3, 0, 'Выделения образовательных программ:',
         create_format(wb, bold=True, halign='left', border=False, indent=1))
ws.write(num_row + 4, 0, 'Новые образовательные программы',
         create_format(wb, font_color='#660099', halign='left', border=False, indent=1))
ws.write(num_row + 5, 0, 'Полностью платные образовательные программы',
         create_format(wb, italic=True, halign='left', border=False, indent=1))
ws.write(num_row + 6, 0, 'Образовательные программы 2023 года, не осуществляющие набор в 2024 году',
         create_format(wb, font_color='#0066CC', halign='left', border=False, indent=1))

# setting up print page
ws.set_zoom(80)
ws.autofilter(4, 0, num_row, num_col)
ws.freeze_panes(5, 1)
ws.set_landscape()
ws.repeat_rows(first_row=3, last_row=4)
ws.set_column(0, 0, 57)
ws.set_paper(8)
ws.set_margins(top=0.4, bottom=0.4, left=0.1, right=0.1)
ws.fit_to_pages(width=1, height=0)
ws.set_footer('&RСтраница &P из &N')
ws.set_column(1, col_correspondence['green'] - 1, 12.17)
ws.set_column(col_correspondence['green'], col_correspondence['green'], 2)
ws.set_column(col_correspondence['green'] + 1, 30, 12.17)
ws.set_row(4, 250)
ws.set_row(3, 30)

wb.close()

# rename files

if LVL == BAC:
    os.rename(pathStat + "workbook.xlsx",
              pathStat +
              f"Сравнительная_статистика_бакалавриат_{CUR_DATE.strftime('%d_%m_%Y')}_с_" +
              f"{prev_date.strftime('%d_%m_%Y')}.xlsx")
if LVL >= MAG:
    os.rename(pathStat + "workbook.xlsx",
              pathStat +
              f"Сравнительная_статистика_магистратура_{CUR_DATE.strftime('%d_%m_%Y')}_с_" +
              f"{prev_date.strftime('%d_%m_%Y')}.xlsx")
