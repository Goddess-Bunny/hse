import pandas as pd

mag = pd.read_excel('/Users/s/Documents/анна/mag.xlsx')
mag_benefit = pd.read_excel('/Users/s/Documents/анна/mag_all_adm.xlsx')

mag = mag.set_index('Рег. номер').join(mag_benefit.set_index('Рег. номер'), how='outer', rsuffix='_benefit')

mag.to_excel('/Users/s/Documents/анна/абитуриенты.xlsx')