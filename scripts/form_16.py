import os
import pandas as pd
import datetime as dt


class GottenDataframes:
    """ get dataframe as source df and let filtrating it by insert only filter value  """
    def __init__(self, dataframe):
        self.dataframe = dataframe
        dataframe.reset_index(drop=True, inplace=True)

    def filter_man(self):
        """ create new Dataframe with only Man  """
        df_copy_man = self.dataframe[self.dataframe['Пол'].str.contains('^М')]
        df_copy_man.reset_index(drop=True, inplace=True)
        return df_copy_man

    def filter_woman(self):
        """ create new Dataframe with only Woman  """
        df_copy_woman = self.dataframe[self.dataframe['Пол'].str.contains('^Ж')]
        df_copy_woman.reset_index(drop=True, inplace=True)
        # df_copy_sex = df_copy_sex.value_counts('Пол')  # grouping and counting values
        return df_copy_woman

    def man_amount(self):
        """ total counting value by Man gender """
        ttl_man = GottenDataframes.filter_man(self)
        return ttl_man.value_counts('Пол')[0]

    def woman_amount(self):
        """ total counting value by Woman gender """
        ttl_woman = GottenDataframes.filter_woman(self)
        return ttl_woman.value_counts('Пол')[0]


notif_script = os.environ['USERPROFILE'] + r'\Desktop\OSnotification.ps1'
tabs_dir_name = os.environ['USERPROFILE'] + r'\Desktop\final_tab(anton)'
fn = os.environ['USERPROFILE'] + r'\Desktop\nums.xlsx'
# pd.set_option('display.max_columns', None)

try:
    os.mkdir(tabs_dir_name)
    tabs = os.listdir(tabs_dir_name)
except FileExistsError:
    tabs = os.listdir(tabs_dir_name)
finally:
    prep_tabs = [pd.read_excel(tabs_dir_name + fr'\{tab}') for tab in tabs]


def concat_tabs():
    try:
        final_tab = pd.concat(prep_tabs).drop(columns=['Номер ЭЛН', 'Дата выдачи', 'ФИО нетрудоспособного',
                                                       'СНИЛС', 'Причина нетрудоспособности',
                                                       'Период нетрудоспособности', 'ФИО врача закрывшего ЭЛН',
                                                       'Статус', 'Изменен', 'Статус СФР'])
        final_tab.dropna(axis='columns', how='all', inplace=True)
        final_tab['Дата рождения'] = final_tab['Дата рождения'].astype(dtype='datetime64[ns]')
        return final_tab
    except ValueError:
        os.system(notif_script)
        os.system('powershell kill -name python')


def calculate_age(birth_date):  # convert datetime to integer number - year
    today = dt.datetime.now()
    age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
    return age


df = concat_tabs()
df.reset_index(drop=True, inplace=True)
# df = pd.read_excel(tab_file, dtype={'Дата рождения': 'datetime64[ns]'})  # read file and create dataframe
df.loc[:, "Возраст"] = [calculate_age(df.at[i, 'Дата рождения']) for i
                        in range(len(df['Дата рождения']))]  # creating new column and applying function getting age
df = df.astype(dtype={'Возраст': 'str'})  # converting column data type
df = df[df['Статус ФСС'] == '030-Закрыт']  # filter by col
ready_DF = GottenDataframes(df)  # creating class object
diagnosis = ['A00 - B99', 'A15 - A19', 'C00 - D48', 'C00 - D09', 'D50 - D89', 'E00 - E90', 'E10 - E14', 'F00 - F99',
             'G00 - G99', 'H00 - H59', 'H60 - H95', 'I00 - I99', 'I20 - I25', 'I60 - I69', 'J00 - J99',
             'J00, J01, J04, J05, J06', 'J10 - J11', 'J12 - J18', 'K00 - K93', 'L00 - L99', 'M00 - M99', 'N00 - N99',
             'O00 - O99', 'Q00 - Q99', 'S00 - T98', 'U07', 'O03 - O07', 'Z00 - Z99']
ages = ['15-19', '20-24', '25-29', '30-34', '35-39', '40-44', '45-49', '50-54', '55-59', '60+']

# common_value = ready_DF.filter_man().value_counts('Пол').values[0]   # count total numbers of mans
# common_value = ready_DF.filter_woman().value_counts('Пол').values[0]   # count total numbers of womans


def diag_block_woman():  # counting sex by diag block
    filters = {1: '^A|^B', 2: '^A1[56789]', 3: '^C|^D[01234][012345678]', 4: '^C|^D[0123456789]',
               5: '^D[5678]', 6: '^E', 7: '^E1[01234]', 8: '^F', 9: '^G', 10: '^H[012345]',
               11: '^H[6789]', 12: '^I', 13: '^I2[012345]', 14: '^I6', 15: '^J', 16: '^J0[01456]',17: '^J1[01]',
               18: '^J1[2345678]', 19: '^K', 20: '^L', 21: '^M', 22: '^N', 23: '^O', 24: '^Q',
               25: '^S|^T', 26: '^U07', 27: '^O0[34567]', 28: '^Z'
               }
    woman = ready_DF.filter_woman()
    diag_blocks_dict = {i: woman[woman['Диагноз'].str.contains(v)] for (i, v) in filters.items()}
    return diag_blocks_dict


def diag_block_man():  # counting sex by diag block
    filters = {1: '^A|^B', 2: '^A1[56789]', 3: '^C|^D[01234][012345678]', 4: '^C|^D[0123456789]',
               5: '^D[5678]', 6: '^E', 7: '^E1[01234]', 8: '^F', 9: '^G', 10: '^H[012345]',
               11: '^H[6789]', 12: '^I', 13: '^I2[012345]', 14: '^I6', 15: '^J', 16: '^J0[01456]', 17: '^J1[01]',
               18: '^J1[2345678]', 19: '^K', 20: '^L', 21: '^M', 22: '^N', 23: '^O', 24: '^Q',
               25: '^S|^T', 26: '^U07', 27: '^O0[34567]', 28: '^Z'
               }
    man = ready_DF.filter_man()
    diag_blocks_dict = {i: man[man['Диагноз'].str.contains(v)] for (i, v) in filters.items()}
    return diag_blocks_dict


by_axies0_woman = {k: (v.value_counts('Пол')[0] if not v.empty else 0) for k, v in diag_block_woman().items()}
by_axies0_man = {k: (v.value_counts('Пол')[0] if not v.empty else 0) for k, v in diag_block_man().items()}


def age_groups_man():  # creating dict that contains counted amount of mans separated by age groups
    filters = {1: '^1[56789]',
               2: '^2[01234]',
               3: '^2[56789]',
               4: '^3[01234]',
               5: '^3[56789]',
               6: '^4[01234]',
               7: '^4[56789]',
               8: '^5[01234]',
               9: '^5[56789]',
               10: '^[6789][0123456789]'
               }
    numbers = {k: {i: (v[v['Возраст'].str.contains(l)].value_counts('Пол')[0]
               if not v[v['Возраст'].str.contains(l)].value_counts('Пол').empty else 0)
               for (i, l) in filters.items()} for (k, v) in diag_block_man().items()}
    return numbers


def age_groups_woman():  # creating dict that contains counted amount of womans separated by age groups
    filters = {1: '^1[56789]',
               2: '^2[01234]',
               3: '^2[56789]',
               4: '^3[01234]',
               5: '^3[56789]',
               6: '^4[01234]',
               7: '^4[56789]',
               8: '^5[01234]',
               9: '^5[56789]',
               10: '^[6789][0123456789]'
               }
    numbers = {k: {i: (v[v['Возраст'].str.contains(l)].value_counts('Пол')[0]
               if not v[v['Возраст'].str.contains(l)].value_counts('Пол').empty else 0)
               for (i, l) in filters.items()} for (k, v) in diag_block_woman().items()}
    return numbers


mans = pd.DataFrame.from_dict(age_groups_man(), orient='index')
womans = pd.DataFrame.from_dict(age_groups_woman(), orient='index')
mans.rename(columns={ages.index(i)+1: i for i in ages},
            index={diagnosis.index(i)+1: i for i in diagnosis},
            inplace=True, errors='raise')
womans.rename(columns={ages.index(i)+1: i for i in ages},
              index={diagnosis.index(i)+1: i for i in diagnosis},
              inplace=True, errors='raise')

with pd.ExcelWriter(fn, mode='w', engine='openpyxl') as f:
    mans.to_excel(f, sheet_name='mans', header=True, index_label='Diagnosis')
    womans.to_excel(f, sheet_name='womans', header=True, index_label='Diagnosis')

