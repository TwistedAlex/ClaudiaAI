###
# Author: SHUO LI
# Requirement from C.W: 现在有文档a，和18个文档b，文档a里有个column 是tav tax number，和tax lot from lot，文档b是统一的格式，也有tav tax number。如果tax lot from lot=0， search tav tax number in 文档b，并复制所有match 的tav number所在的row到一个新的文档
# ###

import pandas as pd
import os
from os import listdir
from os.path import isfile, join

#显示所有列
pd.set_option('display.max_columns', None)
#显示所有行
pd.set_option('display.max_rows', None)
#设置value的显示长度为100，默认为50
pd.set_option('max_colwidth',100)

cwd = os.getcwd()
work_dir = input('Please enter your desired working directory: ')
os.chdir(work_dir)
lot_file = input("Please enter the directory of the LOT file: ")

# Retrieve all file dirs end with .xlsx under a few folders in a list
stop = 0
tav_list = list()
while(stop == 0):
    tav_dir = input("Please enter the directory for all TAV files, if no more files needed, please enter STOP")
    if tav_dir == 'STOP':
        stop = 1
    else:
        new_tav_list = [join(tav_dir, f) for f in listdir(tav_dir) if f.endswith(".xlsx")]
        tav_list.extend(new_tav_list)

df_tav = 0
init_df_flag = 1
for f in tav_list:
    # Read excel file via file directory
    new_tav = pd.ExcelFile(f)
    # parse excel file to dataframe. default: 1st sheet.
    df_lot = new_tav.parse()
    folder_name = f.split("\\")
    # init dataframe to hold data from files
    if init_df_flag == 1:
        df_tav = df_lot
        init_df_flag = 0
    else:
        # append new coming excel files to dataframe
        df_tav = df_tav.append(df_lot)
        df_tav.duplicated()

# output duplicated tav
duplicated_tav = df_tav[df_tav.duplicated(keep=False)]
if not duplicated_tav.empty:
    print("Duplicates exist")
    duplicated_tav_writer = pd.ExcelWriter(join(work_dir,'duplicated_tav.xlsx'), engine='xlsxwriter')
    duplicated_tav.to_excel(duplicated_tav_writer, 'Sheet1')
    duplicated_tav_writer.save()
else:
    print("Duplicates do not exist")

# Drop duplicates from tav
df_tav = df_tav.drop_duplicates()
# print(df_tav['TAV Transaction Number (15char)\n*If not supplied number will be assigned'])

lot_xl = pd.ExcelFile(lot_file)
df_lot = lot_xl.parse('Sheet1')
# only select rows with ['lot sold from lot'] == 0
lot_selected = df_lot.loc[df_lot['lot sold from lot'] == 0]

# a list of "TAVTran No" of selected rows in lot file
targeted_tav_set = lot_selected["TAVTran No"].tolist()
# rows, in tav file, whose TAV NO is in the selected TAV NO list
df_tav = df_tav.loc[df_tav['TAV Transaction Number (15char)\n*If not supplied number will be assigned'].isin(targeted_tav_set)]

# Format and output the results
tav_writer = pd.ExcelWriter(join(work_dir,'results.xlsx'), engine='xlsxwriter')
df_tav.to_excel(tav_writer, 'Sheet1')
worksheet = tav_writer.sheets['Sheet1']
workbook = tav_writer.book
text_num_format = workbook.add_format({'num_format': '0'})
worksheet.set_column('B:B', None, text_num_format)
worksheet.set_column('D:D', None, text_num_format)
worksheet.set_column('G:G', None, text_num_format)
tav_writer.save()
