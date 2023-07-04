import pandas as pd
import os
from os import listdir
from os.path import isfile, join, isdir
import shutil, zipfile
import Extract_and_Move_Error

def call(ban_error_file,ban_template,target):
    ban_error = pd.read_csv(ban_error_file, dtype=str, delimiter = "\t")
    ban_list = pd.read_excel(ban_template,sheet_name='New BANs')
    ban_list = ban_list.rename(columns=ban_list.iloc[0]).drop(ban_list.index[0])
    ban_error['concat'] = ban_error['bill_fccid'] + ban_error['cic_list'] + ban_error['original_cic']
    ban_list['concat'] = ban_list['FCCID'] + ban_list['CIC List'] + ban_list['CIC']
    ban_error_filter=ban_error.loc[ban_error['concat'].isin(list(ban_list['concat'].unique()))]
    ban_error_files=sorted(set(list(ban_error_filter['file_name'].str.upper())))
    file = open(r'K:\Client\Service_bureau\Staff Tools\Python\CABS ONE\Files_list.csv', 'a')
    for i in ban_error_files:
        file.write(f'{i}\n')
    file.close()
    Extract_and_Move_Error.extract_and_move(target)