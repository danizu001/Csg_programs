import pandas as pd
from os import listdir
from os.path import isfile, join
import shutil

def move_or_copy_files(path_inv,bans,mode,path_z):
    total_files=[]
    files = [name for name in listdir(path_inv) if isfile(join(path_inv, name))]
    for file in files:
        if mode=='1':
            full_name=file.split('-')
            if len(full_name)>3:
                file_ban=full_name[1]+'-'+full_name[2]+'-'+full_name[3]
                if file_ban in bans:
                    total_files.append(file_ban)
                    shutil.copyfile(path_inv+'\\'+file,path_z+'\\'+file)
            else:    
                file_ban=full_name[1]
                if file_ban in bans:
                    total_files.append(file_ban)
                    shutil.copyfile(path_inv+'\\'+file,path_z+'\\'+file)
        else:
            if file in bans:
                total_files.append(file)
                shutil.move(path_inv+'\\'+file,path_z+'\\'+file)
    return(total_files)

def call(ff_full_path,part,path_inv=r'Q:\SECABS_PA_FF\X_VERIFY_OUTPUT\Passed'):
    path_z='Q:\\SECABS_PA_FF\\X_VERIFY_OUTPUT\\OUTBOUND_EMAIL\\EMAIL\\InvoicesToEmail'
    file_df= pd.read_excel(ff_full_path)
    if part=='1':
        print("Copying files, please wait...")
        emails=file_df[file_df["format"] == 'Email-PDF']
        bans=emails['ban'].to_list()
        total_files=move_or_copy_files(path_inv,bans,'1',path_z)
        print('The total PDF files copied has been: '+str(len(total_files)))
    elif part=='2':
        print("Moving files, please wait...")
        emails=file_df[(file_df["format"] == 'Secabs') & (file_df["transfer"] == 'Email')]
        filenames=emails['secabs_filename'].drop_duplicates().to_list()
        total_files=move_or_copy_files(path_inv,filenames,'2',path_z)
        print('The total SECABS files moved has been: '+str(len(total_files)))
    
    

