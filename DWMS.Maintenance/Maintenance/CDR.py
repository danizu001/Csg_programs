import os 
import fnmatch
import pandas as pd
import subprocess
import glob

def change_bat_file(edit,path_in,text):
    for i in range(len(edit)):
        new = open(path_in+'\\new.bat', "x")
        new.write(text[i])
        os.remove(path_in+'\\'+edit[i])
        new.close()
        os.rename(path_in+'\\new.bat', path_in+'\\'+edit[i])   

def run_bat_file(path_in,run):
    for i in run:
        subprocess.run([path_in+'\\'+i])

def check_double_ban(ban,CDR):
    if ban.endswith('E'):
        ban_change=ban[:-1]+'T'
    elif ban.endswith('T'):
        ban_change=ban[:-1]+'E'
    if CDR['CDR'].str.contains(ban_change).any():
        ban_flag=True
    else:
        ban_flag=False
    return ban_flag,ban_change

def combine(path_in,i,CDR):
    text2=''
    with open(path_in+'\\'+i, "rt") as bat_file:
        text = bat_file.readlines()
    ban=text[0][38:48]
    text=text[0].split('\\')
    text[4]=text[4].replace('*.dsp Q:',ban+'\\*.dsp Q:')
    text.insert(8,ban)
    text='\\'.join(text)
    ban_flag,ban_change=check_double_ban(ban,CDR)
    if ban_flag==True:
        with open(path_in+'\\'+i, "rt") as bat_file:
            text2=bat_file.readlines()
        text2=text2[0].split('\\')
        text2[4]=text2[4].replace('*.dsp Q:',ban_change+'\\*.dsp Q:')
        text2.insert(8,ban_change)
        text2[9]=text2[9].replace(ban,ban_change)
        text2='\\'.join(text2)
    return text+'\n' +text2
def compress(path_in,i,CDR):
    text2=''
    with open(path_in+'\\'+i, "rt") as bat_file:
        text = bat_file.readlines()
    ban=text[0][87:97]
    text=text[0].split('\\')
    text.insert(10,ban)
    text='\\'.join(text)
    ban_flag,ban_change=check_double_ban(ban,CDR)
    if ban_flag==True:
        with open(path_in+'\\'+i, "rt") as bat_file:
            text2=bat_file.readlines()
        text2=text2[0].split('\\')
        text2[6]=text2[6].replace(ban,ban_change)
        text2.insert(10,ban_change)
        text2[11]=text2[11].replace(ban,ban_change)
        text2='\\'.join(text2)
    return text+'\n' +text2

text_combine=[]
text_compress=[]
path_in=input('Enter the dispute path for example: Q:\SB_Utilities_Logs\gondan03-pa\OMAPRDICXCI05\disputes \n')
files = glob.glob(path_in+'\\*')
run_combine=[f for f in os.listdir(path_in) if fnmatch.fnmatch(f, '*_combine_a.bat')]
run_compress=[f for f in os.listdir(path_in) if fnmatch.fnmatch(f, '*_compress_a.bat')]
edit_combine=[f for f in os.listdir(path_in) if fnmatch.fnmatch(f, '*_combine_b.bat')]
edit_compress=[f for f in os.listdir(path_in) if fnmatch.fnmatch(f, '*_compress_b.bat')]
CDR=pd.read_excel(r'K:\Client\Peerless_Network\Disputes\CDR_Requests\2022\11_2022\CDRs.xlsx')
for i in edit_combine:
    text_combine.append(combine(path_in,i,CDR))
for i in edit_compress:
    text_compress.append(compress(path_in,i,CDR))
change_bat_file(edit_combine,path_in,text_combine)
change_bat_file(edit_compress,path_in,text_compress)
run_bat_file(path_in,run_combine)
run_bat_file(path_in,run_compress)

for f in files:
    os.remove(f)
