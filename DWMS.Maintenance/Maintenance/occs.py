import pandas as pd
import os
import shutil

def occ (path):
    foundfiles=0
    for file in os.listdir(path):
        if file.startswith("DTP"):
            fpath=os.path.join(path, file)
            donepath=path+'\\Done'
            isExist = os.path.exists(donepath)
            if not isExist:
                os.makedirs(donepath)
            path=fpath
            df= pd.read_excel(path)
            #df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            # df.drop('Unnamed', inplace=True, axis=1)
            if df.columns.str.match("Unnamed").any():
                df= df.loc[:,~df.columns.str.match("Unnamed")]
            txt=path.replace(".xlsx","_import.txt")
            df.to_csv(txt, index=False, sep='\t')
            shutil.move(txt, donepath)
            shutil.move(path, donepath)
            print(df)
            foundfiles=1
    if foundfiles==0:
        print("There are no new OCC files")

