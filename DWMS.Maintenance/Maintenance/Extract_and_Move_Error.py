#Author : Rafek Gorgay

import os
from os import listdir
from os.path import isfile, join, isdir
import shutil, zipfile
import pandas as pd

def extract_and_move(target):
    # clearing the terminal to have the colors working as it is a known bug for windows
    os.system('cls')
    
    not_moved_error_files = [] # this list will contain ZIPPED files that did not have .err.txt
    moved_error_files = 0
    
    # setting the destination location from the user
    
    # saving the files from the CSV into a list and remove the space at the end
    files = pd.read_csv(os.getcwd() + '\\Files_list.csv')
    files = list(files['Files'].replace({"^\s*|\s*$":""}, regex=True))
    
    
    for file in files:
    
        # setting the target_folder to be the same state of the zipped file
        target_folder = target + '\\' + file.split('\\')[3]
        index = 10
        try:
            # getting the index of the err.txt file (always skip the original file)
            files_in_zipped = zipfile.ZipFile(file+'.zip').namelist()
            for item in files_in_zipped:
                if item.endswith('.err.txt') and files_in_zipped.index(item) != 0: 
                    index = files_in_zipped.index(item)
    
            # extract the .err.txt file to the path of target folder
            zipfile.ZipFile(file+'.zip').extract(files_in_zipped[index], path = target_folder)
            moved_error_files += 1
        except:
            not_moved_error_files.append(file)
    
               
    print("\033[0;36;40m" + str(moved_error_files)+  " are extracted and moved from " + str(len(files)) + "\033[1;36;40m")
    
    if len(not_moved_error_files) > 0:
        print("\033[1;31;40m" + "These files need to be fixed" + "\033[0;37;40m")
        for i in not_moved_error_files:
            print(str(not_moved_error_files.index(i)+1) + '- ' + i)
    elif len(files) == moved_error_files:
        print("\033[1;32;40m" + "All Error files were moved sucessfully" + "\033[0;37;40m")
    
    os.system("pause")

def __init__(self):
    target = input("Please enter the location of current billing month (where to move the error files to?): ")
    extract_and_move(target)  
