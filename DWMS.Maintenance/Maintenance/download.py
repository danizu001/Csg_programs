from office365_api import SharePoint
import re
import sys, os
from pathlib import PurePath


def save_file(file_n, file_obj, FOLDER_DEST):
    file_dir_path = PurePath(FOLDER_DEST, file_n)
    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)

def get_file(file_n, folder, FOLDER_DEST):
    file_obj = SharePoint().download_file(file_n, folder)
    save_file(file_n, file_obj, FOLDER_DEST)


def upload_files(folder, FOLDER_NAME, keyword=None,):
    file_list = get_list_of_files(folder)
    for file in file_list:
        if keyword is None or keyword == 'None' or re.search(keyword, file[0]):
            file_content = get_file_content(file[1])
            SharePoint().upload_file(file[0], FOLDER_NAME, file_content)
            
def get_list_of_files(folder):
    file_list = []
    folder_item_list = os.listdir(folder)
    for item in folder_item_list:
        item_full_path = PurePath(folder, item)
        if os.path.isfile(item_full_path):
            file_list.append([item, item_full_path])
    return file_list

def get_file_content(file_path):
    with open(file_path, 'rb') as f:
        return f.read()

def call(FOLDER_NAME,FILE_NAME,FOLDER_DEST,code,NAME_PATTERN=None):
    if FILE_NAME != 'None' and code==0:
        get_file(FILE_NAME, FOLDER_NAME, FOLDER_DEST)
    if FILE_NAME != 'None' and code==1:    
        upload_files(FOLDER_DEST, FOLDER_NAME, NAME_PATTERN )