from genericpath import exists
import os
import shutil
import verify_zip

def clean_bad_processed_files(states):
    already_done=[]
    for state_path in states:
        files_state=[]
        clean=state_path+"\\Clean Up\\Clean_Scrpt\\"
        isExist = os.path.exists(clean)
        if not isExist:
            # Create a new directory because it does not exist 
            os.makedirs(clean)
            print("Clean Zip has been created")
        files=os.listdir(state_path)
        clear_files=[file for file in files if os.path.isfile(state_path+file) and '.zip' not in file.lower() and '.err.txt' not in file.lower() and '.drop.txt' not in file.lower() and '.java.log' not in file.lower() and '.out' not in file.lower() and '_sort.txt' not in file.lower() ]
        for total_file in clear_files :
            if exists(state_path+total_file+'.err.txt') and exists(state_path+total_file+'.drop.txt') and exists(state_path+total_file+'.java.log') and exists(state_path+total_file+'.out') and not exists(state_path+total_file+'_sort.txt'):
                original_file=total_file
                if exists(state_path+original_file):
                    if state_path+original_file not in already_done:
                        print("Original File found: "+state_path+original_file)
                        files_state.append(state_path+original_file)
                        already_done.append(state_path+original_file)
                        error_file=os.path.join(state_path, total_file+'.err.txt')
                        shutil.move(error_file, clean)
                        print("File moved: "+error_file)
                        drop_file=os.path.join(state_path, total_file+'.drop.txt')
                        shutil.move(drop_file, clean)
                        print("File moved: "+drop_file)
                        java_file=os.path.join(state_path, total_file+'.java.log')
                        shutil.move(java_file, clean)
                        print("File moved: "+java_file)
                        out_file=os.path.join(state_path, total_file+'.out')
                        shutil.move(out_file, clean)
                        print("File moved: "+out_file)
                        print("\n")
            if not exists(state_path+total_file+'.err.txt') and not exists(state_path+total_file+'.drop.txt') and exists(state_path+total_file+'.java.log') and not exists(state_path+total_file+'.out') and not exists(state_path+total_file+'_sort.txt'):
                original_file=total_file
                if exists(state_path+original_file):
                    if state_path+original_file not in already_done:
                        print("Original File found: "+state_path+original_file)
                        java_file=os.path.join(state_path, total_file+'.java.log')
                        shutil.move(java_file, clean)
                        print("File moved: "+java_file)
                        print('\n')
    for file in files_state :
        if not exists(file):
            print("File doesnt exist: "+file)

def call(path):
    states=[]
    content = next(os.walk(path))[1]
    try:
        content.remove('badrecords')
        content.remove('CABSOUT')
        bad=verify_zip.unzipped_files(path, content)
        wrong=verify_zip.wrong_zipped(path,bad)
    except:
        bad=verify_zip.unzipped_files(path, content)
        wrong=verify_zip.wrong_zipped(path,bad)
    for i in wrong:
        states.append(path+'\\'+i+'\\')
    clean_bad_processed_files(states)