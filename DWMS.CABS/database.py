import os,datetime,pyodbc
from os import listdir
from os.path import isfile, join, isdir
import excel as xls
from datetime import datetime

# creating log file in case that it was not exist
if not (os.path.isfile('log.txt')):
    log_file = open('log.txt', 'w')
else:
    log_file = open('log.txt', 'a')

log_file.write('((((((((((((((((((((((((((((((((((((((((((((-----------------------&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&-----------------------))))))))))))))))))))))))))))))))))))))))))))\n\n')
log_file.write('Today\'s date : ' + datetime.now().strftime("%d/%m/%Y %H:%M:%S")+'\n')


# establishing the database connection 
db_location = os.getcwd()+"\\Reports Database.mdb"
try:
    conn = pyodbc.connect(Driver="{Microsoft Access Driver (*.mdb)}" , DBQ=db_location)
except:
    conn = pyodbc.connect(Driver="{Microsoft Access Driver (*.mdb, *.accdb)}" , DBQ=db_location)

cursor = conn.cursor()

bill_type_dic = {'SW':'\\Switched Billing','FA':'\\Facility Billing', 'RC':'\\Recip_Comp'
                    ,'356DMI':'\\356D_MI','509BOH':'\\509B_OH' ,'590GIN':'\\590G_IN','':''}


# Renames the reports to follow each client standards 
def renaming_reports (company,month,year,bill_type):
    
    log_file.write('========================================================================= Renaming Files =========================================================================\n')
    print('Going to rename files')
    number_of_files=0

    before_renaming = []
    after_renaming = []
    not_OSA_TTS_files=[]
    prefix_dict = {}
    billdate_dict = {}
    code1=[]
    code2=[]
    code3=[]

    # Get the prefix (client name + bill date), report name and report codes from the database.
    cursor.execute('select * from Prefix')
    for row in cursor.fetchall():
        prefix_dict[row.Client]=row.Prefix
        billdate_dict[row.Client]=row.BillDate
    cursor.execute('select * from '+company)
    for row in cursor.fetchall():
        before_renaming.append(row.BeforeRenaming)
        after_renaming.append(row.AfterRenaming)
        code1.append('' if row.Code1 is None else row.Code1)
        code2.append('' if row.Code2 is None else row.Code2)
        code3.append('' if row.Code3 is None else row.Code3)

    # checking if the billing is NT SW to exclude OSA and TTS from the naming of the reports that are located under not_OSA_TTS_files from the database.
    if company == 'Neutral_Tandem' and bill_type == 'SW':
        cursor.execute('select not_OSA_TTS_files from Neutral_Tandem')
        for row in cursor.fetchall():
            not_OSA_TTS_files.append(row.not_OSA_TTS_files)

    # setting the path of the billing
    path = "K:\\Client\\" + company + "\\Reports\\" + year + "\\" + month + "_" + year + bill_type_dic[bill_type]
    # gets all files from the billing month
    files_list = [ f for f in listdir(path) if isfile(join(path, f))]

    # naming all reports
    # the loop below takes code1 and add it to the name of the file and saves, in case that the file exists with the same name it takes code2 and saves and so on till code3
    for file in files_list:
        for report in before_renaming:
            index = before_renaming.index(report)
            if file.upper().startswith(str(report).upper()):
                extension = file.split('.')
                if company == 'Neutral_Tandem'and bill_type == 'SW':
                    if report in not_OSA_TTS_files:
                        file_name = prefix_dict[company+bill_type] + '_' + month + year[2:] +'_'
                    elif extension[0].endswith('_02') or extension[0].endswith('_04'):
                        file_name = prefix_dict[company+bill_type+'T']+ '_' + month +billdate_dict[company+bill_type+'T'] + year[2:] +'_'
                    else:
                        file_name = prefix_dict[company+bill_type+'O']+ '_' + month +billdate_dict[company+bill_type+'O'] + year[2:] +'_'
                else:
                    file_name = prefix_dict[company+bill_type] + '_' + month +billdate_dict[company+bill_type]+ year[2:] +'_'

                try:
                    new_name = file_name + code1[index].upper() +after_renaming[index]+'.'+extension[1]
                    os.rename(path+ "\\" +file, path+ "\\"+new_name)
                    number_of_files+=1
                except:
                    try:
                        new_name = file_name + code2[index].upper() +after_renaming[index]+'.'+extension[1]
                        os.rename(path+ "\\" +file, path+ "\\"+new_name)
                        number_of_files+=1
                    except:
                        new_name = file_name + code3[index].upper() +after_renaming[index]+'.'+extension[1]
                        os.rename(path+ "\\" +file, path+ "\\"+new_name)
                        number_of_files+=1


                print(file + ' Renamed to -> ' + new_name)
                log_file.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + ' | '+ file + ' -> ' + new_name + '\n')
                break   
    print(number_of_files,' files were renamed')
    log_file.write('\t\t\t\t\t\t\t\t\t- '+str(number_of_files)+' were reneamed - \n')
    log_file.flush()
    os.fsync(log_file.fileno())

# this function converts XLS files to XLSX, because we can not work with Old version of Excel
# No more need for this function, because auto reports releases the latest version of Excel
def convert_files(path):
    log_file.write('========================================================================= Converting Files =========================================================================\n')
    
    number_of_files=0
    not_needed_files_ends = ['AUR_History_Stats.xls','AUR History Stats Report.xls','AUR_History_Report.xls','Billing Notes.xls', 'Bill Count.xls','FGA_Usage_Balance_Details.xls','Usage by Day Report.xls','SW BANS Trending Report.xls']
    not_needed_files_starts = ['Fairpoint_moucomp','fusctax','Intrastate Terminating','Revenue_NECA','FF']
    filesList = [f for f in listdir(path) if isfile(join(path, f))]
    for file in filesList:
        if ((file.endswith(".xls")) and (file.endswith(tuple(not_needed_files_ends)) is False) and (file.startswith(tuple(not_needed_files_starts)) is False)):
            print('Convert to xlsx | ' + file)
            xls.convert_xls_to_xlsx(path + "\\" + file)
            log_file.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") +' | ' + file+ "was converted to XLSX version\n")
            number_of_files+=1
    print(number_of_files,' files were converted')
    log_file.write('\t\t\t\t\t\t\t\t\t- '+str(number_of_files)+' files were converted -\n')
    log_file.flush()
    os.fsync(log_file.fileno())

# get the location of reports and save it into the database under "mps_database"
def get_all_files(company,month,year,bill_type):
    
    path = 'K:\\Client\\' + company + '\\Reports\\' + year + "\\" + month + "_" + year + bill_type_dic[bill_type]
    if company == 'MTA' or company == 'SELECTRONICS':
        convert_files(path)
    mps_code_list =[]
    report_name_list=[]

    # save the reports from the MPS to list to check later if they exist in the billing folder. If yes we get the location of the report.
    cursor.execute('select * from mps_database')
    for row in cursor.fetchall():
        mps_code_list.append(row.MPSLocation)
        report_name_list.append(row.ReportName)

    
    print("\033[1;36;40m" +'Importing files location to database' + "\033[0;37;40m")

    # getting all files that are located in the billing folder
    files_list = [file for file in listdir(path) if isfile(join(path, file))]

    # saving the report location to the table "mps_database"
    for file in files_list:  
        name=file.split("_")
        extension = name[-1].split('.')
        try:
            if 'OSA' in name or 'TTS' in name:
                if name[-2]+name[1] in mps_code_list :
                    index = mps_code_list.index(name[-2]+name[1])
                    cursor.execute('update mps_database set FileLocation= ? where ID= ?' , (path+'\\'+file,index+1))  

                elif extension[0]+name[1] in report_name_list:
                    index = report_name_list.index(extension[0]+name[1])
                    cursor.execute('update mps_database set FileLocation= ? where ID= ?' , (path+'\\'+file,index+1)) 
            else:

                if name[-2] in mps_code_list and file.startswith('MPS') != True:
                    index = mps_code_list.index(name[-2])
                    cursor.execute('update mps_database set FileLocation= ? where ID= ?' , (path+'\\'+file,index+1)) 
                elif file.startswith('MPS'):
                    index = mps_code_list.index('MPS')
                    cursor.execute('update mps_database set FileLocation= ? where ID= ?' , (path+'\\'+file,index+1)) 
                elif extension[0] in report_name_list:
                    index = report_name_list.index(extension[0])
                    cursor.execute('update mps_database set FileLocation= ? where ID= ?' , (path+'\\'+file,index+1)) 
        except:
            print(file,' is not named properly.',name)
    conn.commit()
    print("\033[1;36;40m" +"Importing to Database is done" + "\033[0;37;40m")
    return path

# get cell location for the report to write its value in the MPS
def retrieve_coordinates(code, option = ''):
    mps_code_list =[]
    cursor.execute('select * from mps_database')
    for row in cursor.fetchall():
        mps_code_list.append(row.MPSLocation)

    if code+option in mps_code_list:
        index = mps_code_list.index(code+option)
        cursor.execute('select * from mps_database where ID= ?', (index+1))
        return cursor.fetchall()[0][4]

# get the report location
def retrieve_path(code, option = ''):
    mps_code_list =[]
    cursor.execute('select * from mps_database')
    for row in cursor.fetchall():
        mps_code_list.append(row.MPSLocation)

    if code+option in mps_code_list:
        index = mps_code_list.index(code+option)
        cursor.execute('select * from mps_database where ID= ?', (index+1))
        return cursor.fetchall()[0][3]

# clear all values from table "mps_database" column "FileLocation"
def clean_database():
    cursor.execute('UPDATE mps_database SET FileLocation = null')
    conn.commit()