import reports
import database as db
import os,shutil
import clients
import warnings

# Disable showing warnings
with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")


    function_dict = {'Selectronics': clients.selectronics, 'BENDTEL': clients.bendtel, 'MIEAC':clients.mieac, 'ONVOY':clients.onvoy
                    ,  'MTASW':clients.mta_sw , 'MTAFA':clients.mta_fa , 'MTARC':clients.mta_rc, "Neutral_TandemFA":clients.nt_fa
                    ,"GCISW":clients.gci_sw,'Neutral_TandemSW':clients.nt_sw , 'GCIFA':clients.gci_fa, 'American_Broadband356DMI':clients.amb_356DMI
                    ,'American_Broadband509BOH':clients.amb_509BOH,'American_Broadband590GIN':clients.amb_590GIN,'Peerless_Network':clients.peerless}


    #cleaning the database before start working. clearing the terminal
    db.clean_database()
    os.system('cls')

    # asking the user for the company which he is working on
    company=str(input("Please enter the company name (as it is shown in K drive): "))

    reports.pdf.company = company
    reports.temp_month = input('Please enter the month in 2 digits? for example (05) ')
    reports.temp_year = input('Please enter the year in 4 digits? for example (2021) ')
    month,year = reports.temp_month,reports.temp_year

    # If the company is one of the next ones, the bill type will have to be specified  
    if company == "MTA" or company == "Neutral_Tandem" or company == "American_Broadband" or company == "GCI":
        print('Please enter bill type (SW, FA or RC (RC is only for MTA), IN for AMBB Indiana, MI for AMBB Michigan, OH for AMBB Ohio)')
        bill_type=str(input()).upper()
    else: bill_type=''

    # Asking the user for the section he is working on 
    part = input("Please enter the group (PRE, POST1, POST2): ").upper()

    # Start to rename files and get thier locations
    db.renaming_reports(company, month, year, bill_type)
    reports.macro_path = db.get_all_files(company,month, year, bill_type)

    # getting the path of the MPS
    reports.xls.mps_path = db.retrieve_path('MPS') 

    # moving the Macro file in case of Peerless
    if company == 'Peerless_Network':
        shutil.copy(os.getcwd()+'\\Peerless_Network_Macros.xlsm',reports.macro_path)

    # Calling functions based on the client and section 
    function_dict[company+bill_type](part)

    # Removing macro files in case of Peerless
    if company == 'Peerless_Network':
        os.remove(billing_folder+'\\Peerless_Network_Macros.xlsm')    

    os.system("pause")