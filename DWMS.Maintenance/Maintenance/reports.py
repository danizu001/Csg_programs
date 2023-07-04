import Maintenance
import os
import shutil
import verify_zip
import Copy_Invoices_FF
import Frontend
import ClliMapping
import Trending
import CleanUp
import Error_files
import dashboard
import download
import edit_excel
import mps_audit
import occs
import thresholding
import audit_dashboard

dates={'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun'
    ,'07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
dict_verifyZIP={'Peerless_Network':'PeerlessNetwork','Neutral_Tandem':'NTandem','American_Broadband':'AmericanBroadband','ONVOY':'Onvoy'}
def run_reports(company,bill_type,month,year):
    path_all='K:\\Client\\'+company+'\\Reports\\'+year+'\\'+month+'_'+year+'\\'
    #dict=[path reports,mps_name,Send to client folder,standard file name,trending_number_SW]
    dict_paths={'Peerless_Network':[path_all,'MPS_All_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','Peerless_'+month+'05'+year[2:],''],
                'GCI':[path_all+'Switched Billing\\','MPS_GCI_SW_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','GCI_SW_'+month+'21'+year[2:],'3'],
                'MTA':[path_all+'Switched Billing\\','MPS_SW_3015_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','MTA_SW_'+month+'10'+year[2:],'4'],
                'BendTel':[path_all,'MPS_OR_9627_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','9627_OR_'+month+'22'+year[2:],'2'],
                'American_BroadbandOH':[path_all+'509B_OH\\','MPS_OH_509B_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','509B_OH_'+month+'20'+year[2:],'2'],
                'American_BroadbandMI':[path_all+'356D_MI\\','MPS_MI_356D_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','356D_MI_'+month+'25'+year[2:],'1'],
                'American_BroadbandIN':[path_all+'590G_IN\\','MPS_IN_590G_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','590G_IN_'+month+'25'+year[2:],'3'],
                'Neutral_Tandem':[path_all+'Switched Billing\\','MPS_SW_Multi_Combined_'+month+'_'+year[2:]+'.xlsx','Send to Client\\',['Combined_OSA_'+month+'05'+year[2:],'Combined_TTS_'+month+'01'+year[2:]],''],
                'MIEAC':[path_all,'MPS_MN_8811_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','MIEAC_'+month+'01'+year[2:],''],
                'ONVOY':[path_all,'MPS_SW_Multi_Combined_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','ONVOY_'+month+'01'+year[2:],''],
                'Neutral_TandemFA':[path_all+'Facility Billing\\','MPS_FAC_Multi_Combined_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','Combined_IQNT_'+month+'01'+year[2:],''],
                'GCIFA':[path_all+'Facility Billing\\','MPS_GCI_FA_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','GCI_FA_'+month+'01'+year[2:],''],
                'MTAFA':[path_all+'Facility Billing\\','MPS_FAC_3015_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','MTA_FA_'+month+'10'+year[2:],''],
                'MTARC':[path_all+'Recip_Comp\\','MPS_RC_MULTI_'+month+'_'+year[2:]+'.xlsx','Send to Client\\','MTA_RC_'+month+'05'+year[2:],'']}
    if bill_type=='Payments':
        path='K:\\Client\\'+company+'\\Payments\\'+year+'\\'+month+'_'+year
        files=os.listdir(path)
        try:
            os.makedirs(path+'\\Done')    
            print("Directory Created ")
        except FileExistsError:
            print("Directory already exists")
        if company!='Neutral_Tandem':
            for i in files:
                if i!='Done':
                    Maintenance.call(company,path+'\\'+i)
                    shutil.move(path+'\\'+i,path+'\\Done')
                    if(i.endswith('.xls')):
                        file=i.replace('.xls','_import.xls')
                        shutil.move(path+'\\'+file,path+'\\Done')
                        file_txt=file.replace('.xls','.txt')
                        shutil.move(path+'\\'+file_txt,path+'\\Done')
                    else:
                        file=i.replace('.xlsx','_import.xlsx')
                        shutil.move(path+'\\'+file,path+'\\Done')
                        file_txt=file.replace('.xlsx','.txt')
                        shutil.move(path+'\\'+file_txt,path+'\\Done')
        else:
            NT_files=os.listdir(path)
            for i in NT_files:
                if i!='Done':
                   comment=Maintenance.Payment_NT(path+'\\'+i)
                   print(comment) 
            NT_files=os.listdir(path)
            for i in NT_files:
                if i!='Done':
                    shutil.move(path+'\\'+i,path+'\\Done')
        print('Payments are done')
    elif bill_type=='Verify zip':
        company=dict_verifyZIP[company]
        path = 'Q:\\'+company+'\\'+year+'_'+month
        verify_zip.call(path)
        print('Verify Zip is already done')
    elif bill_type=='Volume trending':
        FILE_NAME='Customer Volumes Trending Study_FY'+str(year[2:])+'.xlsx'
        # FILE_NAME='Customer Volumes Trending Study_FY23.xlsx'
        if company=='American_Broadband':
            mps_list=[dict_paths[company+'OH'][:2],dict_paths[company+'MI'][:2],dict_paths[company+'IN'][:2]]
            revenue=[dict_paths[company+'OH'][2:4],dict_paths[company+'MI'][2:4],dict_paths[company+'IN'][2:4]]
        elif company=='GCI' or company=='Neutral_Tandem':
            mps_list=[dict_paths[company][:2],dict_paths[company+'FA'][:2]]
        elif company=='MTA':
            mps_list=[dict_paths[company][:2],dict_paths[company+'FA'][:2],dict_paths[company+'RC'][:2]]
            revenue=[dict_paths[company][2:4],dict_paths[company+'FA'][2:4],dict_paths[company+'RC'][2:4]]
        else: 
            mps_list=[dict_paths[company][:2]]
            revenue=[dict_paths[company][2:4]]
        FOLDER_DEST='K:\\Client\\Service_bureau\\Audit\\Customer Volumes Trending\\Script_connection\\'
        download.call('CABS/Service Bureau - Audit',FILE_NAME,FOLDER_DEST,0)
        edit_excel.call(FOLDER_DEST,mps_list, company, month,FILE_NAME,revenue,'_Revenue Analysis.xlsx')
        download.call('CABS/Service Bureau - Audit',FILE_NAME,FOLDER_DEST,1)
        print('Volume trending is already done')
    elif bill_type=='Move Secabs':
        path="Q:\\SECABS_PA_FF\\Z_VERIFY_INPUT\\To Move"
        files=os.listdir(path)
        path_ff=path+'\\'+files[0]
        Copy_Invoices_FF.call(path_ff,'2')
        dest="Q:\\SECABS_PA_FF\\Z_VERIFY_INPUT\\"+files[0]
        shutil.move(path_ff, dest)
    elif bill_type=='Clli mapping':
        path=dict_paths['Peerless_Network'][0]+'\\Other Reports\\CLLI Mapping'
        command=Frontend.Switch_Facility('Part 1','Part 2')
        if command=='Part 1':
            ClliMapping.call(path+'\\wire_center.txt',path+'\\Clli_map.xlsx','0',path)
            print('Cabs ONE is done')
        if command=='Part 2':
            ClliMapping.call(path+'\\wire_center_b.txt',path+'\\Clli_map.xlsx','1',"")
            print('Cabs ONE is done')
    elif bill_type=='Bill Count Trending':
        if company=='American_Broadband':
            command=Frontend.OH_MI_IN()
            path = dict_paths[company+command][0]+dict_paths[company+command][1]
            Trending.call('1',path,dict_paths[company+command][4])
            print('Cabs ONE is done')
        else:
            path = dict_paths[company][0]+dict_paths[company][1]
            Trending.call(dict_paths[company][4],path)
            print('Cabs ONE is done')
    elif bill_type=='Clean Up':
        company=dict_verifyZIP[company]  
        path = 'Q:\\'+company+'\\'+year+'_'+month
        CleanUp.call(path)
        print('Verify Zip is already done')
    elif bill_type=='Error files':
        ban_error_file=dict_paths[company][0]+'Other Reports\\BanError.txt'
        ban_template=dict_paths[company][0]+'Other Reports\\New BANs_template.xlsx'
        target='Q:\\PeerlessNetwork\\'+year+'_'+month
        Error_files.call(ban_error_file,ban_template,target)
    elif bill_type=='Dashboard':
        dashboard_path="K:\\Client\\Service_bureau\\Audit\\Production Dashboard\\Production_Dashboard_Data_"+str(year)+".xlsx"
        month_dash=dates[month]
        year_dash=year[2:]
        if company=='Peerless_Network':
            mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\MPS_All_'+month+'_'+year[2:]+'.xlsx'
            client='PEERLESS'
            bill_type='SW'
            dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        if company=='GCI':
            command=Frontend.Switch_Facility('Switch','Facility')
            if command=='Switch':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Switched Billing\\MPS_GCI_SW_'+month+'_'+year[2:]+'.xlsx'
                client='GCI'
                bill_type='SW'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
            if command=='Facility':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Facility Billing\\MPS_GCI_FA_'+month+'_'+year[2:]+'.xlsx'
                client='GCI'
                bill_type='FA'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)   
        if company=='MTA':
            command=Frontend.Switch_Facility('Switch','Recip Comp')
            if command=='Switch':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Switched Billing\\MPS_SW_3015_'+month+'_'+year[2:]+'.xlsx'
                mps2='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Facility Billing\\MPS_FAC_3015_'+month+'_'+year[2:]+'.xlsx'

                client='MTA'
                bill_type='SW'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash,mps2)
            if command=='Recip Comp': #change to RC
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Recip_Comp\\MPS_RC_MULTI_'+month+'_'+year[2:]+'.xlsx'
                client='MTA'
                bill_type='RC'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        if company=='Neutral_Tandem':
            command=Frontend.Switch_Facility('Switch','Facility')
            if command=='Switch':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Switched Billing\\MPS_SW_Multi_Combined_'+month+'_'+year[2:]+'.xlsx'
                client='NEUTRAL TANDEM'
                bill_type='SW'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
            if command=='Facility':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\Facility Billing\\MPS_FAC_Multi_Combined_'+month+'_'+year[2:]+'.xlsx'
                client='NEUTRAL TANDEM'
                bill_type='FA'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        if company=='BendTel':
                mps='K:\\Client\\'+company+'\\reports\\'+year+'\\'+month+'_'+year+'\\MPS_OR_9627_'+month+'_'+year[2:]+'.xlsx'
                client='BENDTEL'
                bill_type='SW'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        if company=='American_Broadband':
            command=Frontend.OH_MI_IN()
            if command=='OH':
                mps='K:\\Client\\American_Broadband\\Reports\\'+year+'\\'+month+'_'+year+'\\509B_OH\\MPS_OH_509B_'+month+'_'+year[2:]+'.xlsx'
                client='AMBB'
                bill_type='OH'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
            if command=='MI':
                mps='K:\\Client\\American_Broadband\\Reports\\'+year+'\\'+month+'_'+year+'\\356D_MI\\MPS_MI_356D_'+month+'_'+year[2:]+'.xlsx'
                client='AMBB'
                bill_type='MI'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
            if command=='IN':
                mps='K:\\Client\\American_Broadband\\Reports\\'+year+'\\'+month+'_'+year+'\\590G_IN\\MPS_IN_590G_'+month+'_'+year[2:]+'.xlsx'
                client='AMBB'
                bill_type='IN'
                dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)  
        if company=='ONVOY':
            mps='K:\\Client\\ONVOY\\Reports\\'+year+'\\'+month+'_'+year+'\\MPS_SW_Multi_Combined_'+month+'_'+year[2:]+'.xlsx'
            client='ONVOY'
            bill_type='SW'
            dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        if company=='MIEAC':
            mps='K:\\Client\\MIEAC\\Reports\\'+year+'\\'+month+'_'+year+'\\MPS_MN_8811_'+month+'_'+year[2:]+'.xlsx'
            client='MIEAC'
            bill_type='SW'
            dashboard.call(dashboard_path, mps, client, bill_type, month_dash, year_dash)
        FILE_NAME='Production_Dashboard_Data_'+str(year)+' - Update.xlsx'
        FOLDER_DEST='K:\\Client\\Service_bureau\\Audit\\Production Dashboard\\Update teams\\'
        download.call('CABS/Service Bureau - Audit/Dashboard Script',FILE_NAME,FOLDER_DEST,1)
        os.remove("K:\\Client\\Service_bureau\\Audit\\Production Dashboard\\Update teams\\Production_Dashboard_Data_"+str(year)+" - Update.xlsx")
        print('Cabs One is done')
    elif bill_type== "Audit MPS":
        if company=='Neutral_Tandem' or company=='GCI' :
            command=Frontend.Switch_Facility("Switch","FA")
            if command=='Switch':
                btype='SW'
            else:
                btype='FA'
        elif company=='MTA':
            command=Frontend.Switch_Facility_Recip()#RECIP COMP
            if command=='SW':
                btype='SW'
            elif command=='FA':
                btype='FA'
            else:
                btype="RC"
        elif company=='American_Broadband':
            command=Frontend.OH_MI_IN()
            if command=='OH':
                btype='OH'
            elif command=='IN':
                btype='IN'
            else:
                btype='MI'
        else:
            btype=''
        mps_audit.call(company,btype,year)
        print("CABS ONE is done")
    elif bill_type=='OCC':
        path='K:\\Client\\'+company+'\\OCCS\\'+year+'\\'+month+'_'+year
        occs.occ(path)
        print("CABS ONE is done")
    elif bill_type=='Thresholding':
            command=Frontend.Switch_Facility("Prethresh","PosThresh")
            path='K:\\Client\\'+company+'\\Reports\\'+year+'\\'+month+'_'+year+'\\'
            if command=="Prethresh":
                thresholding.pre_threshold(path)
            else:
                thresholding.pos_threshold(path)
            print("CABS ONE is done")
    elif bill_type=='Audit Dashboard':
        audit_dashboard.call(str(year))
        print("CABS ONE is done")



