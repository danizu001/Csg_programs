import pandas as pd
import numpy as np
import os
from datetime import datetime
from openpyxl import load_workbook
import openpyxl.utils.cell

def Verify_Bans_Invoices(column_bans, column_invoices, file):
    file=file.split("\\")
    file[-1]="inv_bal.xls"
    file="\\".join(file)
    report_df = pd.read_excel(file, converters={'user_invoice_num':str})
    bans_list = list(set(column_bans) - set(report_df['ban']))
    invoice_list = list(set(column_invoices) - set(report_df['user_invoice_num']))
    return bans_list, invoice_list

def Payment_ABB(file):
    value=[]
    value2=0
    report_df = pd.read_excel(file,sheet_name='Payment Template',converters={'user_invoice_num':str})
    len_data = len(report_df['Carrier Name'].dropna())
    report_df = report_df.drop(report_df[len_data:].index)
    report_df['Invoice #'] = [str(i).zfill(11) for i in report_df['Invoice #']]
    if(round(report_df['Total Check Amount'].sum(),2)==round(report_df['Detail Amount (by invoice)'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Total Check Amount'].sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    report_df['Total Check Amount'] = report_df['Total Check Amount'].fillna(0)
    try:
        report_df['Check #'] = pd.to_numeric(report_df['Check #'], downcast='integer')
    except:
        pass
    report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    report_df['Billing Account'] = [i.replace(' ','') for i in report_df['Billing Account']]
    for i in report_df['Total Check Amount']:
        if(i != 0):
            value.append(i)
            value2=i
        else:
            value.append(value2)
    report_df['Total Check Amount']=value
    try:
        bans_list, invoice_list = Verify_Bans_Invoices(report_df['Billing Account'], report_df['Invoice #'], file)
        bans_list=', '.join(bans_list)
        invoice_list=', '.join(invoice_list)
        print('Missing bans: '+bans_list+"\nMissing invoices: "+invoice_list)
        file=file.replace('.xlsx','_import.xlsx')
        file_txt=file.replace('.xlsx','.txt')
        report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
        report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    except:
        file=file.replace('.xlsx','_import.xlsx')
        file_txt=file.replace('.xlsx','.txt')
        report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
        report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    return comment

def Payment_Bendtel(file):
    value=[]
    value2=0
    report_df = pd.read_excel(file,sheet_name='Payment received',header=None)
    limit=list(report_df.iloc[2])
    limit=[i for i in range(limit.index('Detail Amount (by invoice)')+1)]
    report_df = report_df[limit]
    report_df.columns=report_df.iloc[2]
    report_df = report_df.drop(report_df[:4].index).reset_index(drop=True)
    report_df['Billing Account'] = report_df['Billing Account'].replace('Be sure the columns equal',float('NaN'))
    count=0
    for i in range(len(report_df['Billing Account'])):
        if type(report_df['Billing Account'][i]) is float:
            count+=1
        else:
            count=0
        if count==2:
            data_length=i-1
            break
    report_df=report_df.iloc[:data_length]
    report_df['Total Check Amount'] = report_df['Total Check Amount'].fillna(0)
    report_df['Check #'] = report_df['Check #'].fillna(0)
    report_df['Carrier Name'] = report_df['Carrier Name'].fillna(0)
    report_df['Payment Date'] = report_df['Payment Date'].fillna(0)
    report_df = report_df.dropna().reset_index()
    for i in range(len(report_df['Total Check Amount'])):
        if(report_df['Total Check Amount'][i] != 0):
            value.append([report_df['Total Check Amount'][i],report_df['Check #'][i], report_df['Carrier Name'][i], report_df['Payment Date'][i]])
            value2=[report_df['Total Check Amount'][i],report_df['Check #'][i], report_df['Carrier Name'][i], report_df['Payment Date'][i]]
        else:
            value.append(value2)
    if(round(report_df['Total Check Amount'].sum(),2)==round(report_df['Detail Amount (by invoice)'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Total Check Amount'].sum(),2))
    else:
        comment="The value amount doesn't match correctly"    
    value_list = np.array(value)
    transpose_value = value_list.T
    transpose_value = transpose_value.tolist()
    report_df['Total Check Amount']=transpose_value[0]
    report_df['Check #']=transpose_value[1]
    report_df['Carrier Name']=transpose_value[2]
    report_df['Payment Date']=transpose_value[3]
    report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    report_df.drop(columns=report_df.columns[0], axis=1, inplace=True)
    if(file.endswith('.xlsx')):
        file=file.replace('.xlsx','_import.xlsx')
        file_txt=file.replace('.xlsx','.txt')
        report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
        report_df.to_csv (file_txt, index = False, header=False,sep='\t')
    else:
        file=file.replace('.xls','_import.xls')
        file_txt=file.replace('.xls','.txt')
        report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
        report_df.to_csv (file_txt, index = False, header=False,sep='\t')
    return comment

def Payment_GCI(file):
    report_df = pd.read_excel(file,sheet_name='Payment Template',header=None)
    limit=list(report_df.iloc[0])
    limit=[i for i in range(limit.index('Detail Amount (by invoice)')+1)]
    report_df = report_df[limit]
    report_df.columns = report_df.iloc[0]
    report_df = report_df.drop(report_df.index[0])
    report_df = report_df.dropna()
    if report_df.columns[0] == 'Total Wire Amount':
        report_df.rename(columns = {'Total Wire Amount':'Total Check Amount'}, inplace = True)
    report_df['Total Check Amount']=round(report_df['Total Check Amount'].astype(float),2)
    if(round(report_df['Total Check Amount'].unique().sum(),2)==round(report_df['Detail Amount (by invoice)'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Total Check Amount'].unique().sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    try:
        report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    except:
        report_df['Payment Date'] = [datetime.strptime(i,'%m.%d.%y') for i in report_df['Payment Date']]
        report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    report_df['Detail Amount (by invoice)'] = [round(i,2) for i in report_df['Detail Amount (by invoice)']]
    if file.endswith('.xls'): 
        file=file.replace('.xls','_import.xls')
        file_txt=file.replace('.xls','.txt')
    elif file.endswith('.xlsx'):
        file=file.replace('.xlsx','_import.xlsx')
        file_txt=file.replace('.xlsx','.txt')
    report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    return comment

def Payment_Mieac(file,sheet):
    value=[]
    value2=0
    report_df = pd.read_excel(file,sheet_name=sheet, header=None)
    column=list(report_df.iloc[0])
    report_df.columns=column
    report_df = report_df.drop(report_df.index[0])
    limit=[i for i in range(len(report_df['Pay Type'].dropna()))]
    report_df = report_df.iloc[limit]
    limit = column.index('Gross Amount')
    report_df = report_df[column[:limit+1]]
    if(round(report_df['Total Check Amount'].sum(),2)==round(report_df['Gross Amount'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Total Check Amount'].sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    report_df['Total Check Amount'] = report_df['Total Check Amount'].fillna(0)
    for i in report_df['Total Check Amount']:
        if(i != 0):
            value.append(round(i,2))
            value2=round(i,2)
        else:
            value.append(value2)
    report_df['Total Check Amount']=value
    report_df['Transaction Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Transaction Date']]
    file=file.replace('.xlsx','_import.xlsx')
    file_txt=file.replace('.xlsx','.txt')
    report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    return comment

def Payment_Mta(file):
    report_df = pd.read_excel(file, header=None)
    column=list(report_df.iloc[0])
    report_df.columns=column
    report_df = report_df.drop(report_df.index[0])
    limit=[i for i in range(len(report_df['Check #'].dropna()))]
    report_df = report_df.iloc[limit]
    if(round(report_df['Check Amount'].unique().sum(),2)==round(report_df['Detail Amount (by invoice)'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Check Amount'].unique().sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    try:
        report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    except:
        report_df['Payment Date'] = [datetime.strptime(i,'%m/%d/%Y').date() for i in report_df['Payment Date']]
        report_df['Payment Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Payment Date']]
    file=file.replace('.xlsx','_import.xlsx')
    file_txt=file.replace('.xlsx','.txt')
    report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    return comment

def Payment_Selectronics(file):
    value=[]
    value2=0
    limit=[0,1,2,4,3,5,6]
    report_df = pd.read_excel(file, header=None)
    report_df=report_df[limit]
    column=list(report_df.iloc[0].dropna())
    report_df.columns=column
    report_df = report_df.drop(report_df.index[0])
    limit=[i for i in range(len(report_df['CK#'].dropna()))]
    report_df = report_df.iloc[limit]
    if(round(report_df['Check Amount'].sum(),2)==round(report_df['Amount Applied to Invoice'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Check Amount'].sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    report_df['Check Amount'] = report_df['Check Amount'].fillna(0)
    for i in report_df['Check Amount']:
        if(i != 0):
            value.append(i)
            value2=i
        else:
            value.append(value2)
    report_df['Check Amount']=value
    report_df['Check Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Check Date']]
    file=file.replace('.xlsx','_import.xlsx')
    file_txt=file.replace('.xlsx','.txt')
    report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    return comment  

def Payment_NT(file):
    value=[]
    n=[]
    n2=[]
    value2=0
    report_df = pd.read_excel(file, sheet_name='NT Payments', header=None) #Read the excel
    column=list(report_df.iloc[0]) # Name columns
    report_df.columns=column #Put the columns names with the correct name 
    report_df = report_df.drop(report_df.index[0]) #Take off the row 0
    limit=[i for i in range(len(report_df['Pay Type'].dropna()))] #see the range of the number of rows
    report_df = report_df.iloc[limit] #Take the rows that we need
    limit = column.index('Gross Amount') 
    report_df = report_df[column[:limit+1]]#Take the columns that we need
    if(round(report_df['Total Check Amount'].sum(),2)==round(report_df['Gross Amount'].sum(),2)):
        comment="The value amount match correctly " +str(round(report_df['Total Check Amount'].sum(),2))
    else:
        comment="The value amount doesn't match correctly"
    report_df['Total Check Amount'] = report_df['Total Check Amount'].fillna(0)
    cont=0
    for i in report_df['Total Check Amount']: #fill the 0 with the number also is taking the lines of each check amount
        cont=cont+1
        if(i != 0):
            n.append(cont)
            value.append(i)
            value2=i
        else:
            value.append(value2)
    n.append(len(report_df['Total Check Amount'])+1)
    report_df['Total Check Amount']=value
    report_df['Transaction Date'] = [i.strftime("%#m/%#d/%Y") for i in report_df['Transaction Date']]
    report_df['Total Check Amount'] = [round(i,2) for i in report_df['Total Check Amount']]
    file=file.replace('.xlsx','_import.xlsx')
    file_txt=file.replace('.xlsx','.txt')
    file_txt2=file_txt.replace('.txt','')
    report_df.to_excel(file,sheet_name='Payment Template', engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None, header=False,sep='\t')
    restart=n[0]
    for i in range(len(n)): #Divide the file txt in 8 checks amounts
        if i%8==0:
            if(n[i]-restart)>=1000:
                n2.append(n[i-4])
                n2.append(n[i])
            else:
                n2.append(n[i])
            restart=n[i]
    if not n[-1] in n2:
        n2.append(n[-1])
    for i in range(len(n2)-1):
        with open(file_txt, 'r') as fp:
            x = fp.readlines()[n2[i]-1:n2[i+1]-1]
            with open(file_txt2+'{0}'.format(i)+'.txt', 'w') as fout:
                fout.writelines(x)
    return comment

def call(company,file):
    if company=='American_Broadband':
        comment=Payment_ABB(file)
    elif company=='BendTel':
        comment=Payment_Bendtel(file)
    elif company=='GCI':
        comment=Payment_GCI(file)
    elif company=='MIEAC':
        comment=Payment_Mieac(file,'MIEAC Payments')
    elif company=='MTA':
        comment=Payment_Mta(file)
    elif company=='Selectronics':
        comment=Payment_Selectronics(file)  
    elif company=='ONVOY':
        comment=Payment_Mieac(file,'Onvoy End Office Payments')    
    print(comment)