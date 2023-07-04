import pandas as pd
import math
def Adjustments_Bendtel(file,inv_bal):
    bill_date=[]
    report_df = pd.read_excel(file, header=None)
    inv_bal_df = pd.read_excel(inv_bal)
    limit=report_df[0].tolist().index('Customer')
    report_df = report_df[limit:]
    column_name=report_df.loc[limit]
    report_df.columns=column_name
    report_df = report_df.drop(limit)
    report_df['SOC']=report_df['SOC'].fillna(method='ffill')
    report_df=report_df.dropna(subset=['Invoice Number']).reset_index(drop=True)
    report_df=report_df.fillna(method='ffill')
    invoices=list(report_df['Invoice Number'])
    for i in invoices:
        bill_date.append(inv_bal_df['bill_date'].where(inv_bal_df['user_invoice_num'] == i).dropna().reset_index(drop=True).loc[0])
    report_df['Effective Date'] = bill_date
    report_df['Effective Date'] = [i.strftime("%#m/%#d/%Y %H:%M") for i in report_df['Effective Date']]
    column_name=['BAN','Invoice Number', 'Effective Date', 'Phrase Code', 'Description','PON','Phrase Code','SOC','Total Adjustment','Interstate Amount','Intrastate Amount','Local Amount']
    report_df=report_df[column_name]
    column_name[2]='Bill_Date'
    file=file.replace('.xls','_import.xls')
    file_txt=file.replace('.xls','.txt')
    report_df.to_excel(file, engine='xlsxwriter', index=False)
    report_df.to_csv (file_txt, index = None,sep='\t') 
    comment='good'
    return comment
def Adjustments_ABB(file):
    pass
def Adjustments_GCI(file):
    pass
def Adjustments_Mieac(file):
    pass
def Adjustments_Mta(file):
    pass
def Adjustments_Selectronics(file):
    pass
def Adjustments_NT(file):
    pass
company=input("Select the company with the number:\n1. ABB \n2. Bendtel \n3. GCI \n4. Mieac\n5. Mta\n6. Selectronics\n7. Neutral Tandem\n")
file=input("Enter the path of the adjustment\n")
inv_bal=input("Enter the path of the invoice balance\n")
if company=='1':
    comment=Adjustments_ABB(file)
elif company=='2':
    comment=Adjustments_Bendtel(file, inv_bal)
elif company=='3':
    comment=Adjustments_GCI(file)
elif company=='4':
    comment=Adjustments_Mieac(file)
elif company=='5':
    comment=Adjustments_Mta(file)
elif company=='6':
    comment=Adjustments_Selectronics(file) 
elif company=='7':
    comment=Adjustments_NT(file)    
print(comment)