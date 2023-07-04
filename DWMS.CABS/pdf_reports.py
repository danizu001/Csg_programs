import pdfplumber
import re
import PyPDF2
import pandas as pd
import openpyxl 
import os
from datetime import date, timedelta, datetime
import numpy as np

company=''
# This function takes the report location "Usage Balancing report" and return the total value of all FCCIDs for the 4 number [AUR, MOU Tabs, Billed, Difference AUR vs Billed]
# for normal usage balancing the bill type will be 'NA'. For MTA SW and RC we need to filter by BAN (RC BANs have O in the BAN)(SW BANS start with 3015D or 3015A)
# For MTA SW we use bill_type='SW'. For MTA RC we use bill_type='RC'
def usage_balancing(report_location, bill_type="NA"):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:
        total = [0, 0, 0, 0]  # total is a list to contain the totals of the 4 columns

        if bill_type == "NA":
            for page in pdf.pages:
                file = page.extract_text()  # reading the text of every page
                position = 0
                for t in range(file.count("Bill Fccid Total")):  # getting the number of "Bill Fccid Total"
                    text = file[file.find("Bill Fccid Total") : file.find("Bill Fccid Total")+ 100]  # saving the line of "Bill Fccid Total" into text
                    text_split = text.split()  # split every word in the text
                    
                    # adding the value to total
                    total[0] += float(re.sub("[^-\d\.]", "", text_split[4]))
                    total[1] += float(re.sub("[^-\d\.]", "", text_split[5]))
                    total[2] += float(re.sub("[^-\d\.]", "", text_split[6]))
                    # since MIEAC does not have the last total as a number, we do not include it
                    # then we check if the value is negative
                    if text_split[7].startswith('-')and company != 'MIEAC':
                        total[3] += float(re.sub("[^-\d\.]", "", text_split[7]))*-1
                    elif company!='MIEAC':
                        total[3] += float(re.sub("[^-\d\.]", "", text_split[7]))

                    position = file.find("Bill Fccid Total")  # getting the next "Bill Fccid Total" position
                    file = file[position + 100 : -1]  # adjusting the file "page" to exclude the counted "Bill Fccid Total"

        elif bill_type == "RC":
            for page in pdf.pages:
                file = page.extract_text()  # reading the text of every page
                position = 0

                for t in range(file.count("BAN Total ")):  # getting the number of "BAN Total "

                    text = file[file.find("BAN Total ") : file.find("BAN Total ") + 75]  # saving the line of "BAN Total " into text
                    text_split = text.split()  # split every word in the text

                    # copying only the BANS that have O in the BAN "RC BANs"
                    if text_split[2].find("O") > 0:

                        # adding the value to total
                        total[0] += float(re.sub("[^-\d\.]", "", text_split[3]))
                        total[1] += float(re.sub("[^-\d\.]", "", text_split[4]))
                        total[2] += float(re.sub("[^-\d\.]", "", text_split[5]))
                        if text_split[6].startswith('-')and company != 'MIEAC':
                            total[3] += float(re.sub("[^-\d\.]", "", text_split[6]))*-1
                        elif company!='MIEAC':
                            total[3] += float(re.sub("[^-\d\.]", "", text_split[6]))


                    position = file.find("BAN Total ")  # getting the next "BAN Total " position
                    file = file[position + 20 : -1]  # adjusting the file "page" to exclude the counted "BAN Total 3015O"

        elif bill_type == "SW":
            for page in pdf.pages:
                file = page.extract_text()  # reading the text of every page
                position = 0

                for t in range(file.count("BAN Total 3015")):  # getting the number of "BAN Total 3015"

                    text = file[file.find("BAN Total 3015") : file.find("BAN Total 3015") + 75]  # saving the line of "BAN Total 3015" into text
                    text_split = text.split()  # split every word in the text

                    # copying only the BANS that do not equal to 3015O "RC BANs"
                    if text_split[2].startswith("3015O") is False:
                        # adding the value to total
                        total[0] += float(re.sub("[^-\d\.]", "", text_split[3]))
                        total[1] += float(re.sub("[^-\d\.]", "", text_split[4]))
                        total[2] += float(re.sub("[^-\d\.]", "", text_split[5]))
                        if text_split[6].startswith('-')and company != 'MIEAC':
                            total[3] += float(re.sub("[^-\d\.]", "", text_split[6]))*-1
                        elif company!='MIEAC':
                            total[3] += float(re.sub("[^-\d\.]", "", text_split[6]))
                    position = file.find("BAN Total 3015")  # getting the next "BAN Total 3015" position
                    file = file[position + 20 : -1]  # adjusting the file "page" to exclude the counted "BAN Total 3015"

        pdf.close()  # closing the file
    return str(total)  # returning the four numbers


# This function takes the report location "Aged Trial Balance Report" and return the Grand Total value
def aged_trial_balance(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report
        total = file[file.find("GRAND TOTAL") : -1]  # getting the Grand Total part into total
        total_split = total.split()  # spliting the numbers into list
        pdf.close()

    # we convert the negative format ($1.00) to -1.00
    if total_split[7].endswith(')'):
        return float(re.sub("[^-\d\.]", "", total_split[7]))*-1
    else:
        return float(re.sub("[^-\d\.]", "", total_split[7]))


# This function is going to return the location of payment report and return the total
# This fuction does not do any kind of filter. we will use the excel version to do filtration
def payment(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report
        total = file[file.find("Page") - 30 : file.find("Page")]  # getting the Total part into total
        total_split = total.split()  # spliting the numbers into list
        pdf.close()

        if (total_split[-1] == "Date"):  # if it does not have a value and it has date it returns 0
            return 0       
        # we convert the negative format ($1.00) to -1.00
        elif total_split[-1].startswith('('):
            float(re.sub("[^-\d\.]", "", total_split[-1].replace("$", "")))*-1
        else:
            return float(re.sub("[^-\d\.]", "", total_split[-1].replace("$", "")))


# This function is to get the location of usage by day report and returns the total
def usage_by_day(report_location):

    total = 0
    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        for page in pdf.pages:
            file = page.extract_text()  # reading the text of every page
            position = 0
            for t in range(file.count("Totals for FCCID")):  # getting the number of "Totals for FCCID"
                text = file[file.find("Totals for FCCID") : file.find("Totals for FCCID") + 60]  # saving the line of "Totals for FCCID" into text
                text_split = text.split()  # split every word in the text

                # adding the value to total
                total += float(re.sub("[^-\d\.]", "", text_split[6]))

                position = file.find("Totals for FCCID")  # getting the next "Totals for FCCID" position
                file = file[position + 100 : -1]  # adjusting the file "page" to exclude the counted "Totals for FCCID"
        pdf.close()  # closing the file
    return total


# This function is going to return the total value of Adjustment report 
# This fuction does not do any kind of filter. we will use the excel version to do filtration
def adjustment(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report

        # checking if there is adjustment or not, if there is "Grand Total" it will return the value.If not it will return 0
        if file.find("Grand Total") > 0:
            total = file[file.find("Grand Total") : -1]  # getting the Grand Total part into total
            total_split = total.split()  # spliting the numbers into list
            # we convert the negative format ($1.00) to -1.00
            if total_split[2].startswith('('):
                return float(re.sub("[^-\d\.]", "", total_split[2].replace("$", "")))*-1
            else:    
                return float(re.sub("[^-\d\.]", "", total_split[2].replace("$", "")))
        else:
            return 0


# This function is going to return the total current revenue value of SW Bans Trending report 
def bans_trending_report(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report
        text = file[file.find("Totals") : -1]  # getting the text part of the last table "Total"
        revenue = text[text.find("Revenue") : -1]  # getting the text part of the revenue
        revenue_split = revenue.split()  # spliting the numbers into list
    return float(re.sub("[^-\d\.]", "", revenue_split[12]))


# This function is going to return the total value of OCCs report
# This fuction does not do any kind of filter. we will use the excel version to do filtration
def occ(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report

        # checking if there is OCC or not, if there is "Grand Total" it will return the value.If not it will return 0
        if file.find("Grand Total") > 0:
            total = file[file.find("Grand Total") : -1]  # getting the Grand Total part into total
            total_split = total.split()  # spliting the numbers into list
            # we convert the negative format ($1.00) to -1.00
            if total_split[2].startswith('('):
                return float(re.sub("[^-\d\.]", "", total_split[2]))*-1
            else:
                return float(re.sub("[^-\d\.]", "", total_split[2]))
        else:
            return 0
        pdf.close()


# This function takes the report location "transactions summary report" and return the total value of  "credit"
def transaction_summary(report_location,bill_type='NA'):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:
        total = [0,0]  # total is a list to contain the total of the 2 value "credit and debit"
        for page in pdf.pages:
            file = page.extract_text()  # reading the text of every page
            position = 0
            if bill_type == 'SW':
                for t in range(file.count('Switched "D" Bills')):
                    file = file[file.find('Switched "D" Bills'): -1]
                    text = file[file.find("TOTAL:") : file.find("TOTAL:") + 40]
                    text_split = text.split()  # split every word in the text
                    # adding the value to total
                    total[0] += float(re.sub("[^-\d\.]", "", text_split[1]))
                    total[1] += float(re.sub("[^-\d\.]", "", text_split[2]))

                    position = file.find("TOTAL:")  # getting the next "TOTAL:" position
                    file = file[position + 50 : -1]  # adjusting the file "page" to exclude the counted "TOTAL:"
            elif bill_type == 'FA':
                 for t in range(file.count('Facility Bills')):
                    file = file[file.find('Facility Bills'): -1]
                    text = file[file.find("TOTAL:") : file.find("TOTAL:") + 40]
                    text_split = text.split()  # split every word in the text
                    # adding the value to total
                    total[0] += float(re.sub("[^-\d\.]", "", text_split[1]))
                    total[1] += float(re.sub("[^-\d\.]", "", text_split[2]))

                    position = file.find("TOTAL:")  # getting the next "TOTAL:" position
                    file = file[position + 50 : -1]  # adjusting the file "page" to exclude the counted "TOTAL:"
            elif bill_type == 'RC':
                 for t in range(file.count('Switched "O" Bills')):
                    file = file[file.find('Switched "O" Bills'): -1]
                    text = file[file.find("TOTAL:") : file.find("TOTAL:") + 40]
                    text_split = text.split()  # split every word in the text
                    # adding the value to total
                    total[0] += float(re.sub("[^-\d\.]", "", text_split[1]))
                    total[1] += float(re.sub("[^-\d\.]", "", text_split[2]))

                    position = file.find("TOTAL:")  # getting the next "TOTAL:" position
                    file = file[position + 50 : -1]  # adjusting the file "page" to exclude the counted "TOTAL:":

            else :
                for t in range(file.count("TOTAL")):  # getting the number of "TOTAL:"
                    text = file[file.find("TOTAL:") : file.find("TOTAL:") + 40]  # saving the line of "TOTAL:" into text
                    text_split = text.split()  # split every word in the text
                    # adding the value to total
                    total[0] += float(re.sub("[^-\d\.]", "", text_split[1]))
                    total[1] += float(re.sub("[^-\d\.]", "", text_split[2]))

                    position = file.find("TOTAL:")  # getting the next "TOTAL:" position
                    file = file[position + 50 : -1]  # adjusting the file "page" to exclude the counted "TOTAL:"
            
        pdf.close()  # closing the file

        # checking if the 2 number are matching to return the value, if no return error.
        if total[0] == total[1]:
            return total[1]
        else:
            return "numbers are not matching"


# This function is going to return the total revenue value of MMR 
def mmr(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report
        text = file[file.find("Grand Totals") : -1]  # getting the text part of the last table "Grand Totals"
        revenue = text[text.find("Total Revenue") : -1]  # getting the text part of the revenue
        revenue_split = revenue.split()  # spliting the numbers into list

    return float(re.sub("[^-\d\.]", "", revenue_split[9]))


# This function takes the report location "transactions summary report" and return the total value of  "late payment charges"
def late_payment_charge_from_TS(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:
        total = 0
        for page in pdf.pages:
            file = page.extract_text()  # reading the text of every page
            position = 0
            for t in range(file.count("Late Payment Charges")):  # getting the number of "Late Payment Charges"
                text = file[file.find("Late Payment Charges") : file.find("Late Payment Charges")+ 40]  # saving the line of "Late Payment Charges" into text
                text_split = text.split()  # split every word in the text

                # adding the value to total
                if text_split[4].startswith('('):
                    total -= float(re.sub("[^-\d\.]", "", text_split[4]))
                else:
                    total += float(re.sub("[^-\d\.]", "", text_split[4]))

                position = file.find("Late Payment Charges")  # getting the next "Late Payment Charges" position
                file = file[position + 40 : -1]  # adjusting the file "page" to exclude the counted "Late Payment Charges"
        pdf.close()  # closing the file
    return total


# This function is going to return the total value of ADJ and OCCs report 
# This fuction does not do any kind of filter
def adj_occ_for_AB(report_location):

    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:

        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report

        # checking if there is adj or not, if there is "Total ADJ:" it will return the value. If not it will return 0
        if file.find("Total ADJ:") > 0:
            total = file[file.find("Total ADJ:") : -1]  # getting the Grand Total part into total
            total_split = total.split()  # spliting the numbers into list
            # we convert the negative format ($1.00) to -1.00
            if total_split[2].startswith('('):
                return float(re.sub("[^-\d\.]", "", total_split[2].replace("$", "")))*-1
            else:
                return float(re.sub("[^-\d\.]", "", total_split[2].replace("$", "")))
        else:
            return 0


# This function is going to return the total value of minutes and revenue by rate element report 
def min_rev_by_rate_element(report_location):
    # opening the report with pdfplumber as pdf
    with pdfplumber.open(report_location) as pdf:
        file = pdf.pages[-1].extract_text()  # reading the text of the last page of the report
        file = file[file.find("page")-250:file.find("page")-1]  # getting the text part of the last table "Total"
        text = file[file.find("Total"):-1]
        text_split = text.split()  # spliting the numbers into list
        # we convert the negative format ($1.00) to -1.00
    if text_split[12].startswith('('):
        return float(re.sub("[^-\d\.]", "", text_split[12].replace("$", "")))*-1
    else:
        return float(re.sub("[^-\d\.]", "", text_split[12].replace("$", ""))) 


# checks if exception analysis is empty or not
def exception_analisys(report_location):
    value=''
    # open the pdf file
    object = PyPDF2.PdfFileReader(report_location)

    # define keyterms
    String = "No Exceptions to report"

    # searching if report is empty
    PageObj = object.getPage(0)
    Text = PageObj.extractText() 
    ResSearch = re.search(String, Text)
    if ResSearch:
        value='0'
    return(value)
    

# this function checks if the mou entries report is empty or not.
#If the report is not empty, it will write in the comment of the MPS what values appear in the report
def mou_entries(report_location, uid_values):

    # openeing the pdf report
    pdf=report_location.replace('txt','pdf')
    pdf_file = open(pdf,'rb')
    read_pdf = PyPDF2.PdfFileReader(pdf_file)
    page = read_pdf.getPage(0)
    page_content = page.extractText()

    # getting the UIDs to match it with what is in the MPS
    location_starting = page_content.find('s') #This part look the range of the FCCIDs in the PDF to validate with the MPS
    location_first_colom = page_content.find(':')
    location_FFCID = page_content.find('F')
    location_second_colom = page_content.find(':',location_first_colom+1)
    final_uid = int(page_content[location_first_colom+1:location_starting-1])
    begining_uid = int(page_content[location_second_colom+1:location_FFCID-1])


    if(begining_uid== uid_values[0] and final_uid== uid_values[1]):
        try:
            tt_result_percentage = []
            txtf = pd.read_csv(report_location, sep=",",header=None,skiprows=1)
            data_txtDF = pd.DataFrame(txtf)
            for i in range(11):# Split the name of the column with the value and olny take the values
                data_txtDF[i] = data_txtDF[i].str.split("=",1).apply(lambda l: "".join(l[1]))
            data_txtDF.columns=['BILL FCCID','CIC','BAN','CIC_LIST','CLLI','JURIS','TT','USAGE PERIOD','TANDEM MESSAGES','MINUTES','RATE PERIOD'] #Put names for each column
            data_txtDF['MINUTES'] = round(pd.to_numeric(data_txtDF['MINUTES']),2) #Round all the column of minutes
            value = round(data_txtDF['MINUTES'].sum()) #Start the stadistic
            sum_min_per_jur = data_txtDF.groupby('JURIS').MINUTES.apply(lambda g: g.sum()).reset_index() #Sum all the minutes per Jurisdiction with the properly clasification
            juris_percentage = round(100*(sum_min_per_jur['MINUTES']/value),2) #Put the percentage per Jurisdiction
            min_per_tt = data_txtDF.groupby(['JURIS','TT']).MINUTES.apply(lambda g: g.sum()).reset_index() #Clasify the TT per jurisdiction
            comment2=""
            comment=""
            min_per_tt['TT']=min_per_tt['TT'].replace(' 01','01: 700').replace(' 02','02: 800').replace(' 03','03: 900').replace(' 04','04: OPH(0+)').replace(' 05','05: OPH(0-)').replace(' 06','06: International DD').replace(' 07','07: Domestic DD').replace(' 08','08: International OPH').replace(' 09','09: Directory Assistance').replace(' 10','10: Terminating Minutes').replace(' 11','11: Terminating 800').replace(' 12','12:	WATS').replace(' 13','13: Switched 56 Kbps').replace(' 14','14: 800 Database-POTS Translation').replace(' 15','15: 800 Database-IC Identification').replace(' 16','16: Originating Minutes').replace(' 17','17: Orig. Minutes - Non-Jurisdictional').replace(' 18','18: Domestic DDD - Non-Jurisdictional').replace(' 27','27: 500').replace(' 98','98: Terminating Local - EAS').replace(' 99','99: Terminating Local')
            for i in range(len(sum_min_per_jur['JURIS'])): #This loop is for put the percentage for each TT in each Jurisdiction
                for j in range(len(min_per_tt['JURIS'])):
                    if(sum_min_per_jur['JURIS'][i]==min_per_tt['JURIS'][j]):
                        tt_result_percentage.append(round(100*(min_per_tt['MINUTES'][j]/sum_min_per_jur['MINUTES'][i]),2))
            for i in range(len(sum_min_per_jur['JURIS'])): #This loop is for do the comment for the excel cell
                for j in range(len(tt_result_percentage)):
                    if(sum_min_per_jur['JURIS'][i]==min_per_tt['JURIS'][j]):
                        comment2= comment2 + min_per_tt['TT'][j] +": " + str(tt_result_percentage[j]) +"% "
                comment= comment + "The Jurisdiction: " + str(sum_min_per_jur['JURIS'][i]) + " has "+ str(juris_percentage[i]) +"% [TT: " + str(comment2) + "]\n"
                comment2=""

                
        except Exception:
            txtf = pd.read_csv(report_location, sep='\t')
            data_txtDF = pd.DataFrame(txtf)
            if(begining_uid== uid_values[0] and final_uid== uid_values[1]):
                value=0
                comment = "The UIDs are matching correctly"
        
        excel=report_location.replace('txt','xlsx')
        data_txtDF.to_excel(excel,index=False)

    else:
        value = "The UIDs don't match"
        comment = 0
        
    pdf_file.close()
    os.remove(report_location)
    return value , comment


# this function checks if adj not posted report is empty or not.
def adjs_not_posted(report_location):
    value=''
    # open the pdf file
    object = PyPDF2.PdfFileReader(report_location)

    # define keyterms
    String = "Fccid"

    # searching if report is empty
    PageObj = object.getPage(0)
    Text = PageObj.extractText() 
    ResSearch = re.search(String, Text)
    if ResSearch:
        comment='Adjs not posted is not empty please post before continuing'
    else:
        value='X'
        comment=0
    return value, comment


# this function checks if bill completion report is empty or not.
def bill_completion(report_location):
    value=''
    # open the pdf file
    object = PyPDF2.PdfFileReader(report_location)

    # define keyterms
    String = "Bill Fccid"

    # searching if report is empty
    PageObj = object.getPage(0)
    Text = PageObj.extractText() 
    ResSearch = re.search(String, Text)
    if ResSearch:
        comment='Please check the report, there are errors'
    else:
        value='X'
        comment=0
    return value, comment


#This function gets the location of the report, but from the previous billing month
def get_last_month_file(file):
    file_array = file.split('_')
    file_date = file_array[-3]
    folder_array = file.split('\\')
    folder_year = int(folder_array[4])
    month = int(file_date[:2]) 
    year = int(file_date[-2:]) 
    if month == 1:
        year = year-1
        folder_year = str(folder_year-1)
        month = 12
    else:
        month = month -1
        folder_year = str(folder_year)

        
    file_date = file_date.replace(file_date[:2],str(month).zfill(2),1)
    file_date = str(year).zfill(2).join(file_date.rsplit(file_date[-2:],1))
    file_array[-3] = file_date
    folder_month = str(month).zfill(2)+'_'+folder_year
    file_old = '_'.join(file_array)
    file_old = file_old.split('\\')
    folder_array[-1] = file_old[-1]
    folder_array[4] = folder_year 
    folder_array[5] = folder_month
    folder_old = '\\'.join(folder_array)
    
    return folder_old

# This function checks if the Factor or SOC change audit report is empty or not.
# it also checks if the dates are correct or not.
def soc_factor_change(code,report_location):
    
    if report_location.endswith('.xlsx'):
            report_location=report_location.replace('xlsx','pdf')

    # opening current month report
    with pdfplumber.open(report_location) as pdf:
        page = pdf.pages[0].extract_text().split()
        if code == '10':    # if the report is Factor changes
            word = 'between'
            from_date = 1
            to_date = 3
            current_produced = datetime.strptime(page[0],"%M/%d/%Y")
        else:   # If the report is SOC changes
            word = 'Updated'
            from_date = 3
            to_date = 6
            current_produced = datetime.strptime(pdf.pages[-1].extract_text().split()[-2],"%M/%d/%Y")

        # getting from and to dates from the pdf
        from_date_current =  datetime.strptime(page[page.index(word)+from_date], "%M-%d-%Y")
        to_date_current =  datetime.strptime(page[page.index(word)+to_date],"%M-%d-%Y")

    # get the dates info from previous month report
    old_month_report = get_last_month_file(report_location)
    with pdfplumber.open(old_month_report) as pdf2:
        page2 = pdf2.pages[0].extract_text().split()
        to_date_prev =  datetime.strptime(page2[page2.index(word)+to_date],"%M-%d-%Y")
    
    # check if the dates are correct
    if from_date_current == to_date_prev + timedelta(days=1) and current_produced == to_date_current:
        value='X'
        # check is the report is empty or not
        if page.count('FCCID') >2 or 'CIC' in page:
            comment='The dates are right and there are changes this month'
        else:
            comment='The dates are right and the report is empty'  
    else:
        value='X'
        comment='The paramemter dates for the report are wrong'

    return value,comment

def sbt_gci(occ_report):
    with pdfplumber.open(occ_report) as pdf:
        page=[]
        bool_bill=[]
        for i in range(len(pdf.pages)):
            page+=(pdf.pages[i].extract_text().split())
        index_amount=[i for i,val in enumerate(page) if val=='Company'][1:]
        index_ban=[i for i,val in enumerate(page) if val=='Date']
        index_billed=[i for i,val in enumerate(page) if val=='billed'][1:]
        index_amount=list(map(lambda x: x + 7, index_amount))
        index_ban=list(map(lambda x: x + 4, index_ban))
        nppage = np.array(page)
        index_delbilled=[i for i,val in enumerate(nppage[index_amount]) if val=='billed']
        data_amount=list(nppage[index_amount])
        data_ban=list(nppage[index_ban])
        for i in index_billed:
            if page[i-26]=='Company':
                bool_bill.append(True)
            else:
                bool_bill.append(False)
        for i in range(len(index_delbilled)):
            if bool_bill[i]==False:
                data_amount[index_delbilled[i]]=0
                data_ban[index_delbilled[i]]=0
            else:
                data_amount[index_delbilled[i]]=0
        data_amount=list(filter(lambda num: num != 0, data_amount))
        data_ban=list(filter(lambda num: num != 0, data_ban))
        df_data={'ban':data_ban,'amount_billed':data_amount}
        df = pd.DataFrame(df_data)
        return(df)



    