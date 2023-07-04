from enum import unique
from tkinter import *
from tkinter import ttk
import tkinter.font as font
import database as db
import os,shutil
import clients
import reports
from tkcalendar import Calendar
from datetime import date, datetime
from tkinter.ttk import Progressbar
import time
import subprocess
import warnings
from shutil import copyfile
from os import listdir
from os.path import isfile, join

# Disable showing warnings

class mainapp(Frame):
    def __init__(self, master):
        #self.master = master
        
        self.canvas = Canvas(background='grey31',height=720,highlightthickness=0)
        #canvas.create_line(15, 25, 200, 25)
        self.canvas.create_rectangle(0, 0, 140, 700,
            outline="grey10", fill="grey10",)
        self.canvas.place(x=0)

        # Creating labels
        self.title = Label(master, text="CABS",bg='grey31',fg='light cyan')
        self.title.place(x=400, y=20)
        self.title.config(font=("Segoe UI Historic", 44 ))
        self.reports = Label(master, text="Reports\n to be printed:",bg='grey10',fg='light cyan')
        self.reports.place(x=10, y=10)
        self.reports.config(font=("Segoe UI Historic", 12 ))
        self.analyst = Label(root, text="Analyst\nAssistant",bg='grey31',fg='light cyan')
        self.analyst.place(x=550, y=15)
        self.analyst.config(font=("Segoe UI Historic", 28 ))
        self.company = Label(root, text="1. Customer",bg='grey31')
        self.company.place(x=141, y=90)
        self.company.config(font=("Segoe UI Light", 24))
        self.bill_type = Label(root, text="2. Bill Type",bg='grey31')
        self.bill_type.place(x=141, y=250)
        self.bill_type.config(font=("Segoe UI Light", 24))
        self.group = Label(root, text="3. Group",bg='grey31')
        self.group.place(x=500, y=250)
        self.group.config(font=("Segoe UI Light", 24))
        self.dates = Label(root, text="4. Dates",bg='grey31')
        self.dates.place(x=141, y=410)
        self.dates.config(font=("Segoe UI Light", 24))
        self.credits = Label(root, text="Developed by: Nicolas, Rafek, Daniel",bg='grey31')
        self.credits.place(x=370, y=675)
        self.credits.config(font=("Verdana", 10))
        
        #Creating companies radio buttons
        self.v = IntVar()
        self.v.set(None)  # initializing the choice, i.e. Python

        companies = [("Selectronics", 1),
                ("Bendtel", 2),
                    ("MIEAC", 3),
                    ("Onvoy", 4),
                    ("MTA", 5),
                    ("American Broadband", 6),
                    ("Peerless", 7),
                    ("Neutral Tandem", 8),
                    ("GCI",9)]
        #Creating a radio button for each company   
        y=130
        y2=130
        y3=130
        self.company_rad_butt=[]
        self.buttonFont = font.Font(family='Segoe UI Light', size=14, weight='bold')
        for comp, val in companies:
            self.company_rad=Radiobutton(root, 
                        text=comp,
                        foreground='mint cream', 
                        variable=self.v, 
                        command=self.select_company,
                        background='grey31',highlightthickness=0,font=self.buttonFont,
                        value=val, selectcolor='light sea green',activebackground='grey31',activeforeground='mint cream')
            if len(self.company_rad_butt) <= 2:
                x=141
                y=y+30
                self.company_rad.place(x=x,y=y)
            elif len(self.company_rad_butt) <= 5:
                x=330
                y2=y2+30
                self.company_rad.place(x=x,y=y2)
            else:
                x=550
                y3=y3+30
                self.company_rad.place(x=x,y=y3)
            self.company_rad_butt.append(self.company_rad)

        #Creating companies radio buttons for BILL TYPES
        self.j = IntVar()
        self.j.set(None)  # initializing the choice, i.e. Python

        billtypes = [("Switched", 1),
                ("Facility", 2),
                    ("Recip Comp", 3),
                    ("Ohio", 4),
                    ("Michigan", 5),
                    ("Indiana", 6)]
        #Creating a radio button for each BILL TYPE   
        y4=280
        y2=280
        self.billtypes_rad_butt=[]
        self.buttonFont = font.Font(family='Segoe UI Light', size=14, weight='bold')
        for bt, val in billtypes:
            y=y+30
            self.billtypes_rad=Radiobutton(root, 
                        text=bt,
                        foreground='mint cream', 
                        variable=self.j, 
                        command=self.select_bill_type,
                        background='grey31',highlightthickness=0,font=self.buttonFont, 
                        value=val,selectcolor='light sea green',activebackground='grey31',activeforeground='mint cream')
            if len(self.billtypes_rad_butt) <= 2:
                x=141
                y4=y4+30
                self.billtypes_rad.place(x=x,y=y4)
            else:
                x=330
                y2=y2+30
                self.billtypes_rad.place(x=x,y=y2)
            self.billtypes_rad_butt.append(self.billtypes_rad)


        #Creating companies radio buttons for GROUPS
        self.p = IntVar()
        self.p.set(None)  # initializing the choice, i.e. Python

        groups = [("Pre Production (FGA)", 1),
                ("Post Production", 2),
                    ("Post Thresholding\nPost Deletion", 3)]
        #Creating a radio button for each company   
        y=280
        self.group_rad_butt=[]
        self.buttonFont = font.Font(family='Segoe UI Light', size=14, weight='bold')
        for gr, val in groups:
            y=y+30
            self.group_rad=Radiobutton(root, 
                        text=gr,
                        foreground='mint cream', 
                        variable=self.p,
                        background='grey31',highlightthickness=0,font=self.buttonFont, 
                        command=self.select_group,
                        value=val,selectcolor='light sea green',activebackground='grey31',activeforeground='mint cream')
            self.group_rad.place(x=550,y=y)
            self.group_rad_butt.append(self.group_rad)


        # Creating a Button
        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.start =Button(root,text='Start',height = 1,relief=FLAT, 
                width = 15,font=self.buttonFont,fg='white', bg='DarkSlateGray4',bd=0, highlightthickness =0,command = self.get_answer)
        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.rename =Button(root,text='Move Files',height = 1,relief=FLAT,
                width = 15,font=self.buttonFont,fg='white', bg='DarkOrange1',bd=0, highlightthickness =0,command = self.move_files_parms)
        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.quit = Button(root, text = 'Quit', command = root.destroy,height = 1,relief=FLAT, 
                width = 15,font=self.buttonFont,fg='white',bd=0, highlightthickness =0, bg='red4')
        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.reset = Button(root, text = 'Reset', command = self.reset, height = 1,relief=FLAT, 
                width = 15,font=self.buttonFont,fg='white',bd=0, highlightthickness =0, bg='steel blue')


        # Set a relative position of button
        self.rename.place(x=600, y=500)
        self.start.place(x=410, y=500)
        self.quit.place(x=600, y=590)
        self.reset.place(x=410, y=590)


        # Progress bar widget
        self.style = ttk.Style(root)
        # add label in the layout
        self.style.layout('text.Horizontal.TProgressbar', 
                    [('Horizontal.Progressbar.trough',
                    {'children': [('Horizontal.Progressbar.pbar',
                                    {'side': 'left', 'sticky': 'ns'})],
                        'sticky': 'nswe'}), 
                    ('Horizontal.Progressbar.label', {'sticky': ''})])
        # set initial text
        self.style.configure('text.Horizontal.TProgressbar', text='0 %')

        self.progress = Progressbar(root, orient = HORIZONTAL,
        length = 100, mode = 'determinate', style='text.Horizontal.TProgressbar')
        self.progress.place(x=500,y=450,width=185)


        #Getting todays date to set it in the calendar
        todays_date = date.today()
        year=todays_date.year
        month=todays_date.month
        day=todays_date.day

        # Add Calendar
        self.cal = Calendar(root, selectmode = 'day',
               year = year, month = month,
               day = day,date_pattern="y-mm-dd" )
        self.cal.place(x=150, y=470)
        #Array to store the report labels
        self.report_labels=[]

    #Method to update the progress bar
    def bar(self):
        #updating progress bar
        self.progress['value'] = 20
        #Updating the number over the progress bar
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(20))
        root.update_idletasks()
        time.sleep(1)
    
        self.progress['value'] = 40
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(40))
        root.update_idletasks()
        time.sleep(1)
    
        self.progress['value'] = 50
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(50))
        root.update_idletasks()
        time.sleep(1)
    
        self.progress['value'] = 60
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(60))
        root.update_idletasks()
        time.sleep(1)
    
        self.progress['value'] = 80
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(80))
        root.update_idletasks()
        time.sleep(1)

        self.progress['value'] = 100
        self.style.configure('text.Horizontal.TProgressbar', 
                    text='{:g} %'.format(100))




    def select_company(self): #Disables bill type depending on the company
        if self.v.get() == 1:
            self.company="Selectronics"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0:
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
            self.group_rad_butt[2].configure(state = DISABLED)
        elif self.v.get() == 2:
            self.company="BendTel"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0:
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 3:
            self.company="MIEAC"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0:
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 4:
            self.company="ONVOY"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0:
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 5:
            self.company="MTA"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0 and i != 1 and i != 2 :
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 6:
            self.company="American_Broadband"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 3 and i != 4 and i != 5 :
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 7:
            self.company="Peerless_Network"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0 :
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
        elif self.v.get() == 8:
            self.company="Neutral_Tandem"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0 and i != 1 :
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
                self.group_rad_butt[2].configure(state = DISABLED)
        elif self.v.get() == 9:
            self.company="GCI"
            i=0
            for radio_button in self.billtypes_rad_butt:
                if i != 0 and i != 1 :
                    self.billtypes_rad_butt[i].configure(state = DISABLED)
                i=i+1
                self.group_rad_butt[2].configure(state = DISABLED)
        self.billtype=''
        self.part=''
        return(self.company)  

    def select_bill_type(self): #Depending on the app disables the group depending on the bill type
        if self.j.get() == 1:
            if self.company == 'MTA' or self.company == 'Neutral_Tandem' or self.company == 'GCI'  :
                self.billtype="SW"
                self.group_rad_butt[2].configure(state = DISABLED)
            else:
                self.billtype= ""
        elif self.j.get() == 2:
            if self.company == 'MTA' or self.company == 'Neutral_Tandem' or self.company == 'GCI'  :
                self.billtype="FA"
            self.group_rad_butt[0].configure(state = DISABLED)
            self.group_rad_butt[2].configure(state = DISABLED)
        elif self.j.get() == 3:
            self.billtype="RC"
        elif self.j.get() == 4:
            self.billtype="509BOH"
        elif self.j.get() == 5:
            self.billtype="356DMI"
        elif self.j.get() == 6:
            self.billtype="590GIN"
        return(self.billtype)  

    def select_group(self): #Getting the report names depending on the company bill type and group
        if self.p.get() == 1:
            self.part="PRE"
        elif self.p.get() == 2:
            self.part="POST1"
        elif self.p.get() == 3:
            self.part="POST2"
        #Getting reports
        reportnames=self.get_reports()
        y=50
        
        for report in reportnames:
            y=y+40
            self.report_label=Label(root,text=report,bg="grey10",fg="light cyan")
            self.report_label.place(x=10,y=y)
            self.report_labels.append(self.report_label)
        return(self.part)
    
    def get_reports(self):
        #Add new reports here
        reports=[(0,'BAN Error Report'),(1,'CLLI Error Report'),(2,'Usage By Day'), (3,'Usage by Day 3 Month'), (4,'Zip and Error'),(5,'UBAL'), (6,'Payments'),
                (7,'SWBT') ,(8,'FABT'),(9,'Adjustment Report'),(10,'OCC Report'),(11,'ATB PDF'),(12,'ATB EXCEL'),(13,'MMR BY BAN PDF'),(14,'Transaction Summary'),
                (15,'LPC'),(16,'Revenue Analisys'),(17,'Billed By Period'),(18,'FA Circuit Charges\nBilled Disconnected'), (19,'Invoice Balance'),(20,'Adjustment and \nOCC Report'),
                (21,'MMR With LPC'), (22,'MMR EXCEL'),(23,'MMR By CLLI'),(24,'Billing Review'),(25,'MMR by Rate Element'), (26,'FA Summary Charges'),
                (27,'FA Charges by CIC\nby CLLI'),(28,'FUSC Charges by\nCircuit'),(29,'Accounting Detailed'),(30,'Accounting Report'),(31,'SW Usage Sumary \nCharges'),
                (32,'Exception Analisys'),(33,'MOU Entries Error'),(34,'Adj Not Posted'),(35,'Active BANs'),(36,'MMR Bill Date'),(37,'GL Extract'),(38,'Factor Change'),(39,'SOC Change'),
                (40,'Number Of Invoices'),(41,'AUR No/With Bill Date'),(42,'OCC Billed'),(43,'Bill Completion'),(44,'Pre Deletions Under 5'),(45,'Pos Deletions Under 5')]
        comp_type=self.company.upper()+self.billtype.upper()+self.part.upper()
        if comp_type == "SELECTRONICSPRE":
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[5],reports[6],reports[32],reports[33],reports[34]]
        elif comp_type == 'SELECTRONICSPOST1':
            unique_reports=[reports[7],reports[8],reports[9],reports[10],reports[11],reports[12],reports[13],reports[14],reports[15],reports[16],reports[17],reports[18],
                            reports[19],reports[40],reports[43]]
        elif comp_type == 'BENDTELPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[5],reports[6]]
        elif comp_type == 'BENDTELPOST1':
            unique_reports=[reports[5],reports[7],reports[9],reports[10],reports[43],reports[44]]
        elif comp_type == 'BENDTELPOST2':
            unique_reports=[reports[45],reports[5],reports[7],reports[13],reports[14],reports[17],reports[19],reports[20],reports[40]]
        elif comp_type == 'MIEACPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34],reports[38],reports[39]]
        elif comp_type == 'MIEACPOST1':
            unique_reports=[reports[7],reports[9],reports[10],reports[11],reports[13],reports[14],reports[43]]
        elif comp_type == 'MIEACPOST2':
            unique_reports=[reports[7],reports[11],reports[12],reports[13],reports[14],reports[16],reports[17],reports[19],reports[21],reports[22],reports[23],reports[24],reports[40],reports[41]]
        elif comp_type == 'ONVOYPRE':
            unique_reports=[reports[0],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34],reports[38],reports[39]]
        elif comp_type == 'ONVOYPOST1':
            unique_reports=[reports[7],reports[9],reports[10],reports[11],reports[13],reports[14],reports[43]]
        elif comp_type == 'ONVOYPOST2':
            unique_reports=[reports[7],reports[11],reports[12],reports[13],reports[14],reports[16],reports[17],reports[19],reports[21],reports[22],reports[23],reports[24],reports[40],reports[41]]
        elif comp_type == 'MTASWPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[5],reports[38],reports[39]]
        elif comp_type == 'MTASWPOST1':
            unique_reports=[reports[6],reports[7],reports[9],reports[10],reports[5],reports[12],reports[13],reports[23],reports[25],reports[14],reports[17],reports[31],reports[43]]
        elif comp_type == 'MTARCPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[5]]
        elif comp_type == 'MTARCPOST1':
            unique_reports=[reports[5],reports[6],reports[7],reports[9],reports[10],reports[12],reports[13],reports[23],reports[25],reports[14],reports[17],reports[31],reports[43]]
        elif comp_type == 'MTAFAPOST1':
            unique_reports=[reports[6],reports[8],reports[9],reports[10],reports[12],reports[14],reports[26],reports[27],reports[28],reports[42],reports[43]]
        elif comp_type == 'GCISWPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34],reports[38],reports[39]]
        elif comp_type == 'GCISWPOST1':
            unique_reports=[reports[6],reports[7],reports[9],reports[10],reports[5],reports[12],reports[13],reports[14],reports[17],reports[23],reports[25],reports[31],reports[40],reports[43]]
        elif comp_type == 'GCIFAPOST1':
            unique_reports=[reports[6],reports[8],reports[9],reports[18],reports[26],reports[27],reports[10],reports[14],reports[12],reports[28],reports[40],reports[42],reports[43]]
        elif comp_type == 'AMERICAN_BROADBAND590GINPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34]]
        elif comp_type == 'AMERICAN_BROADBAND590GINPOST1':
            unique_reports=[reports[5],reports[7],reports[20],reports[43],reports[44]]
        elif comp_type == 'AMERICAN_BROADBAND590GINPOST2':
            unique_reports=[reports[45],reports[5],reports[7],reports[11],reports[12],reports[13],reports[14],reports[17],reports[40]]
        elif comp_type == 'AMERICAN_BROADBAND356DMIPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34]]
        elif comp_type == 'AMERICAN_BROADBAND356DMIPOST1':
            unique_reports=[reports[5],reports[7],reports[20],reports[43],reports[44]]
        elif comp_type == 'AMERICAN_BROADBAND356DMIPOST2':
            unique_reports=[reports[45],reports[5],reports[7],reports[11],reports[12],reports[13],reports[14],reports[17],reports[40]]
        elif comp_type == 'AMERICAN_BROADBAND509BOHPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[6],reports[32],reports[33],reports[34]]
        elif comp_type == 'AMERICAN_BROADBAND509BOHPOST1':
            unique_reports=[reports[5],reports[7],reports[20],reports[43],reports[44]]
        elif comp_type == 'AMERICAN_BROADBAND509BOHPOST2':
            unique_reports=[reports[45],reports[5],reports[7],reports[11],reports[12],reports[13],reports[14],reports[17],reports[40]]
        elif comp_type == 'NEUTRAL_TANDEMSWPRE':
            unique_reports=[reports[0],reports[1],reports[2],reports[3],reports[4],reports[5],reports[6],reports[32],reports[33],reports[34]]
        elif comp_type == 'NEUTRAL_TANDEMSWPOST1':
            unique_reports=[reports[5],reports[7],reports[9],reports[10],reports[11],reports[12],reports[13],reports[22],reports[23],reports[14],reports[16],reports[17],reports[40],reports[43]]
        elif comp_type == 'NEUTRAL_TANDEMFAPOST1':
            unique_reports=[reports[6],reports[8],reports[9],reports[18],reports[26],reports[27],reports[10],reports[14],reports[12],reports[22],reports[29],reports[30],reports[40],reports[43]]
        elif comp_type == 'PEERLESS_NETWORKPRE':
            unique_reports=[reports[1],reports[3],reports[4],reports[5],reports[6],reports[33],reports[32],reports[34]]
        elif comp_type == 'PEERLESS_NETWORKPOST1':
            unique_reports=[reports[35],reports[7],reports[16],reports[5],reports[12],reports[13],reports[14],reports[20],reports[9],reports[10],reports[43]]
        elif comp_type == 'PEERLESS_NETWORKPOST2':
            unique_reports=[reports[5],reports[7],reports[11],reports[12],reports[22],reports[13],reports[14],reports[16],reports[17],reports[7],reports[19],reports[36],
            reports[37],reports[40]]

        report_names = [ j for i, j in unique_reports ]
        return report_names

    def move_files(self,array,path):
        mypath=path #Regular folder
        mypath1=path+'Send to Client\\' #Send to client folder
        for report in array:
            try:
                onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f)) ] #Getting all files inside a list
                for f in onlyfiles:
                    file=f.split('_')
                    if len(file)>2: #Only will work on reports not other kinds of files
                        if self.company.upper()+self.billtype.upper() !='NEUTRAL_TANDEMSW':
                            if '_'+file[-2].upper()==report or '_'+file[-2].upper()==report+"-3MONTHS" : #Checking that the code matches one of the ones on the list in move files params
                                newfile=f.replace(report,'')#Taking out the code from the old file
                                source=mypath+f #Full source path (path+file name) 
                                dest=mypath1+newfile#Full destination path (path+file name) 
                                os.rename(source,dest)#Renaming the file (and moving)
                                if report.find("19") == -1:
                                    copyfile(dest, source)#Copying the file from the send to client folder to regular folder
                                print(file[-1]+" has been renamed and moved")
                                db.log_file.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + ' | '+file[-1]+" has been renamed and moved\n")
                        else:
                            if file[1] != 'TTS':
                                ntpath=mypath1+'OSA-TSA Reports\\'
                            else:
                                ntpath=mypath1+'TTS Reports\\'
                            if '_'+file[-2].upper()==report or '_'+file[-2].upper()==report+"-3MONTHS" :
                                newfile=f.replace(report,'')
                                source=mypath+f
                                dest=ntpath+newfile
                                os.rename(source,dest)
                                if report.find("19") == -1:
                                    copyfile(dest, source)
                                print(file[1]+" "+file[-1]+" has been renamed and moved")
                                db.log_file.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + ' | '+file[1]+" "+file[-1]+" has been renamed and moved\n")
            except:
                print("Report "+report+" not found or there is a copy already in the send to client folder")
                db.log_file.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + ' | '+'Report '+report+' not found or there is a copy already in the send to client folder\n')
    def move_files_parms(self):
        db.log_file.write('========================================================================= Moving Files =========================================================================\n')
        #Add reports and clients here
        reports=['_4A','_4B','_4F','_4H','_5D','_5E','_5J-BOB','_8','_12Q','_13E','_13F','_13G','_14','_18Q','_18R','_18S','_19A','_19G','_19K','_19M','_19V','_19W',
            '_19AL','_19B','_19E','_19F','_19H','_19I','_19T','_19R','_19S','_19Q','_19X','_19Z','_19AC','_19AW','_19AX','_19AY','_19AB','_19AD','_19AF','_19AH','_19AI',
            '_19AJ','_19AU','_19BA','_19BB','_19BC','_19BJ','_19BK','_20D','_20E']
        specrep={'SELECTRONICS':['_11','_12M','_12N','_12O','_12P','_12I','_12J','_12K','_21B'],
            'MIEAC':['_12O','_12P'],'ONVOY':['_5C','_12O','_12P'],'MTAFA':['_12B','_12O','_12P','_20','_20A'],
            'MTARC':['_12B','_12O','_12P'],'MTASW':['_12B','_12O','_12P'],'NEUTRAL_TANDEMSW':['_12B','_12O','_12P'],
            'NEUTRAL_TANDEMFA':['_12B','_12O','_12P'],'GCISW':['_12B','_12O','_12P','_21B'],'GCIFA':['_12B','_12O','_12P','_20','_20A']}
        date=self.cal.get_date()
        date=list(date)
        year=date[0]+date[1]+date[2]+date[3]
        month=date[5]+date[6]
        db.get_all_files(self.company,month, year, self.billtype)
        mps_path = db.retrieve_path('MPS')
        path = os.path.split(mps_path)
        path=path[0]+"\\" #Getting regular folder path
        self.move_files(reports,path)
        if self.company.upper()!='AMERICAN_BROADBAND' and self.company.upper()!='BENDTEL' and self.company.upper()!='PEERLESS_NETWORK':
            self.move_files(specrep[self.company.upper()+self.billtype.upper()],path) 
        print("The Reports have been moved\n")
        db.log_file.write("The Reports have been moved\n")
        db.log_file.flush()
        os.fsync(db.log_file.fileno())


    def get_answer(self):
        if self.bill_type != '' and self.part != '':
            #Disabling buttons after clicking start
            self.start['state']='disabled'
            self.reset['state']='disabled'
            self.rename['state']='disabled'
            self.quit['state']='disabled'
            #Getting the date from the calendar
            date=self.cal.get_date()
            date=list(date)
            year=date[0]+date[1]+date[2]+date[3]
            month=date[5]+date[6]
            reports.temp_month = month
            reports.temp_year = year
            print("You have selected "+self.company+self.billtype+self.part+" for month "+month+" and year "+year+"\n")
            db.log_file.write("You have selected "+self.company+self.billtype+self.part+" for month "+month+" and year "+year+"\n")
            
            
            #getting database file and MPS locations
            db_location = db.db_location
            
            
            #reports.temp_month = input('Please enter the month in 2 digits? for example 05 ')
            #reports.temp_year = input('Please enter the month in 4 digits? for example 2021 ')
            #month,year = reports.temp_month,reports.temp_year
            #cleaning the database before start working
            db.clean_database()
            reports.pdf.company = self.company
            #Getting the path where the report files and MPS are located
            db.renaming_reports(self.company, month, year, self.billtype)
            #Getting the path where the report files and MPS are located
            
            db.get_all_files(self.company,month, year, self.billtype)
            reports.xls.mps_path = db.retrieve_path('MPS') 
            #reports.macro_path = db.get_all_files(self.company,month, year, self.billtype)
            if self.company == 'Peerless_Network':
                reports.macro_path = db.get_all_files(self.company,month, year, self.billtype)                                                                              
                shutil.copy(os.getcwd()+'\\Peerless_Network_Macros.xlsm',reports.macro_path)


            #Calling functions based on the client and section 
            function_dict = {'Selectronics': clients.selectronics, 'BendTel': clients.bendtel, 'MIEAC':clients.mieac, 'ONVOY':clients.onvoy
                    ,  'MTASW':clients.mta_sw , 'MTAFA':clients.mta_fa , 'MTARC':clients.mta_rc, 'GCIFA':clients.gci_fa,"Neutral_TandemFA":clients.nt_fa
                    ,"GCISW":clients.gci_sw,'Neutral_TandemSW':clients.nt_sw,'American_Broadband590GIN':clients.amb_590GIN, 
                    'American_Broadband356DMI':clients.amb_356DMI,'American_Broadband509BOH':clients.amb_509BOH,'Peerless_Network':clients.peerless}
            #Calling functions based on the client and section 
            self.bar()
            #threading.Thread(target=self.update).start()
            function_dict[self.company+self.billtype](self.part)
            #Enabling buttons
            self.start['state']='normal'
            self.reset['state']='normal'
            self.quit['state']='normal'
            self.rename['state']='normal'    
        else:
            print("You have not selected a bill type or group\n")
    #Method to reset the radio buttons and the Report labels 
    def reset(self):
        i=0
        j=0
        k=0
        self.v.set(None)
        self.j.set(None)
        self.p.set(None)
        for radio_button in self.billtypes_rad_butt:
            self.billtypes_rad_butt[i].configure(state = NORMAL)
            i=i+1
        for radio_button in self.group_rad_butt:
            self.group_rad_butt[j].configure(state = NORMAL)
            j=j+1
        if len(self.report_labels)>0:
            for report_label in self.report_labels:
                self.report_labels[k].destroy()
                k=k+1
        print("Resetting Done\n")



#starts here

if __name__ == "__main__":

    # creating lof file in case that it was not exist
    if not (os.path.isfile('log.txt')):
        log_file = open('log.txt', 'w')


    # Creating a tkinter window
    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        root = Tk()
        os.system('cls')
        # Initialize tkinter window with dimensions 300 x 250             
        root.geometry('800x700')
        root.configure(bg='grey31')
        root.title("CABS Analyst Assistant")
        app=mainapp(master=root)
    
    root.mainloop()
