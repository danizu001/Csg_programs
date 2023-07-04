from tkinter import *
import tkinter
import os
from PIL import ImageTk, Image
import tkinter.font as font
from datetime import date, datetime
from tkcalendar import Calendar
import reports
import documents
import ctypes as ct


def add_csg_logo():
    image1 = Image.open("logo.png")
    image2= image1.resize((220,70), Image.ANTIALIAS)
    test = ImageTk.PhotoImage(image2)
    label1 = tkinter.Label(image=test,borderwidth=0, highlightthickness=0)
    label1.image = test
    label1.place(x=30, y=15)    

def activate_butttom(self,array_number):
    for i in range(14):
        if i in array_number:
            self.billtypes_rad_butt[i-1].configure(state = NORMAL)
        else:
            self.billtypes_rad_butt[i-1].configure(state = DISABLED)

def deactivate_companies(self,number):
    for i in range(10):
        if i==number:
            self.company_rad_butt[i-1].configure(state = NORMAL)
        else:
            self.company_rad_butt[i-1].configure(state = DISABLED)  
            
def makeSomething(value):
    global variable
    variable = value
    root1.quit()
    
def Switch_Facility(text1, text2):
    global root1
    root1 = tkinter.Tk()
    root1.title("Do you want switch or facility")
    root1.geometry('400x100')
    Button_yes = Button(root1, text=text1,command=lambda *args: makeSomething(text1),height=5, width=10).pack(side='left')
    Button_no = Button(root1, text=text2,command=lambda *args: makeSomething(text2),height=5, width=10).pack(side='right')
    root1.mainloop()
    return variable

def OH_MI_IN():
    global root1
    root1 = tkinter.Tk()
    root1.title("Choose the state")
    root1.geometry('400x100')
    Button_yes = Button(root1, text='OH',command=lambda *args: makeSomething('OH'),height=5, width=10).pack(side='left')
    Button_no = Button(root1, text='MI',command=lambda *args: makeSomething('MI'),height=5, width=10).pack(side='right')
    Button_middle = Button(root1, text='IN',command=lambda *args: makeSomething('IN'),height=5, width=10)
    Button_middle.place(x=160, y=7)
    root1.mainloop()
    return variable

def Switch_Facility_Recip():
    global root1
    root1 = tkinter.Tk()
    root1.title("Select the bill type")
    root1.geometry('400x100')
    Button_yes = Button(root1, text='SW',command=lambda *args: makeSomething('SW'),height=5, width=10).pack(side='left')
    Button_no = Button(root1, text='FA',command=lambda *args: makeSomething('FA'),height=5, width=10).pack(side='right')
    Button_middle = Button(root1, text='RC',command=lambda *args: makeSomething('RC'),height=5, width=10)
    Button_middle.place(x=160, y=7)
    root1.mainloop()
    return variable  
        
class mainapp(Frame):
    def __init__(self, master):
        self.tit = PhotoImage(file = "title.png")
        self.title = Label(master, image=self.tit,borderwidth=0, highlightthickness=0)
        self.title.place(x=300, y=7)
        add_csg_logo()
        self.company = Label(root, text="1. Customer",bg='grey2',fg='darkorange2')
        self.company.place(x=30, y=100)
        self.company.config(font=("Segoe UI Light", 24))
        self.bill_type = Label(root, text="2. Utilities",bg='grey2',fg='darkorange2')
        self.bill_type.place(x=30, y=250)
        self.bill_type.config(font=("Segoe UI Light", 24))
        self.credits = Label(root, text="A N&D Creation",bg='grey2',fg="dodgerblue2")
        self.credits.place(x=30, y=560)
        self.credits.config(font=("Segoe UI Light", 14))
        
        #Creating companies radio buttons
        self.v = IntVar()
        self.v.set(None)  # initializing the choice, i.e. Python

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
        self.buttonFont = font.Font(family='Segoe UI Light', size=14,weight='bold',slant='italic')
        for comp, val in companies:
            self.company_rad=Radiobutton(root, 
                        text=comp,
                        foreground='mint cream', 
                        variable=self.v, 
                        command=self.select_company,
                        background='grey2',highlightthickness=0,font=self.buttonFont,
                        value=val, selectcolor='darkgoldenrod',activebackground='grey2',activeforeground='mint cream')
            if len(self.company_rad_butt) <= 2:
                x=30
                y=y+30
                self.company_rad.place(x=x,y=y)
            elif len(self.company_rad_butt) <= 5:
                x=250
                y2=y2+30
                self.company_rad.place(x=x,y=y2)
            else:
                x=510
                y3=y3+30
                self.company_rad.place(x=x,y=y3)
            self.company_rad_butt.append(self.company_rad)
        #Creating companies radio buttons for BILL TYPES
        self.j = IntVar()
        self.j.set(None)  # initializing the choice, i.e. Python

        billtypes = [("Clli mapping", 1),
                ("Verify zip", 2),
                    ("Payments", 3),
                    ("Volume trending", 4),
                    ("Error files", 5),
                    ("Move Secabs", 6),
                    ("Bill Count Trending", 7),
                    ("Clean Up", 8),
                    ("Dashboard",9),
                    ("Audit MPS",10),
                    ("OCC",11),
                    ("Thresholding",12),
                    ("Audit Dashboard",13)
                    ]
        #Creating a radio button for each BILL TYPE   
        y4=270
        y2=270
        self.billtypes_rad_butt=[]
        self.buttonFont = font.Font(family='Segoe UI Light', size=14,weight='bold', slant='italic')
        for bt, val in billtypes:
            y=y+30
            self.billtypes_rad=Radiobutton(root, 
                        text=bt,
                        foreground='mint cream', 
                        variable=self.j, 
                        command=self.select_bill_type,
                        background='grey2',highlightthickness=0,font=self.buttonFont, 
                        value=val,selectcolor='darkgoldenrod',activebackground='grey2',activeforeground='mint cream')
            if len(self.billtypes_rad_butt) <= 6:
                x=30
                y4=y4+30
                self.billtypes_rad.place(x=x,y=y4)
            else:
                x=250
                y2=y2+30
                self.billtypes_rad.place(x=x,y=y2)
            self.billtypes_rad_butt.append(self.billtypes_rad)
        for i in range(len(self.billtypes_rad_butt)):
            self.billtypes_rad_butt[i-1].configure(state = DISABLED)
            
        #Getting todays date to set it in the calendar
        todays_date = date.today()
        year=todays_date.year
        month=todays_date.month
        day=todays_date.day

        # Add Calendar
        self.calFont= font.Font(family='constantia', size=10)
        self.cal = Calendar(root, selectmode = 'day',
               year = year, month = month,
               day = day,date_pattern="y-mm-dd", background="black", disabledbackground="black", bordercolor="black", 
               headersbackground="turquoise4", normalbackground="dodgerblue4", foreground='white', 
               normalforeground='white', headersforeground='white', font=self.calFont)
        self.cal.place(x=450, y=280)
        #Array to store the report labels
        self.report_labels=[]
        
        self.photo = PhotoImage(file = "play.png")
        self.docu = PhotoImage(file = "docs.png")
        self.rst = PhotoImage(file = "reset.png")

        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.start =Button(root,image=self.photo,borderwidth=0,highlightthickness=0,command = self.get_answer)
        self.buttonFont = font.Font(family='Segoe UI Light', size=16, weight='bold')
        self.document =Button(root,image=self.docu,borderwidth=0,highlightthickness=0,command = self.docs)
        self.reset =Button(root,image=self.rst,borderwidth=0,highlightthickness=0,command = self.reset)
        #self.document =Button(root,text='Documentation',height = 1,relief=FLAT,width = 15,font=self.buttonFont,fg='white', bg='DarkOrange1',bd=0, highlightthickness =0,command = self.docs)
        # Set a relative position of button
        self.start.place(x=460, y=520)
        self.document.place(x=550, y=520)
        self.reset.place(x=630, y=520)
    
    def get_answer(self):
        if self.bill_type != '':
            #Disabling buttons after clicking start
            self.start['state']='disabled'
            self.document['state']='disabled'
            #Getting the date from the calendar
            date=self.cal.get_date()
            date=list(date)
            year=date[0]+date[1]+date[2]+date[3]
            month=date[5]+date[6]
            reports.temp_month = month
            reports.temp_year = year
            print("You have selected "+self.company+' '+self.billtype+" for month "+month+" and year "+year+"\n")
            reports.run_reports(self.company,self.billtype,month,year)
            #Enabling buttons
            self.start['state']='normal'
            self.document['state']='normal'    
        else:
            print("You have not selected a bill type or group\n")
            
    def reset(self):
        i=0
        j=0
        k=0
        self.v.set(None)
        self.j.set(None)
        for radio_button in self.billtypes_rad_butt:
            self.billtypes_rad_butt[i].configure(state = NORMAL)
            i=i+1
        for radio_button in self.company_rad_butt:
            self.company_rad_butt[j].configure(state = NORMAL)
            j=j+1
        if len(self.report_labels)>0:
            for report_label in self.report_labels:
                self.report_labels[k].destroy()
                k=k+1
        print("Resetting Done\n")
        
    def docs(self):
        if self.bill_type != '':
            self.start['state']='disabled'
            self.document['state']='disabled'
            documents.open_doc(self.billtype)
            #Enabling buttons
            self.start['state']='normal'
            self.document['state']='normal'
            
    def select_company(self): #Disables bill type depending on the company
        if self.v.get() == 1:
            self.company="Selectronics"
            activate_butttom(self,[])
            deactivate_companies(self, 1)
        elif self.v.get() == 2:
            self.company="BendTel"
            activate_butttom(self,[3,4,7,8,9,10,13])
            deactivate_companies(self, 2)
        elif self.v.get() == 3:
            self.company="MIEAC"
            activate_butttom(self,[3,4,6,8,9,10,13])
            deactivate_companies(self, 3)
        elif self.v.get() == 4:
            self.company="ONVOY"
            activate_butttom(self,[2,3,4,8,9,10,13])
            deactivate_companies(self, 4)
        elif self.v.get() == 5:
            self.company="MTA"
            activate_butttom(self,[3,4,7,8,9,10,13])
            deactivate_companies(self, 5)
        elif self.v.get() == 6:
            self.company="American_Broadband"
            activate_butttom(self,[3,4,7,8,9,10,13])
            deactivate_companies(self, 6)
        elif self.v.get() == 7:
            self.company="Peerless_Network"
            activate_butttom(self,[1,2,4,5,6,8,9,10,11,12,13])
            deactivate_companies(self, 7)
        elif self.v.get() == 8:
            self.company="Neutral_Tandem"
            activate_butttom(self,[2,3,4,6,8,9,10,13])
            deactivate_companies(self, 8)
        elif self.v.get() == 9:
            self.company="GCI"
            activate_butttom(self,[3,4,6,7,8,9,10,13])
            deactivate_companies(self, 9)
        return(self.company)

    def select_bill_type(self): #Depending on the app disables the group depending on the bill type
        if self.j.get() == 1:
            self.billtype="Clli mapping"
        elif self.j.get() == 2:
            self.billtype="Verify zip"
        elif self.j.get() == 3:
            self.billtype="Payments"
        elif self.j.get() == 4:
            self.billtype="Volume trending"
        elif self.j.get() == 5:
            self.billtype="Error files"
        elif self.j.get() == 6:
            self.billtype="Move Secabs"
        elif self.j.get() == 7:
            self.billtype="Bill Count Trending"
        elif self.j.get() == 8:
            self.billtype="Clean Up"    
        elif self.j.get() == 9:
            self.billtype="Dashboard" 
        elif self.j.get() == 10:
            self.billtype="Audit MPS" 
        elif self.j.get() == 11:
            self.billtype="OCC" 
        elif self.j.get() == 12:
            self.billtype="Thresholding"
        elif self.j.get() == 13:
            self.billtype="Audit Dashboard" 



        return(self.billtype)  

if __name__ == '__main__':
    root = tkinter.Tk()
    os.system('cls')            
    root.geometry('750x600')
    root.configure(bg='grey2')
    root.resizable(False, False)
    root.title("CABS ONE")
    root.iconbitmap('icon.ico')
    app=mainapp(master=root)
    root.mainloop()