# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 15:51:04 2022

@author: Engineering
"""
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox
import tkinter.scrolledtext as scrolledtext
from requests import get
from json import loads 
import pandas as pd
from datetime import datetime, timedelta, date
from datetimerange import DateTimeRange

from traceback import format_exception
from re import split as resplit
from math import sin, cos, radians
from os import path, listdir ,remove
import sys
from glob import glob
import win32com.client
from shutil import rmtree
pathwin32=win32com.__gen_path__
directory_contents = listdir(pathwin32)
for item in directory_contents:    
    if path.isdir(path.join(pathwin32, item)):        
        rmtree(path.join(pathwin32, item))

from PIL import Image, ImageTk




class Demo1:
    
        
    def __init__(self, master):
        
        
        
        master.report_callback_exception = self.report_callback_exception
        
        self.today= date.today()
        self.master = master
        self.ICAOinpasSTR=""
        self.ETAdtinpasSTR=""
        self.evaloutSTR=""
        self.TAFraw=""
        self.imported=""
        
        self.TAFstringvar = tk.StringVar()
        self.EvalOutvar = tk.StringVar()
        if getattr(sys, 'frozen', False):
            self.application_path = path.dirname(sys.executable)
        elif __file__:
            self.application_path = path.dirname(__file__)
        
        
        self.alternatelist=pd.read_excel(path.join(self.application_path,'add_files','alternate','alternate.xlsx'), sheet_name="MINIMUMS", header=1, index_col=0)
        self.alternatelist1=pd.read_excel(path.join(self.application_path,'add_files','alternate','alternate.xlsx'), sheet_name="PREFERRED ALTERNATES", header=11, usecols=["Station","Alt1", "Alt2", "Alt3", "Alt4", "Alt5"])
        
        
        
        self.frame = tk.Frame(self.master, background="#000000", bd=3)
        
       
        self.reseticon = Image.open(path.join(self.application_path,'add_files','icon','reset.png'))
        
        self.reseticon = self.reseticon.resize((10, 10))
        self.reseticon= ImageTk.PhotoImage(self.reseticon, master=self.master)
        
        self.CYicon = Image.open(path.join(self.application_path,'add_files','icon','CYGri.png'))
        self.sizer=200
        self.CYicon = self.CYicon.resize( (self.sizer, (int(self.sizer*.122))),Image.ANTIALIAS)
        self.CYicon= ImageTk.PhotoImage(self.CYicon, master=self.master)
        self.Cyiconlabel = tk.Label(self.frame, image = self.CYicon, background="#000000")
        self.CYversion="V1.01"
        self.Versionlabel = tk.Label(self.frame, text=self.CYversion, font=("Ariel", 10), background="#000000", foreground="#BFD0D7")
        self.Sloganlabel = tk.Label(self.frame, text="Eller bilmez, Canyoldas bilir...", font=("Ariel", 10, "italic"), background="#000000", foreground="#BFD0D7")
       
        self.Sloganlabel.place(x=10, y=768)
        
        #self.resizable(0, 0)
        #self.frame.columnconfigure(3, weight=3)
        self.frame.place(height=800, width=900)
        self.Versionlabel.place(x=10, y=715)
        self.Cyiconlabel.place(x=10, y=740)
        #self.frame.rowconfigure(index, weight)┌
        
        
        self.canvas1=tk.Canvas(self.frame,height= 250, width=220, background= "#21232A")
        self.canvas1.place(x=10, y=10)
        
        
        
        
        
       
        
        
        
        
        
        self.ICAOlbl = tk.Label(self.frame, text = "APRT", wraplength= 150, justify="left", background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
        self.ICAOlbl.grid(column=0, row=1, sticky=tk.W , padx=(20, 5), pady=(25,10))
        
        self.ICAO = tk.Text(self.frame, height = 1, width = 10, background="#40434E",  font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        self.ICAO.grid(column=1, row=1, sticky=tk.W , padx=5,  pady=(25,10))
        self.ICAO.bind('<Return>', self.chooseICAObyEnter)
        
        
        self.RWYlbl = tk.Label(self.frame, text = "RWY", wraplength= 150, justify="left",  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
        self.RWYlbl.grid(column=0, row=2, sticky=tk.W , padx=(20, 5), pady=10)
        
        self.RWY = tk.Text(self.frame,  height = 1, width = 10, background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        self.RWY.grid(column=1, row=2, sticky=tk.W , padx=5, pady=10)
        self.RWY.bind('<Return>', self.chooseRWYbyEnter)
        
        
        self.ETAdtlbl = tk.Label(self.frame, text = "ETA", wraplength= 150, justify="left",  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
        self.ETAdtlbl.grid(column=0, row=3, sticky=tk.W , padx=(20, 5), pady=10)
        
        self.ETAdt = tk.Text(self.frame, height = 1, width = 10, background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        self.ETAdt.grid(column=1, row=3, sticky=tk.W , padx=5, pady=(10,0))
        self.ETAdt.bind('<Return>', self.chooseETA)
        
        
        #self.Datelbl = tk.Label(self.frame, text = "ETA", wraplength= 150, justify="left",  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
        #self.ETAdtlbl.grid(column=0, row=3, sticky=tk.W , padx=(20, 5), pady=10)
        #tk.Text(self.frame, height = 1, width = 10, background="#21232A", font=("Ariel", 10), foreground="#FE8438", borderwidth=0)
        
        self.DateTstyle = ttk.Style(self.frame)
        self.DateTstyle.theme_use("clam")
        
        self.DateTstyle.configure("deneme.TEntry", fieldbackground="#2B2D36", bd="#21232A", highlightthickness=0, borderwidth=0, foreground="#16CA76")
        
        #self.DateTstyle.configure('my.DateEntry', fieldbackground='red')
        self.DateT =DateEntry(self.frame, width=10, font=('Arial',10), date_pattern='y-mm-dd', style= 'deneme.TEntry')
        #self.DateT.insert('insert', datetime.strftime(self.today, '%Y-%m-%d'))
        self.DateT.grid(column=1, row=4, sticky=tk.W , padx=5, pady=(0,10))
        #self.DateT.bind('<Button-1>', self.chooseDate)
        
        
        
        
        self.runbutton1 = tk.Button(self.frame, text = 'CALCULATE', width = 15, height=1, font=("Ariel", 15), command = self.mainEngine, background="#017BC7", foreground="#FFFFFF")
        self.runbutton1.grid(column=0, row=5 , ipadx=3, ipady=3, padx=(20,10), pady=10, columnspan=3)
        
        
        self.resetbutton1 = tk.Button(self.frame, image=self.reseticon, width = 15, command = self.resetinputs, background="#BFD0D7")
        self.resetbutton1.grid(column=2, row=1, sticky=tk.W , padx=10, pady=(25,10))
        
        self.resetbutton2 = tk.Button(self.frame, image=self.reseticon, width = 15, command = self.resetinputs2, background="#BFD0D7")
        self.resetbutton2.grid(column=2, row=2, sticky=tk.W , padx=10, pady=10)
        
        self.resetbutton3 = tk.Button(self.frame, image=self.reseticon, width = 15, command = self.resetinputs3, background="#BFD0D7")
        self.resetbutton3.grid(column=2, row=3, sticky=tk.W , padx=10, pady=10)

    def is_connected(self):
        try:
            
            res = get('https://www.google.com/')
            
            if (res.status_code):
              
                self.Intstate = "Online"
                #self.canvas2=tk.Canvas(self.frame,height= 10, width=10, background= "#16CA76")
                
            else:
                
                self.Intstate = "Offline"
        except:
            self.Intstate = "Offline"
            #self.canvas2=tk.Canvas(self.frame,height= 10, width=10, background= "red")
        
        #self.canvas2.place(x=10, y=740)
        #self.Internetlabel = tk.Label(self.frame, text=self.Intstate, font=("Ariel", 10), background="#000000", foreground="#BFD0D7")
        
        #self.Internetlabel.config(text=self.Intstate)
        #self.Internetlabel.place(x=25, y=735)
        
      
       


    def report_callback_exception(self, exc_type, exc_value, exc_traceback):
        self.message = ''.join(format_exception(exc_type,
                                             exc_value,
                                             exc_traceback))
        messagebox.showerror('Error',self.message)   
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
            mail.Subject = 'Canyoldas Bug' 
            mail.Body = self.CYversion +"\n" + self.ICAOinpasSTR+ "\n"+ self.RWYinpasSTR + "\n"+self.ETAdtinpasSTR + "\n\n\n"+ self.message
            mail.Send()
            print("success")
        except:
            pass
        

    def resetinputs(self):
        self.ICAO.config(state='normal',  background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        self.RWY.config(state='normal',  background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        self.ETAdt.config(state='normal',  background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        try: 
            self.ICAOinpasSTR=""
            self.ICAO.delete(1.0,"end-1c")
            self.RWYinpasSTR=""
            self.RWY.delete(1.0,"end-1c")
            self.ETAdtinpasSTR=""
            self.ETAdt.delete(1.0,"end-1c")
        except:
           
            pass
            
        try:
            self.subframe1.destroy()
            self.subframe2.destroy()
            self.subframe3.destroy()
            self.subframe4.destroy()
            
            
            
        except:
            pass
        
        try:
            self.filesindir = glob('add_files/excel/*')
            for f in self.filesindir:
                remove(f)
        except:
            pass
        
    def resetinputs2(self):    
        
        self.RWY.config(state='normal',  background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        
        try:
            self.RWYinpasSTR=""
            self.RWY.delete(1.0,"end-1c")
            
        except:
            pass
        
        try:
            self.subframe1.destroy()
            self.subframe2.destroy()
            self.subframe3.destroy()
            self.subframe4.destroy()
         
        except:
            pass
        
        try:
            self.filesindir = glob('add_files/excel/*')
            for f in self.filesindir:
                remove(f)
        except:
            pass
        
        
    def resetinputs3(self):    
        
        self.ETAdt.config(state='normal',  background="#40434E", font=("Ariel", 15), foreground="#BFD0D7", insertbackground="#BFD0D7")
        
        try:
            
            self.ETAdtinpasSTR=""
            self.ETAdt.delete(1.0,"end-1c")
            
        except:
            pass
        
        try:
            self.subframe1.destroy()
            self.subframe2.destroy()
            self.subframe3.destroy()
            self.subframe4.destroy()
          
        except:
            pass
        
        try:
            self.filesindir = glob('add_files/excel/*')
            for f in self.filesindir:
                remove(f)
        except:
            pass
        

    
    def chooseICAObyEnter(self, event):
        
        
        #self.ICAOlblstringvar = tk.StringVar()
       
        self.ICAOinp= self.ICAO.get(1.0, "end-1c")
        
        self.ICAOinpasSTR=(self.ICAOinp).strip().upper()
        if len((self.ICAOinp).strip())==4:
            #.ICAOlblstringvar.set("Selected airport is "+ self.ICAOinpasSTR)
            #self.ICAOlblselected=tk.Label(self.frame, textvariable = self.ICAOlblstringvar)
            #self.ICAOlblselected.place(relx=.20, rely=.20)
            self.ICAO.delete(1.0,"end-1c")
            self.ICAO.insert("insert",self.ICAOinpasSTR)
            
            self.ICAO.config(state='disabled',  background="#2B2D36", foreground="#16CA76")
            event.widget.tk_focusNext().focus()
        else:
            messagebox.showerror(title="Error", message="Please check airport code!")
            self.ICAO.delete(1.0,"end-1c")
        
    def chooseRWYbyEnter(self, event):
                   
         #self.ICAOlblstringvar = tk.StringVar()
        
         self.RWYinp= self.RWY.get(1.0, "end-1c")
         
         self.RWYinpasSTR=(self.RWYinp).strip().upper()
         if self.RWYinpasSTR[0:2].isnumeric():
             #self.ICAOlblstringvar.set("Selected airport is "+ self.ICAOinpasSTR)
             #self.ICAOlblselected=tk.Label(self.frame, textvariable = self.ICAOlblstringvar)
             #self.ICAOlblselected.place(relx=.20, rely=.20)
             self.RWY.delete(1.0,"end-1c")
             self.RWY.insert("insert",self.RWYinpasSTR)
             
             self.RWY.config(state='disabled', background="#2B2D36", foreground="#16CA76")
             self.RWYdirection=int(self.RWYinpasSTR[0:2])*10
             event.widget.tk_focusNext().focus()
         else:
             messagebox.showerror(title="Error", message="Please check RWY code!")
             self.RWY.delete(1.0,"end-1c")
    
    def chooseETA(self, event):
        
  
        
        self.ETAdtinp= self.ETAdt.get(1.0, "end-1c")
        #
        self.ETAdtinpasSTR=(self.ETAdtinp).strip()
      
        if self.checkifDT(self.ETAdtinpasSTR):
            #self.ETAdtstringvar.set("Selected ETA is "+ self.ETAdtinpasSTR)
            #self.ETAdtselected=tk.Label(self.frame, textvariable = self.ETAdtstringvar)
            #self.ETAdtselected.place(relx=.60, rely=.20)
            
            self.ETAdt.config(state='disabled', background="#2B2D36", foreground="#16CA76")
            
        else:
            messagebox.showerror(title="Error", message="Please check ETA format!")
            self.ETAdt.delete(1.0,"end-1c")
    
    
    #def chooseDate(self, event):
        
     #   self.cal=DateEntry(self.DateT, width=12, background='darkblue',
       #             foreground='white', borderwidth=2, year=2010)
      #  self.cal.pack(pady = 20)
        
    
    def checkifDT(self, checkdt):
        
        self.checkdt=checkdt
        
        try:
            datetime.strptime(self.checkdt, '%H:%M')
            return True
        except:
            return False
        
    def checkifinterval(self, importedTAFstr):
        try:
            self.interval_list=self.importedTAFstr.split("/")
           
            if (self.interval_list[0].isnumeric()) and (self.interval_list[1].isnumeric()):
                
                return True
            else: 
                return False
        except Exception as E:
            
            return False
        
    def checkifinterval1(self,TAFshort_parsedstr):
        try:
            self.interval_list=self.TAFshort_parsedstr.split("/")
           
            if (self.interval_list[0].isnumeric()) and (self.interval_list[1].isnumeric()):
                
                return True
            else: 
                return False
        except Exception as E:
            
            return False



    def TAFaslist(self, TAFraw ,arrivalhr, arrivalhrm1, arrivalhrp1):
        self.TAFsplitted=str(self.TAFraw).replace("['", "")
        self.TAFsplitted=self.TAFsplitted.replace("']", "")
        self.inperiod_color="#FE8438"
        self.notinperiod_color="#FFFFFF"
        self.changeindiclist=["BECMG", "TL", " FM", "PROB30", "PROB40", "TEMPO"]
        self.TAFsplitted=resplit(r'(BECMG |PROB30 |PROB40 |TEMPO | FM | TL )', self.TAFsplitted)
        self.TAFsplitted=[self.x for self.x in self.TAFsplitted if not self.x==""]
        self.finallist=[]
        self.cntTAFstr=0
        while self.cntTAFstr<len(self.TAFsplitted):
            if not (self.TAFsplitted[self.cntTAFstr]).strip() in self.changeindiclist:
                
                self.finallist.append(self.TAFsplitted[self.cntTAFstr])
                self.cntTAFstr+=1
            else:
                if not (self.TAFsplitted[self.cntTAFstr+1]).strip() in self.changeindiclist:
                    self.finallist.append(self.TAFsplitted[self.cntTAFstr]+ self.TAFsplitted[self.cntTAFstr+1])
                    self.cntTAFstr+=2
                else:
                    self.finallist.append(self.TAFsplitted[self.cntTAFstr]+ self.TAFsplitted[self.cntTAFstr+1]+self.TAFsplitted[self.cntTAFstr+2])
                    self.cntTAFstr+=3
        self.TAFparsedDB=pd.DataFrame([], columns=["TAFline", "Color"]) 
        
        
        self.dtRange1 = DateTimeRange(self.arrivalhrm1, self.arrivalhrp1)
        
        
        for cnt2 in range(len(self.finallist)):
            if not "BECMG" in self.finallist[cnt2]:
                self.dummylist=(self.finallist[cnt2]).split(" ")
                
                for cnt3 in range(len(self.dummylist)):
                    self.importedTAFstr=self.dummylist[cnt3]
                    if self.checkifinterval(self.importedTAFstr):
                        
                        self.interval_splitted=(self.importedTAFstr).split("/")
                        self.day1=int(self.interval_splitted[0][0:2])
                        self.hour1=int(self.interval_splitted[0][2:4])
                        if self.hour1==24:
                            self.hour1=0
                            self.day1=self.day1+1
                        self.day2=int(self.interval_splitted[1][0:2])
                        self.hour2=int(self.interval_splitted[1][2:4])
                        if self.hour2==24:
                            self.hour2=0
                            self.day2=self.day2+1
                        if self.day1<=self.day2:
                            self.month1=self.arrivalhr.month
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month
                            self.year2=self.arrivalhr.year
                        elif self.day1==self.arrivalhr.day:
                            self.month1=self.arrivalhr.month
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month+1
                            self.year2=self.arrivalhr.year
                        else:
                            self.month1=self.arrivalhr.month+1
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month+1
                            self.year2=self.arrivalhr.year
                        
                        self.dtRange2 = DateTimeRange(datetime(self.year1, self.month1, self.day1, self.hour1), datetime(self.year2, self.month2, self.day2, self.hour2))
                        
                        if self.dtRange2.is_intersection(self.dtRange1):
                            
                            self.TAFparsedDB.at[cnt2, "TAFline" ]=self.finallist[cnt2]
                            self.TAFparsedDB.at[cnt2, "Color" ]=self.inperiod_color
                        else:
                            
                            self.TAFparsedDB.at[cnt2, "TAFline" ]=self.finallist[cnt2]
                            self.TAFparsedDB.at[cnt2, "Color" ]=self.notinperiod_color
                    
                        
                       
            else: 
                
                self.dummylist=(self.finallist[cnt2]).split(" ")
                for cnt3 in range(len(self.dummylist)):
                    self.importedTAFstr=self.dummylist[cnt3]
                    if self.checkifinterval(self.importedTAFstr):
                        self.interval_splitted=(self.importedTAFstr).split("/")
                        self.day1=int(self.interval_splitted[0][0:2])
                        self.hour1=int(self.interval_splitted[0][2:4])
                        if self.hour1==24:
                            self.hour1=0
                            self.day1=self.day1+1
                        self.day2=int(self.interval_splitted[1][0:2])
                        self.hour2=int(self.interval_splitted[1][2:4])
                        if self.hour2==24:
                            self.hour2=0
                            self.day2=self.day2+1
                        if self.day1<=self.day2:
                            self.month1=self.arrivalhr.month
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month
                            self.year2=self.arrivalhr.year
                        elif self.day1==self.arrivalhr.day:
                            self.month1=self.arrivalhr.month
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month+1
                            self.year2=self.arrivalhr.year
                        else:
                            self.month1=self.arrivalhr.month+1
                            self.year1=self.arrivalhr.year
                            self.month2=self.arrivalhr.month+1
                            self.year2=self.arrivalhr.year
                        
                        
                        
                        if datetime(self.year1, self.month1, self.day1, self.hour1)<=self.arrivalhrm1:
                           
                            self.TAFparsedDB.at[cnt2, "TAFline" ]=self.finallist[cnt2]
                            self.TAFparsedDB.at[cnt2, "Color" ]=self.inperiod_color
                        else:
                            self.TAFparsedDB.at[cnt2, "TAFline" ]=self.finallist[cnt2]
                            self.TAFparsedDB.at[cnt2, "Color" ]=self.notinperiod_color
                        
                        
        self.TAFparsedDB=self.TAFparsedDB.dropna()
      
                        
                
    def AlternateList(self, ICAOinpasSTR):
        self.dummyalternatelist=self.alternatelist[self.alternatelist["Station"]==self.ICAOinpasSTR]
        self.dummyalternatelist=self.dummyalternatelist.reset_index(drop=True)
        self.dummyalternatelist1=self.alternatelist1[self.alternatelist1["Station"]==self.ICAOinpasSTR]
        self.dummyalternatelist1=self.dummyalternatelist1.reset_index(drop=True)
        
        
        
        
        
    def RunAPI(self, selectedairportforTAF):
        self.hdr = {"X-API-Key": "6805e1ead04f42f38838e3f14b"}
        self.req = get("https://api.checkwx.com/taf/"+ self.selectedairportforTAF +"/decoded", headers=self.hdr)
        
        self.req2 = get("https://api.checkwx.com/taf/"+ self.selectedairportforTAF, headers=self.hdr)
        self.resp2 = loads(self.req2.text)
        if not len(str(self.resp2["data"]))>5:
            messagebox.showinfo(title="Error", message="TAF for "+ self.selectedairportforTAF + " could not be found!")
            

    def DataParseManipulation(self, respinp):
        self.TAFdb=pd.json_normalize(self.respinp, "forecast")
        try:
            self.TAFdb=self.TAFdb.explode("conditions", ignore_index=True)
            self.TAFdb1=self.TAFdb.pop("conditions")
            self.TAFdb = self.TAFdb.join(pd.json_normalize(self.TAFdb1))
        except:
            self.TAFdb["code"]=""
            self.TAFdb["prefix"]=""
            self.TAFdb["text"]=""
            
        try:
            self.TAFdb=self.TAFdb.explode("clouds", ignore_index=True)
            self.TAFdb2=self.TAFdb.pop("clouds")
            self.TAFdb = self.TAFdb.join(pd.json_normalize(self.TAFdb2), rsuffix='_clouds')
        except:
            self.TAFdb["code_clouds"]=""
            
            
            
        if not "code" in self.TAFdb:
            self.TAFdb["code"]="" 
        if not "prefix" in self.TAFdb:
            self.TAFdb["prefix"]=""    
        if not "change.probability" in self.TAFdb:
            self.TAFdb["change.probability"]=""
        if not "code_clouds" in self.TAFdb:
            self.TAFdb["code_clouds"]=""

        if not "change.indicator.code" in self.TAFdb:
            self.TAFdb["change.indicator.code"]=""       
        if not "CBtext" in self.TAFdb:
            self.TAFdb["CBtext"]=""
            
        self.TAFdb["timestamp.from"]=pd.to_datetime(self.TAFdb["timestamp.from"], format="%Y-%m-%dT%H:%MZ")
        self.TAFdb["timestamp.to"]=pd.to_datetime(self.TAFdb["timestamp.to"],format="%Y-%m-%dT%H:%MZ")
        try:
            self.TAFdb["change.time_becoming"]=pd.to_datetime(self.TAFdb["change.time_becoming"], format="%Y-%m-%dT%H:%M:%SZ")
        except KeyError: 
            self.TAFdb["change.time_becoming"]=""
        self.TAFdb=self.TAFdb.fillna("")
        self.TAFdb=self.TAFdb.drop_duplicates()
        self.TAFdb=self.TAFdb.reset_index(drop=True)
        
    
    def dummyCBcreate(self):
        
        self.TAFshort_parsed=(self.req2.text).split(" ")
        self.dummyCB=pd.DataFrame([], columns=["cloud", "ft", "day1","hour1" ,"day2", "hour2","CB"])
        for cnt1 in range(len(self.TAFshort_parsed)):
            if (("BKN" in self.TAFshort_parsed[cnt1]) or ("OVC" in self.TAFshort_parsed[cnt1])) and ("CB" in self.TAFshort_parsed[cnt1]):
               
                self.gotobackcnt=1
                
            
                while True:
                    self.TAFshort_parsedstr=self.TAFshort_parsed[cnt1-self.gotobackcnt]
                    if (cnt1-self.gotobackcnt)<=1:
                        break 
                   
                    if self.checkifinterval1(self.TAFshort_parsedstr)==False:
                        self.gotobackcnt+=1
                        
                    else:
                        break
                self.interval_splitted=self.TAFshort_parsed[cnt1-self.gotobackcnt].split("/")
                self.day1=int(self.interval_splitted[0][0:2])
                self.hour1=int(self.interval_splitted[0][2:4])
                if self.hour1==24:
                    self.hour1=0
                    self.day1=self.day1+1
                self.day2=int(self.interval_splitted[1][0:2])
                self.hour2=int(self.interval_splitted[1][2:4])
                if self.hour2==24:
                    self.hour2=0
                    self.day2=self.day2+1
                
                self.cloud=self.TAFshort_parsed[cnt1][0:3]
                self.ft=int(self.TAFshort_parsed[cnt1][3:6])*100
                self.dummyCB=pd.concat([self.dummyCB, pd.DataFrame({"cloud":[self.cloud], "ft": [self.ft], "day1": [self.day1], "hour1": [self.hour1], "day2": [self.day2], "hour2": [self.hour2], "CB": [self.TAFshort_parsed] })])
                
        self.dummyCB=self.dummyCB.reset_index(drop=True) 
       
    
    
    def AlternateEvaluation(self):
        
        self.dummyTAFdb=self.TAFdb[((self.TAFdb["timestamp.from"]<=self.arrivalhrm1) & (self.TAFdb["timestamp.to"]>=self.arrivalhrm1)) | ((self.TAFdb["timestamp.from"]<=self.arrivalhrp1) & (self.TAFdb["timestamp.to"]>=self.arrivalhrp1))]
        self.dummyTAFdb=self.dummyTAFdb.reset_index(drop=True)
        
        for cnt1 in range(len(self.dummyTAFdb)):
            
            self.creteria1=0
            for cnt2 in range(len(self.dummyCB)):
                
                
                
                if (self.dummyTAFdb.at[cnt1, "code_clouds"]==self.dummyCB.at[cnt2, "cloud"]) and (self.dummyTAFdb.at[cnt1, "base_feet_agl"]==self.dummyCB.at[cnt2, "ft"]) and \
                    ((self.dummyTAFdb.at[cnt1, "timestamp.from"]).day==self.dummyCB.at[cnt2, "day1"]) and ((self.dummyTAFdb.at[cnt1, "timestamp.from"]).hour==self.dummyCB.at[cnt2, "hour1"]):
                    if ((self.dummyTAFdb.at[cnt1, "timestamp.to"]).day==self.dummyCB.at[cnt2, "day2"]) and ((self.dummyTAFdb.at[cnt1, "timestamp.to"]).hour==self.dummyCB.at[cnt2, "hour2"]):
                        self.creteria1+=1
                    elif ((self.dummyTAFdb.at[cnt1, "change.time_becoming"]).day==self.dummyCB.at[cnt2, "day2"]) and ((self.dummyTAFdb.at[cnt1, "change.time_becoming"]).hour==self.dummyCB.at[cnt2, "hour2"]):
                        self.creteria1+=1
                    
                    
            if self.creteria1>=1:
                self.dummyTAFdb.at[cnt1, "CB_creteria"]=1
            else:
                self.dummyTAFdb.at[cnt1, "CB_creteria"]=0
            
            
            if "wind.gust_kts" in self.dummyTAFdb.columns:
                if not self.dummyTAFdb.at[cnt1, "wind.gust_kts"]=="":
                    self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[cnt1, "wind.gust_kts"]
            try:
                if self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[0, "wind.speed_kts"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "wind.speed_kts"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=""
                
            
            try:
                if self.dummyTAFdb.at[cnt1, "ceiling.feet"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "ceiling.feet"]=self.dummyTAFdb.at[0, "ceiling.feet"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "ceiling.feet"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "ceiling.feet"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "ceiling.feet"]=""
            
            try:
                if self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=self.dummyTAFdb.at[0, "visibility.meters_float"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "visibility.meters_float"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=""
            
        
            
            if self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=="":
                self.dummyTAFdb.at[cnt1, "evaluation_output1"]=0
            else:
                
                if self.dummyTAFdb.at[cnt1, "wind.speed_kts"]>20:
                    self.dummyTAFdb.at[cnt1, "evaluation_output1"]=1
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output1"]=0
                
            if self.dummyTAFdb.at[cnt1, "CB_creteria"]==1:
                self.dummyTAFdb.at[cnt1, "evaluation_output2"]=1
            else:
                self.dummyTAFdb.at[cnt1, "evaluation_output2"]=0
            try:
                
                if self.dummyalternatelist.at[0,"Alternate VIS."]>self.dummyTAFdb.at[cnt1, "visibility.meters_float"]:
                    self.dummyTAFdb.at[cnt1, "evaluation_output3"]=1
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
            except:
                self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
            
            if (self.dummyTAFdb.at[cnt1, "code_clouds"]=="BKN") or (self.dummyTAFdb.at[cnt1, "code_clouds"]=="OVC"):
                if self.dummyalternatelist.at[0,"Alternate CEI."]>self.dummyTAFdb.at[cnt1, "ceiling.feet"]:
                    self.dummyTAFdb.at[cnt1, "evaluation_output4"]=1
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0
            else:
                self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0
            
            self.dummyTAFdb.at[cnt1, "evaluation_output"]=self.dummyTAFdb.at[cnt1, "evaluation_output1"]+self.dummyTAFdb.at[cnt1, "evaluation_output2"]+self.dummyTAFdb.at[cnt1, "evaluation_output3"]+self.dummyTAFdb.at[cnt1, "evaluation_output4"]
            
    
    
    def EvaluationEngine(self):
        
        self.condition_rank0=["NSW","-DZ", "DZ", "-RA", "RA", "-SHRA", "SHRA",""]
        self.condition_rank1=["+SHRA", "+RA", "-SN", "-SHSNRA", "TCU", "-SHRASN"]
        self.condition_rank2=["BR","-BR","+BR","-FG","FG","+FG","-FU","FU","+FU", "-VA", "+VA", "VA", "-DU", "+DU", "DU", "SA", "-SA", "+SA", "HZ", "+HZ", "-HZ","-PY", "+PY","PY"]
        
        self.dummyTAFdb=self.TAFdb[((self.TAFdb["timestamp.from"]<=self.arrivalhrm1) & (self.TAFdb["timestamp.to"]>=self.arrivalhrm1)) | ((self.TAFdb["timestamp.from"]<=self.arrivalhrp1) & (self.TAFdb["timestamp.to"]>=self.arrivalhrp1))]
        self.dummyTAFdb=self.dummyTAFdb.reset_index(drop=True)
        
        if len(self.dummyTAFdb)==0:
            
            messagebox.showerror(title="Error", message="No weather forecast found for the input date !")
            return
            
        
        for cnt1 in range(len(self.dummyTAFdb)):
            
            self.dummyTAFdb.at[cnt1, "conditionwprefix"]=self.dummyTAFdb.at[cnt1, "prefix"]+self.dummyTAFdb.at[cnt1, "code"]
            
            self.creteria1=0
            for cnt2 in range(len(self.dummyCB)):
                
                #print((TAFdb.at[cnt1, "change.time_becoming"]).day, dummyCB.at[cnt2, "day2"],(TAFdb.at[cnt1, "change.time_becoming"]).hour, dummyCB.at[cnt2, "hour2"]  )
                
                if (self.dummyTAFdb.at[cnt1, "code_clouds"]==self.dummyCB.at[cnt2, "cloud"]) and (self.dummyTAFdb.at[cnt1, "base_feet_agl"]==self.dummyCB.at[cnt2, "ft"]) and \
                    ((self.dummyTAFdb.at[cnt1, "timestamp.from"]).day==self.dummyCB.at[cnt2, "day1"]) and ((self.dummyTAFdb.at[cnt1, "timestamp.from"]).hour==self.dummyCB.at[cnt2, "hour1"]):
                    if ((self.dummyTAFdb.at[cnt1, "timestamp.to"]).day==self.dummyCB.at[cnt2, "day2"]) and ((self.dummyTAFdb.at[cnt1, "timestamp.to"]).hour==self.dummyCB.at[cnt2, "hour2"]):
                        self.creteria1+=1
                        self.dummyTAFdb.at[cnt1, "CBtext"]=str(self.dummyTAFdb.at[cnt1, "code_clouds"])+str(int(self.dummyTAFdb.at[cnt1, "feet"]/100))+"CB"
                    elif ((self.dummyTAFdb.at[cnt1, "change.time_becoming"]).day==self.dummyCB.at[cnt2, "day2"]) and ((self.dummyTAFdb.at[cnt1, "change.time_becoming"]).hour==self.dummyCB.at[cnt2, "hour2"]):
                        self.creteria1+=1
                        self.dummyTAFdb.at[cnt1, "CBtext"]=str(self.dummyTAFdb.at[cnt1, "code_clouds"])+str(int(self.dummyTAFdb.at[cnt1, "feet"]/100))+"CB"
                    
                    
            if self.creteria1>=1:
                self.dummyTAFdb.at[cnt1, "CB_creteria"]=1
            else:
                self.dummyTAFdb.at[cnt1, "CB_creteria"]=0
            
            if "wind.gust_kts" in self.dummyTAFdb.columns:
                if not self.dummyTAFdb.at[cnt1, "wind.gust_kts"]=="":
                    self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[cnt1, "wind.gust_kts"]
            try:
                if self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[0, "wind.speed_kts"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "wind.speed_kts"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "wind.speed_kts"]=""
               
                
            try:
                if self.dummyTAFdb.at[cnt1, "wind.degrees"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "wind.degrees"]=self.dummyTAFdb.at[0, "wind.degrees"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "wind.degrees"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "wind.degrees"]
                            break
                        else:
                            self.tempocnt+=1
                            
                            
            except KeyError:
                self.dummyTAFdb.at[cnt1, "wind.degrees"]=""   
                
            try:
                self.dummyTAFdb.at[cnt1, "Crosswind"]=abs(sin(radians(self.dummyTAFdb.at[cnt1, "wind.degrees"]-self.RWYdirection)))*self.dummyTAFdb.at[cnt1, "wind.speed_kts"]
            except:
                #print("olmadı yar cw yok")
                self.dummyTAFdb.at[cnt1, "Crosswind"]=""
            
            try:
            
                self.dummyTAFdb.at[cnt1, "Tailwind"]=cos(radians(self.dummyTAFdb.at[cnt1, "wind.degrees"]-self.RWYdirection))*self.dummyTAFdb.at[cnt1, "wind.speed_kts"] 
            except:
                #print("olmadı yar tw yok")
                self.dummyTAFdb.at[cnt1, "Tailwind"]=""
                
            try:
                if self.dummyTAFdb.at[cnt1, "code_clouds"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "code_clouds"]=self.dummyTAFdb.at[0, "code_clouds"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "code_clouds"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "code_clouds"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "code_clouds"]=""
                
            try:
                if self.dummyTAFdb.at[cnt1, "ceiling.feet"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "ceiling.feet"]=self.dummyTAFdb.at[0, "ceiling.feet"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "ceiling.feet"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "ceiling.feet"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "ceiling.feet"]=""
                
            try:
                if self.dummyTAFdb.at[cnt1, "conditionwprefix"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "conditionwprefix"]=self.dummyTAFdb.at[0, "conditionwprefix"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "conditionwprefix"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "conditionwprefix"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "conditionwprefix"]=""
            try:
                if self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=="":
                    self.tempocnt=1
                    while True:
                        if (cnt1-self.tempocnt)<1:
                            self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=self.dummyTAFdb.at[0, "visibility.meters_float"]
                            break                
                        if not self.dummyTAFdb.at[cnt1-self.tempocnt, "change.indicator.code"]=="TEMPO":
                            self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=self.dummyTAFdb.at[cnt1-self.tempocnt, "visibility.meters_float"]
                            break
                        else:
                            self.tempocnt+=1
            except KeyError:
                self.dummyTAFdb.at[cnt1, "visibility.meters_float"]=""
                
                
            if self.dummyTAFdb.at[cnt1, "conditionwprefix"] in self.condition_rank1:
                self.dummyTAFdb.at[cnt1, "conditionasinteger"]=1
            elif self.dummyTAFdb.at[cnt1, "conditionwprefix"] in self.condition_rank0: 
                self.dummyTAFdb.at[cnt1, "conditionasinteger"]=0
            elif not self.dummyTAFdb.at[cnt1, "conditionwprefix"] in self.condition_rank2:
                self.dummyTAFdb.at[cnt1, "conditionasinteger"]=2
            else:    
                self.dummyTAFdb.at[cnt1, "conditionasinteger"]=0
            
             
            if self.dummyTAFdb.at[cnt1, "change.indicator.code"]=="TEMPO":
                if (self.dummyTAFdb.at[cnt1, "change.probability"]=="30") or (self.dummyTAFdb.at[cnt1, "change.probability"]=="40"):
                    if self.dummyTAFdb.at[cnt1, "conditionasinteger"]==0:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=0
                    elif self.dummyTAFdb.at[cnt1, "conditionasinteger"]==1:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=1
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=3
                        
                        
                    if self.dummyTAFdb.at[cnt1, "Crosswind"]>20:
                        self.dummyTAFdb.at[cnt1, "evaluation_output1"]=2
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output1"]=0
                        
                    if self.dummyTAFdb.at[cnt1, "CB_creteria"]==1:
                        self.dummyTAFdb.at[cnt1, "evaluation_output2"]=1
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output2"]=0
                    try:
                            
                        if self.dummyalternatelist.at[0,"Alternate VIS."]>self.dummyTAFdb.at[cnt1, "visibility.meters_float"]:
                            self.dummyTAFdb.at[cnt1, "evaluation_output3"]=3
                        else:
                            self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
                    except:
                        self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
                    
                    if (self.dummyTAFdb.at[cnt1, "code_clouds"]=="BKN") or (self.dummyTAFdb.at[cnt1, "code_clouds"]=="OVC"):
                        if self.dummyalternatelist.at[0,"Alternate CEI."]>self.dummyTAFdb.at[cnt1, "ceiling.feet"]:
                            self.dummyTAFdb.at[cnt1, "evaluation_output4"]=3
                        else:
                            self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0

                        
                
                else:
                    if self.dummyTAFdb.at[cnt1, "conditionasinteger"]==0:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=0
                    elif self.dummyTAFdb.at[cnt1, "conditionasinteger"]==1:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=4
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output0"]=5
                        
                        
                    if self.dummyTAFdb.at[cnt1, "Crosswind"]>20:
                        self.dummyTAFdb.at[cnt1, "evaluation_output1"]=4
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output1"]=0
                        
                    if self.dummyTAFdb.at[cnt1, "CB_creteria"]==1:
                        self.dummyTAFdb.at[cnt1, "evaluation_output2"]=4
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output2"]=0
                    try:
                        
                        if self.dummyalternatelist.at[0,"Alternate VIS."]>self.dummyTAFdb.at[cnt1, "visibility.meters_float"]:
                            self.dummyTAFdb.at[cnt1, "evaluation_output3"]=5
                        else:
                            self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
                    except:
                        self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
                    
                    if (self.dummyTAFdb.at[cnt1, "code_clouds"]=="BKN") or (self.dummyTAFdb.at[cnt1, "code_clouds"]=="OVC"):
                        if self.dummyalternatelist.at[0,"Alternate CEI."]>self.dummyTAFdb.at[cnt1, "ceiling.feet"]:
                            self.dummyTAFdb.at[cnt1, "evaluation_output4"]=5
                        else:
                            self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0                 
                    
                    
                    
                    
            else: 
                if self.dummyTAFdb.at[cnt1, "conditionasinteger"]==0:
                    self.dummyTAFdb.at[cnt1, "evaluation_output0"]=0
                elif self.dummyTAFdb.at[cnt1, "conditionasinteger"]==1:
                    self.dummyTAFdb.at[cnt1, "evaluation_output0"]=4
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output0"]=5
                    
                    
                if self.dummyTAFdb.at[cnt1, "Crosswind"]>20:
                    self.dummyTAFdb.at[cnt1, "evaluation_output1"]=4
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output1"]=0
                    
                if self.dummyTAFdb.at[cnt1, "CB_creteria"]==1:
                    self.dummyTAFdb.at[cnt1, "evaluation_output2"]=4
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output2"]=0
                    
                if self.dummyalternatelist.at[0,"Alternate VIS."]>self.dummyTAFdb.at[cnt1, "visibility.meters_float"]:
                    self.dummyTAFdb.at[cnt1, "evaluation_output3"]=5
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output3"]=0
                
                if (self.dummyTAFdb.at[cnt1, "code_clouds"]=="BKN") or (self.dummyTAFdb.at[cnt1, "code_clouds"]=="OVC"):
                    if self.dummyalternatelist.at[0,"Alternate CEI."]>self.dummyTAFdb.at[cnt1, "ceiling.feet"]:
                        self.dummyTAFdb.at[cnt1, "evaluation_output4"]=5
                    else:
                        self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0
                else:
                    self.dummyTAFdb.at[cnt1, "evaluation_output4"]=0  
                    
                    
        
        
        self.dummyTAFdb["evaluation_output"] = self.dummyTAFdb[["evaluation_output0","evaluation_output1","evaluation_output2","evaluation_output3","evaluation_output4"]].max(axis=1)        
          

    def mainEngine(self):
        self.is_connected()
        
        
        if self.Intstate=="Offline":
            
            messagebox.showerror(title="Error", message="No Internet Connection! Please connect to the internet and try again!")
            return
        
        if self.ICAOinpasSTR=="":
            messagebox.showerror(title="Error", message="Please select the Airport!")
            return
        
        if self.ETAdtinpasSTR=="":
            
            messagebox.showerror(title="Error", message="Please select the ETA!")
            return
                                                      
        
        self.AlternateList(self.ICAOinpasSTR)
        
        
        
        if len(self.dummyalternatelist)==0:
            messagebox.showerror(title="Error", message="No minimums found for this airport!")
            return
            
        self.alternateseries=self.dummyalternatelist1.loc[0,:]
        self.dummyaltsforselection=self.alternateseries.tolist()
        self.subframe1 = tk.Frame(self.frame,  highlightthickness=2, height=650,  width=600, background= "#21232A")
        self.subframe1.grid(column=3, row=0, sticky=tk.W , padx=(30, 5), pady=10, rowspan=700)
        self.ETASTR= datetime.strftime(self.DateT.get_date(), "%Y-%m-%d") + " "+ self.ETAdtinpasSTR
        print(self.ETASTR)
        self.ETAdtselectedDT=datetime.strptime(self.ETASTR, '%Y-%m-%d %H:%M')
        self.arrivalhr=self.ETAdtselectedDT
        self.arrivalhrm1=self.arrivalhr-timedelta(hours=1)
        self.arrivalhrp1=self.arrivalhr+timedelta(hours=1)
        self.destalttextlist=("Destination TAF", "Alternate 1 TAF", "Alternate 2 TAF", "Alternate 3 TAF", "Alternate 4 TAF", "Alternate 5 TAF")
        self.dummyaltsforselection = [x for x in self.dummyaltsforselection if str(x) != 'nan']
     
        self.alternate_evaluation_output=[]
        for cnt5 in range(len(self.dummyaltsforselection)):
            
            self.selectedairportforTAF=(self.dummyaltsforselection[cnt5]).replace("*", "")
            
            
            if cnt5==0:
                self.RunAPI(self.selectedairportforTAF)
                self.TAFraw=self.resp2["data"]
               
                self.TAFaslist(self.TAFraw, self.arrivalhr, self.arrivalhrm1, self.arrivalhrp1)
                
            
              
                self.req.raise_for_status()
                self.resp = loads(self.req.text)
                self.resp["data"]
        
            
                self.respinp=self.resp["data"]
                self.DataParseManipulation(self.respinp)
                
                self.dummyCBcreate()
                
                self.EvaluationEngine()
                for cnt1 in range(len(self.dummyTAFdb)):
                    if self.dummyTAFdb.at[cnt1, "evaluation_output"]==0:
                        self.dummyTAFdb.at[cnt1, "response"]="1 Alternate"
                    elif self.dummyTAFdb.at[cnt1, "evaluation_output"]==1:
                        self.dummyTAFdb.at[cnt1, "response"]="2 Alternate"
                    elif self.dummyTAFdb.at[cnt1, "evaluation_output"]==2:
                        self.dummyTAFdb.at[cnt1, "response"]="2 Alternate + Extra fuel for 10 MIN"
                    elif self.dummyTAFdb.at[cnt1, "evaluation_output"]==3:
                        self.dummyTAFdb.at[cnt1, "response"]="2 Alternate + Extra fuel for 15 MIN"
                    elif self.dummyTAFdb.at[cnt1, "evaluation_output"]==4:
                        self.dummyTAFdb.at[cnt1, "response"]="2 Alternate + Extra fuel for 20 MIN"
                    elif self.dummyTAFdb.at[cnt1, "evaluation_output"]==5:
                        self.dummyTAFdb.at[cnt1, "response"]="2 Alternate + Extra fuel for 30 MIN"
                
                
                
                
                self.evaluation_output=self.dummyTAFdb["evaluation_output"].max()
                
                if self.evaluation_output==0:
                    self.evaloutSTR="1 Alternate"
                elif self.evaluation_output==1:
                    self.evaloutSTR="2 Alternate"
                elif self.evaluation_output==2:
                    self.evaloutSTR="2 Alternate + Extra fuel for 10 MIN"
                elif self.evaluation_output==3:
                    self.evaloutSTR="2 Alternate + Extra fuel for 15 MIN"
                elif self.evaluation_output==4:
                    self.evaloutSTR="2 Alternate + Extra fuel for 20 MIN"
                elif self.evaluation_output==5:
                    self.evaloutSTR="2 Alternate + Extra fuel for 30 MIN"
                
                
                self.showDF=self.dummyTAFdb[self.dummyTAFdb["evaluation_output"]==self.evaluation_output]
                self.showDF=self.showDF.reset_index(drop=True)
            
                    
                self.dummyshowDF=self.showDF[["evaluation_output0","evaluation_output1","evaluation_output2","evaluation_output3","evaluation_output4"]]
                self.ColorlistCondition=self.dummyshowDF.values[-1].tolist()
                
                self.ColorlistCondition=['#FF0000' if x==self.evaluation_output else '#FFFFFF' for x in self.ColorlistCondition]
                if self.evaluation_output==0:
                    
                    self.ColorlistCondition=['#FFFFFF' for y in self.ColorlistCondition]
            
                self.WPtext=self.showDF.at[self.showDF.index[-1],"conditionwprefix"]
                if self.WPtext=="":
                    self.WPtext="NIL"
                self.CBtext=self.showDF.at[self.showDF.index[-1],"CBtext"]
                if self.CBtext=="":
                    self.CBtext="NIL"
                self.CCtext=self.showDF.at[self.showDF.index[-1],"code_clouds"]
                if self.CCtext=="":
                    self.CCtext="NIL"
                
                try:
                    
                    self.MinCeiltext=str(int(self.dummyalternatelist.at[0,"Alternate CEI."]))+ " ft"
                except:
                    self.MinCeiltext="NIL"
                
                try:
                    
                    self.ActCeiltext=str(int(self.showDF.at[self.showDF.index[-1],"ceiling.feet"]))+" ft"
                except:
                    self.ActCeiltext="NIL"
                    
                try:
                    
                    self.MinVistext=str(int(self.dummyalternatelist.at[0,"Alternate VIS."]))+ " m"
                except:
                    self.MinVistext="NIL"
                
                try:
                    
                    self.ActVistext=str(int(self.showDF.at[self.showDF.index[-1],"visibility.meters_float"]))+" m"
                except:
                    self.ActVistext="NIL"
                
                
                
               
                
                for cnt7 in range(len(self.showDF)):
                    if self.showDF.at[cnt7,"Tailwind"]>0:
                        self.showDF.at[cnt7,"RWYwind"]=str(int(self.showDF.at[cnt7,"Tailwind"]))+ " TW"
                    elif self.showDF.at[cnt7,"Tailwind"]<=0:
                        self.showDF.at[cnt7,"RWYwind"]=str(int(abs(self.showDF.at[cnt7,"Tailwind"])))+ " HW"
                    else:
                        self.showDF.at[cnt7,"RWYwind"]="0 HW"
                
                self.Windtext=str(self.showDF.at[self.showDF.index[-1],"RWYwind"]) + " / "+ str(int(self.showDF.at[self.showDF.index[-1],"Crosswind"])) + " CW" 
                if self.Windtext=="":
                    self.Windtext="NIL"
                
                self.subframe2 = tk.Frame(self.frame,  highlightthickness=2, height=300,  width=225, background= "#21232A")
                self.subframe2.grid_propagate(0)
                self.subframe2.place(x=10, y=400)
                
                self.WSlbl = tk.Label(self.subframe2, text = "Worst Scenerio",   background="#21232A" ,font=("Ariel", 15, "underline"), foreground="#BFD0D7" )
                self.WSlbl.grid(column=0, row=0, sticky=tk.N , padx=5, pady=10, columnspan=2)
                
                self.WPlbl = tk.Label(self.subframe2, text = "Weather Phenomena: " , wraplength=75, justify='right',  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                self.WPlbl.grid(column=0, row=1, sticky=tk.E , padx=5, pady=3 )
                   
                self.CBlbl = tk.Label(self.subframe2, text = "CB: ", justify='left',  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                self.CBlbl.grid(column=0, row=2, sticky=tk.E , padx=5, pady=3)
                
                
                self.CClbl = tk.Label(self.subframe2, text = "Cloud Condition: ", justify='left',  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                self.CClbl.grid(column=0, row=3, sticky=tk.E , padx=5, pady=3)
                
                self.Ceilminlbl = tk.Label(self.subframe2, text = "Ceiling Minima: ", justify='left',  background="#21232A" ,font=("Ariel", 8, "underline"), foreground="#BFD0D7" )
                self.Ceilminlbl.grid(column=0, row=4 , padx=5, pady=(2,1))
                
                self.Ceilactlbl = tk.Label(self.subframe2, text = "Ceiling Actual: ", justify='left',  background="#21232A" ,font=("Ariel", 8, "underline"), foreground="#BFD0D7" )
                self.Ceilactlbl.grid(column=1, row=4 , padx=5, pady=(2,1))
                
                self.VISminlbl = tk.Label(self.subframe2, text = "Visibility Minima: ", justify='left',  background="#21232A" ,font=("Ariel", 8, "underline"), foreground="#BFD0D7" )
                self.VISminlbl.grid(column=0, row=6 , padx=5, pady=(2,1))
                
                self.VISactlbl = tk.Label(self.subframe2, text = "Visibility Actual: ", justify='left',  background="#21232A" ,font=("Ariel", 8, "underline"), foreground="#BFD0D7" )
                self.VISactlbl.grid(column=1, row=6,padx=5, pady=(2,1))
                
                self.Windlbl = tk.Label(self.subframe2, text = "Wind: ", justify='left',  background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                self.Windlbl.grid(column=0, row=8, sticky=tk.E , padx=5, pady=3)
                
                
                
                
                self.WPtextlbl = tk.Label(self.subframe2, text = self.WPtext , justify='left',  background="#21232A" ,font=("Ariel", 10, "bold"), foreground=self.ColorlistCondition[0] )
                self.WPtextlbl.grid(column=1, row=1, sticky=tk.W , padx=5, pady=3)
                
                self.CBtextlbl = tk.Label(self.subframe2, text = self.CBtext , justify='left',  background="#21232A" ,font=("Ariel", 10, "bold"), foreground=self.ColorlistCondition[2] )
                self.CBtextlbl.grid(column=1, row=2, sticky=tk.W , padx=5, pady=3)
                
                self.CCtextlbl = tk.Label(self.subframe2, text = self.CCtext , justify='left',  background="#21232A" ,font=("Ariel", 10, "bold"), foreground="#BFD0D7" )
                self.CCtextlbl.grid(column=1, row=3, sticky=tk.W , padx=5, pady=3)
                
                self.MinCeiltextlbl = tk.Label(self.subframe2, text = self.MinCeiltext , justify='left',  background="#21232A" ,font=("Ariel", 8, "bold"), foreground="#BFD0D7" )
                self.MinCeiltextlbl.grid(column=0, row=5, padx=5, pady=(1,2))
                
                self.ActCeiltextlbl = tk.Label(self.subframe2, text = self.ActCeiltext , justify='left',  background="#21232A" ,font=("Ariel", 8, "bold"), foreground=self.ColorlistCondition[4] )
                self.ActCeiltextlbl.grid(column=1, row=5, padx=5, pady=(1,2))
                
                self.MinVistextlbl = tk.Label(self.subframe2, text = self.MinVistext, justify='left',  background="#21232A" ,font=("Ariel", 8, "bold"),foreground="#BFD0D7" )
                self.MinVistextlbl.grid(column=0, row=7, padx=5, pady=(1,2))
                
                self.ActVistextlbl = tk.Label(self.subframe2, text = self.ActVistext , justify='left',  background="#21232A" ,font=("Ariel", 8, "bold"), foreground=self.ColorlistCondition[3] )
                self.ActVistextlbl.grid(column=1, row=7,  padx=5, pady=(1,2))
                
                self.Windtextlbl = tk.Label(self.subframe2, text = self.Windtext , justify='left',  background="#21232A" ,font=("Ariel", 10, "bold"), foreground=self.ColorlistCondition[1] )
                self.Windtextlbl.grid(column=1, row=8, sticky=tk.W , padx=5, pady=3)
                
                
                
                
                self.subframe3 = tk.Frame(self.frame,  highlightthickness=2, height=100,  width=225, background= "#21232A")
                self.subframe3.grid_propagate(0)
                self.subframe3.place(x=10, y=290)
                
                self.subframe3.grid_columnconfigure(0, weight=1)
                self.subframe3.grid_rowconfigure(0, weight=1)
                self.EvalOutvar.set(self.evaloutSTR)
                self.OutputLBL=tk.Label(self.subframe3, textvariable = self.EvalOutvar, justify="center", font=("Ariel" , 15,"bold"), background= "#21232A", foreground="#BFD0D7", wraplength=200)
                self.OutputLBL.grid(column=0, row=0, columnspan=100, rowspan=100)
                              
                self.T1=scrolledtext.ScrolledText(self.subframe1, width=65, height=200, background="#2B2D36", state='normal', highlightthickness=2)
                self.dummytext=str(self.selectedairportforTAF)+"\n"
            
                self.T1.tag_configure( "MainDestLabeltag" , justify="left", background="#21232A" ,font=("Ariel", 20, 'bold'), foreground="#BFD0D7")
                self.T1.insert("end", self.dummytext, "MainDestLabeltag") 
                self.T1.insert("end", "\n") 
                
                for cnt4 in range(len(self.TAFparsedDB)):
                    
                    
                    self.dummytext=self.TAFparsedDB.at[cnt4, "TAFline"]+"\n" 
                    self.tag_name = "color-" + self.TAFparsedDB.at[cnt4, "Color"]
                    self.T1.tag_configure(self.tag_name, foreground=self.TAFparsedDB.at[cnt4, "Color"])
                    self.T1.insert("end", self.dummytext, self.tag_name) 
                #self.destaltlbl = tk.Label(self.subframe1, text =self.destalttextlist[cnt5] , justify="left", background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                
                
                #self.maindestlbl = tk.Label(self.subframe1, text =str(self.selectedairportforTAF) ,  )
                
                #self.destaltlbl.place(x=480, y=25+100*cnt5)
                #self.maindestlbl.place(x=20, y=10)
                
                self.T1.insert("end", "\n") 
                self.T1.insert("end", "\n") 
                
                
                self.subframe4 = tk.Frame(self.frame,  highlightthickness=2, height=100,  width=600, background= "#21232A")
                self.subframe4.grid_propagate(0)
                self.subframe4.place(x=260, y=670)
                
                self.RemarkDF=pd.read_excel(path.join(self.application_path,'add_files','alternate','alternate.xlsx'), sheet_name="PREFERRED ALTERNATES", header=11, usecols=["Station","Remarks"])
                self.RemarkDF=self.RemarkDF[self.RemarkDF["Station"]==self.selectedairportforTAF]
                
                self.RemarkDF=self.RemarkDF.reset_index(drop=True)
                self.Remarktext=str(self.RemarkDF.at[0,"Remarks"])
                self.Remarktextlist=self.Remarktext.split("\n")
                
                #print(self.Remarktextlist)
                
                if len(self.Remarktextlist)==1 and str(self.Remarktextlist[0])=="nan":
                    self.Remarktextlist[0]="NIL"
                
                self.T2=scrolledtext.ScrolledText(self.subframe4, width=65, height=200,  state='normal', highlightthickness=2,background="#2B2D36" )
                for cnt8 in range(len(self.Remarktextlist)):
                    self.Remarktextstr=self.Remarktextlist[cnt8].strip()
                    
                    if cnt8 % 2==0:
                        self.remarkbg="#1B1C22"
                        self.remarkfg="#FFFFFF"
                    else:
                        self.remarkbg="#BFD0D7"
                        self.remarkfg="#000000"
    
                        
                        
    
                    
                    self.T2.tag_configure("remark-"+str(cnt8), background=self.remarkbg,foreground=self.remarkfg)
                    self.T2.insert("end", self.Remarktextstr+"\n", "remark-"+str(cnt8)) 
                self.T2.place(x=10, y=5, relwidth=.95, relheight=0.9)
                self.T2.configure(state='disabled') 
                
                
            else:
                self.RunAPI(self.selectedairportforTAF)
                self.TAFraw=self.resp2["data"]
                print(self.TAFraw)
                if len(self.TAFraw)==0:
                    self.TAFraw="NIL"
                    self.dummytext=str(self.selectedairportforTAF)+"\n\n"
                    self.T1.tag_configure( "DestAltLabeltag"+str(cnt5) , justify="left", background="#21232A" ,font=("Ariel", 10, 'bold'), foreground="#BFD0D7" )
                    self.T1.insert("end", self.dummytext,"DestAltLabeltag"+str(cnt5)) 
                    
         
                    self.T1.tag_configure("NILtag", justify="left", foreground="white")
                    self.T1.insert("end", "No TAF found for this alternate! \n" , "NILtag") 
                    
                   
                   
                    self.T1.insert("end", "\n") 
                    self.T1.insert("end", "\n") 
                else:
                    
                    self.TAFaslist(self.TAFraw, self.arrivalhr, self.arrivalhrm1, self.arrivalhrp1)
                    
                
                  
                    self.req.raise_for_status()
                    self.resp = loads(self.req.text)
                    self.resp["data"]
            
                
                    self.respinp=self.resp["data"]
                    self.DataParseManipulation(self.respinp)
                    
                    self.dummyCBcreate()
                    self.AlternateEvaluation()
                    if len(self.dummyTAFdb)==0:
                        self.dummytext=str(self.selectedairportforTAF)+"\n\n"
                        self.T1.tag_configure( "DestAltLabeltag"+str(cnt5) , justify="left", background="#21232A" ,font=("Ariel", 10, 'bold'), foreground="#BFD0D7" )
                        self.T1.insert("end", self.dummytext,"DestAltLabeltag"+str(cnt5)) 
                        
                        for cnt4 in range(len(self.TAFparsedDB)):
                            
                            
                            self.dummytext=self.TAFparsedDB.at[cnt4, "TAFline"]+"\n" 
                            self.tag_name = "color-" 
                            self.T1.tag_configure(self.tag_name, foreground="white")
                            self.T1.insert("end", self.dummytext, self.tag_name) 
                        #self.destaltlbl = tk.Label(self.subframe1, text =self.destalttextlist[cnt5] , justify="left", background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                        self.T1.insert("end", "\n") 
                        self.T1.insert("end", "\n") 
                    else:
                            
                        self.alternate_evaluation_output=self.dummyTAFdb["evaluation_output"].max()
                        
                        
                        if self.alternate_evaluation_output>0:
                            self.alternate_evaluationcolor="#FF0000"
                        else:
                            self.alternate_evaluationcolor="#16CA76"
                             
                        #self.T1=scrolledtext.ScrolledText(self.subframe1, width=65, height=1, background="#2B2D36",  state='normal',  highlightthickness=1)
                        self.dummyalternatelist=self.alternatelist[self.alternatelist["Station"]==self.selectedairportforTAF]
                        self.dummyalternatelist=self.dummyalternatelist.reset_index(drop=True)
                        self.MinVisforalt=str(int(self.dummyalternatelist.at[0,"Alternate VIS."]))
                        self.MinCeilforalt=str(int(self.dummyalternatelist.at[0,"Alternate CEI."]))
                        
                        self.dummytext=str(self.selectedairportforTAF)+"\n"
                        self.T1.tag_configure( "DestAltLabeltag"+str(cnt5) , justify="left", background="#21232A" ,font=("Ariel", 10, 'bold'), foreground=self.alternate_evaluationcolor )
                        self.T1.insert("end", self.dummytext,"DestAltLabeltag"+str(cnt5)) 
                        
                        self.dummytext="Visibility Minima: "+ self.MinVisforalt + "m / Ceiling Minima: " + self.MinCeilforalt+ "ft\n"
                        self.T1.tag_configure( "VisCei"+str(cnt5) , justify="left", background="#21232A" ,font=("Ariel", 8 ), foreground="#BFD0D7" )
                        self.T1.insert("end", self.dummytext,"VisCei"+str(cnt5)) 
                        
                        #print(self.TAFparsedDB)
                        for cnt4 in range(len(self.TAFparsedDB)):
                            
                            
                            self.dummytext=self.TAFparsedDB.at[cnt4, "TAFline"]+"\n" 
                            self.tag_name = "color-" + self.TAFparsedDB.at[cnt4, "Color"]
                            self.T1.tag_configure(self.tag_name, foreground=self.TAFparsedDB.at[cnt4, "Color"])
                            self.T1.insert("end", self.dummytext, self.tag_name) 
                        #self.destaltlbl = tk.Label(self.subframe1, text =self.destalttextlist[cnt5] , justify="left", background="#21232A" ,font=("Ariel", 10), foreground="#BFD0D7" )
                        self.T1.insert("end", "\n") 
                        self.T1.insert("end", "\n") 
                        #self.destaltlbl.place(x=480, y=25+100*cnt5)
                        #self.maindestlbl.place(x=20, y=25+100*cnt5)
                   
                         
                     
                     
                 
                 
             
            self.TAFdb.to_excel(path.join(self.application_path,'add_files','excel',str(self.selectedairportforTAF)+"_ALL.xlsx"))
            self.dummyTAFdb.to_excel(path.join(self.application_path,'add_files','excel',str(self.selectedairportforTAF)+"_PERIOD.xlsx"))
            
        self.T1.place(x=20, y=20, relwidth=.9, relheight=0.9)
        self.T1.configure(state='disabled')  
        
        
        
   


        
def main(): 
    root = tk.Tk()
    root.geometry("900x800")
    root.title("Can Yoldas")
    if getattr(sys, 'frozen', False):
        application_path1 = path.dirname(sys.executable)
    elif __file__:
        application_path1 = path.dirname(__file__)
    root.iconbitmap(path.join(application_path1,'add_files','icon','icon.ico'))
     
    app = Demo1(root)
    #root.after(1000,app.is_connected)   
    root.mainloop()
    

if __name__ == '__main__':
    main()
   