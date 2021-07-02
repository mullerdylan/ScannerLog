from sys import setswitchinterval
from tkinter import *
import sqlite3
import time
import sched
import datetime
import os
from xlsxwriter.workbook import Workbook

root = Tk()
conn = sqlite3.connect('Scan_Log.db')
root.geometry("960x540")
root.title("USPS Scanner Checkout Application")
root.iconbitmap('C:/Users/dtmuller/Desktop/python projects/Tkinter/ScannerLog/scanicon.ico')

# sets background image template 
bg = PhotoImage(file = "uspslogo.png")
#defines canvas settings/removes border highlight
my_canvas = Canvas(root, width = 960, height =540, bd = 0, highlightthickness=0)
my_canvas.pack(fill = "both", expand = True)

#set image in canvas 
my_canvas.create_image(0,0, image = bg, anchor = "nw")
#Creates static text
my_canvas.create_text(500,60, text= "Scan EIN & SN  ", font = ("helvetica",15) ,fill= "white")
my_canvas.create_text(10,520, text= "Comments or questions: dylan.muller@usps.gov", anchor = "nw", font = ("helvetica",8) ,fill= "white")
vartext  = my_canvas.create_text(20,450, text= "", anchor = "nw", font = ("helvetica",15) ,fill= "white")

#Creates input box
IB = Entry(root, font =("Helvetica", 24), width = 20, fg = "black", bd = 0)
#adds boxes to canvas
IBwindow = my_canvas.create_window(500,120,window= IB)

#####start of SQLInput Code####
IB.focus_set()

###create table if one doesn't exits
conn = sqlite3.connect('Scan_Log.db')
c = conn.cursor()
c.execute(""" CREATE TABLE IF NOT EXISTS Scan (
                EIN text,
                EID text,
                IBT text,
                scanstat text
            ) """)
conn.commit
conn.close

##show  all records query
def query():
    conn = sqlite3.connect('Scan_Log.db')
    c = conn.cursor()
    c.execute("SELECT * FROM Scan")
    recs = c.fetchall()
    top = Toplevel()
    top.title('Checked Out Scanners')
    Label(top, text ="EID                  SN           Time                       Status" , font =("Helvetica", 10), width = 40, fg = "black", bg = "gray", bd = 0).pack()
    for row in recs:
        Label(top,text= row).pack()
    conn.commit()
    conn.close()

##shows currently checkout scanners 
def checkout():
    conn = sqlite3.connect('Scan_Log.db')
    c = conn.cursor()
    c.execute("SELECT * FROM Scan WHERE Scanstat = 'Out'")
    recs = c.fetchall()
  ####launching second window
    top = Toplevel()
    top.title('Checked Out Scanners')
    Label(top, text ="EID                  SN           Time                       Status" , font =("Helvetica", 10), width = 40, fg = "black", bg = "gray", bd = 0).pack()
    for row in recs:
        Label(top,text= row).pack()
    conn.commit()
    conn.close()

IBL = IB.get()
IBLength = len(IB.get())
IBLabel = Label(root, text = "")
ST = " Error"
cond = 0
cond2 = 0



query_btn = Button(root, text = "Show All Records", width=20, command = query)
query_window = my_canvas.create_window(860,520,window= query_btn)
checkout_btn = Button(root, text = "Show Checked Out Scanners", width=25, command = checkout)
query_window = my_canvas.create_window(670,520,window= checkout_btn)

EINL = Label(root, text ="EIN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
EINwindow = my_canvas.create_window(25,351,anchor = "nw",window= EINL)
EIDL = Label(root, text ="SN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
EIDwindow = my_canvas.create_window(25,391,anchor = "nw",window= EIDL)

def submit():
    conn = sqlite3.connect('Scan_Log.db')
    IBL = IB.get()
    IBLength = len(IB.get()) 
    global EINL
    global EIN
    global EIDL
    global EID
    global cond
    global cond2
    
    if IBLength == 12:
        ST = " ScannerID"
        #handles leading 000 in badge reader
        EIN = IBL[4:]
        EINL.config(text ="EIN is "+ EIN , font =("Helvetica", 20), width = 20, fg = "black", bg ="green", bd = 0)
        cond = 1
    elif IBLength == 10:
        ST = " EIN"
        EID = IBL
        EIDL.config(text ="SN is "+ EID , font =("Helvetica", 20), width = 20, fg = "black", bg ="green", bd = 0)
        #used to check in scanner, doesn't require checkin condition to be true  
        conn = sqlite3.connect('Scan_Log.db')
        c = conn.cursor()
        c.execute("UPDATE Scan SET scanstat = 'In' WHERE EID = (?)", [EID])
        conn.commit()
        conn.close()
        cond2 = 1
    else:
        ST = " Error"
        if IBL == "checkin" and cond2 == 1:
            EIDL.config(text ="SN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
            EINL.config(text ="EIN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
            my_canvas.itemconfig(vartext, text= "Scanner Checked In ", anchor = "nw", font = ("helvetica",15) ,fill= "green" )
            cond = 0
            cond2 = 0
        elif IBLength>=1:
            EIDL.config(text ="SN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
            EINL.config(text ="EIN " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)    
            cond = 0
            cond2 = 0
 
    IB.delete(0,END)
    root.after(4000,submit)
    if cond == 1 & cond2 == 1:
        EINL.config(text ="EIN " , font =("Helvetica", 20), width = 10, fg = "black",bg = "gray", bd = 0)
        EIDL.config(text ="SN  " , font =("Helvetica", 20), width = 10, fg = "black", bg = "gray", bd = 0)
        rec = str(" EID " + EID + " EIN " + EIN)
        scanstat = "Out" 
        IBT = datetime.datetime.now()
        conn = sqlite3.connect('Scan_Log.db')
        c = conn.cursor()
        # c.execute("UPDATE Scan SET scanstat = 'In' WHERE EID = (?)", [EID])  ##new code
        c.execute("INSERT INTO Scan VALUES (:EIN, :EID, :IBT, :scanstat) ",
            {
                'EIN': EIN,
                'EID': EID,
                'IBT': IBT,
                'scanstat': scanstat, ##new code
            })
        
        ####writes value to excel sheet
        workbook = Workbook('Output.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'EID')
        worksheet.write('B1', 'SN')
        worksheet.write('C1', 'Time')
        worksheet.write('D1', 'Status')

        mysel=c.execute("SELECT * FROM Scan")
        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i+1, j, row[j])
        workbook.close()
        conn.commit()
        conn.close()

        cond = 0
        cond2 = 0
        my_canvas.itemconfig(vartext, text= "Last Succesful Record " + rec, anchor = "nw", font = ("helvetica",15) ,fill= "green" )
        

root.after(0,submit)
SubmitB = Button(root, text = "submit", command = submit)

root.mainloop()