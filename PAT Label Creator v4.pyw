import labels
import string
from string import *
from reportlab.graphics import shapes
import time
import datetime
import calendar
import Tkinter
from Tkinter import Frame,Label,Entry,Button,W,Pack,X,END,LEFT,RIGHT,StringVar,OptionMenu
from tkMessageBox import showerror
import tempfile
import Tkinter as tk
import os
from tkFileDialog import *
from reportlab.lib import colors
from reportlab import rl_settings




top=Tkinter.Tk()

ICON = (b'\x00\x00\x01\x00\x01\x00\x10\x10\x00\x00\x01\x00\x08\x00h\x05\x00\x00'
        b'\x16\x00\x00\x00(\x00\x00\x00\x10\x00\x00\x00 \x00\x00\x00\x01\x00'
        b'\x08\x00\x00\x00\x00\x00@\x05\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'
        b'\x00\x01\x00\x00\x00\x01') + b'\x00'*1282 + b'\xff'*64

_, ICON_PATH = tempfile.mkstemp()
with open(ICON_PATH, 'wb') as icon_file:
    icon_file.write(ICON)


def checkDigits(num):
    wrong=False
    if num.isdigit():
        wrong=True
    return wrong

def checkCorrect(startNo,endNo,month,companyName,custName):
    correct=False

    if companyName=="":
        showerror("Missing Information", "Company Name is empty")
        return correct
##    elif companyTel=="" or checkDigits(companyTel)==False:
##        showerror("Missing Information", "Telephone Number is empty or not a number")
##        return correct
    elif custName=="":
        showerror("Missing Information", "Customer name is empty")
        return correct
    elif startNo=="" or checkDigits(startNo)==False:
        showerror("Missing Information", "Start Number is empty or not a number")
        return correct
    elif endNo=="" or checkDigits(endNo)==False:
        showerror("Missing Information", "End Number is empty or not a number")
        return correct
    elif month=="Select Month":
        showerror("Missing Information", "Month hasn't been selected")
        return correct
    elif int(startNo)>int(endNo):
         showerror("Incorrect Information", "Start Number is less than End Number")
         return correct
    else:
        correct=True
        return correct
    
def companyCorrect(companyName):
    correct=False

    if companyName=="":
        showerror("Missing Information", "Company Name is empty")
        return correct
##    elif companyTel=="" or checkDigits(companyTel)==False:
##        showerror("Missing Information", "Telephone Number is empty or not a number")
##        return correct
    else:
        correct=True
        return correct

def createPassSheet():

    companyName= e1.get()
    companyName=companyName.title()
    custName=e3.get()
    custName=custName.title()
    month=variable.get()
    startNo=e4.get()
    endNo=e5.get()

    companyTel=e6.get()

    

    if checkCorrect(startNo,endNo,month,companyName,custName):
        def draw_label(label, width, height, obj):
            label.add(shapes.String(width/2, 60, "Portable Appliance Test", textAnchor="middle",fontName="Helvetica-Bold", fontSize=8,fillColor="green"))
            label.add(shapes.String(width/2,54, "by "+companyName+"  "+companyTel, textAnchor="middle",fontName="Helvetica-Oblique", fontSize=5,fillColor="green"))
            label.add(shapes.String(0, 53, "___________________________________________", fontName="Helvetica", fontSize=5,fillColor="green"))
            label.add(shapes.String(4, 40, "I.D.: "+str(obj), fontName="Helvetica", fontSize=11,fillColor="green"))
            label.add(shapes.String(4, 29, "Tested: "+month+" "+str(year), fontName="Helvetica", fontSize=9,fillColor="green"))
            label.add(shapes.String(0, 26, "___________________________________________", fontName="Helvetica", fontSize=5,fillColor="green"))
            label.add(shapes.String(width/2, 14, "DO NOT REMOVE ", textAnchor="middle",fontName="Helvetica-Bold", fontSize=11,fillColor="green"))
            label.add(shapes.String(width/2, 3, "Passed", textAnchor="middle",fontName="Helvetica", fontSize=12,fillColor="green"))

        specs = labels.Specification(297, 210, 7, 7, 40, 26, corner_radius=2,left_margin=0,bottom_margin=6,column_gap=2,row_gap=2)
        
            
        sheet = labels.Sheet(specs, draw_label, border=False)

        
        i=int(startNo)
        while i<=int(endNo):
            sheet.add_label(i)
            i=i+1

        fileName=companyName+" - "+custName+" "+month+" "+str(year)+" "+str(startNo)+"-"+str(endNo)+" "+"Passed Labels"
       
    # Save the file and we are done.
        sheet.save(asksaveasfilename(title="File saves as PDF",defaultextension=".pdf",initialfile=fileName))

       
        

def createFailSheet():

    companyName=e1.get()
    companyName=companyName.title()
    companyTel=e6.get()
        
    if companyCorrect(companyName):
        def draw_label(label, width, height, obj):
            
            colour=colors.HexColor(0xFF0000)
            r = shapes.Rect(0, 0, width, height)
            r.fillColor = colour
            r.strokeColor = None
            label.add(r)
            label.add(shapes.String(width/2, 60, "Portable Appliance Test", textAnchor="middle",fontName="Helvetica-Bold", fontSize=8))
            label.add(shapes.String(width/2,54, "by "+companyName+"  "+companyTel,textAnchor="middle",fontName="Helvetica-Oblique", fontSize=5))
            label.add(shapes.String(0, 53, "___________________________________________", fontName="Helvetica", fontSize=5))
            label.add(shapes.String(4, 40, "I.D.: ", fontName="Helvetica", fontSize=12))
            #label.add(shapes.String(4, 29, "Tested: "+month+" "+str(year), fontName="Helvetica", fontSize=10))
            label.add(shapes.String(4, 29, "Tested: ",fontName="Helvetica", fontSize=10))
            label.add(shapes.String(0, 26, "___________________________________________", fontName="Helvetica", fontSize=5))
            label.add(shapes.String(width/2, 14, "DO NOT REMOVE ", textAnchor="middle",fontName="Helvetica-Bold", fontSize=11))
            label.add(shapes.String(width/2, 3, "FAILED", textAnchor="middle",fontName="Helvetica", fontSize=12))

        specs = labels.Specification(297, 210, 7, 7, 40, 26, corner_radius=2,left_margin=0,bottom_margin=6,column_gap=2,row_gap=2)
        
        sheet = labels.Sheet(specs, draw_label, border=False)

       

        sheet.add_labels(range(0, 49))
        fileName=companyName+" - PAT Test Failed Labels"
   
# Save the file and we are done.
        sheet.save(asksaveasfilename(title="File saves as PDF",defaultextension=".pdf",initialfile=fileName))

def clear():

    e1.delete(0, END)
    e3.delete(0, END)
    e4.delete(0, END)
    e5.delete(0, END)
    e6.delete(0, END)
    variable.set("Select Month")

year = datetime.date.today().year

top.title("Light & Easy PAT Label printer")
top.iconbitmap(default=ICON_PATH)

Label(top, text="Testing Company Name").grid(row=0)
Label(top, text="Testing Company Phone").grid(row=1)
Label(top, text="Customer Name").grid(row=2)
Label(top, text="Month").grid(row=5)
Label(top, text="Start Number").grid(row=3)
Label(top, text="End Number").grid(row=4)

e1 = Entry(top)
e6 = Entry(top)
e3 = Entry(top)
e4 = Entry(top)
e5 = Entry(top)
variable = StringVar(top)
variable.set("Select Month")
e2 = OptionMenu(top,variable,"January","February","March","April","May","June","July","August","September","October","November","December")




e1.grid(row=0, column=1,columnspan=2)
e2.grid(row=5, column=1,columnspan=2)
e3.grid(row=2, column=1,columnspan=2)
e4.grid(row=3, column=1,columnspan=2)
e5.grid(row=4, column=1,columnspan=2)
e6.grid(row=1, column=1,columnspan=2)


Button(top, text='Create Passed Sheet', command=createPassSheet,width=15,background="green").grid(row=2, column=4)
Button(top, text='Create Failed Sheet', command=createFailSheet,width=15,background="red").grid(row=3, column=4)
Button(top, text='Clear', command=clear,width=15).grid(row=4, column=4)
Button(top, text='Finished', command=top.quit,width=15).grid(row=5, column=4)

top.mainloop()




