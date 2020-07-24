import io
from io import StringIO
import os
import sys
import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from pathlib import Path
import shutil
import win32api
import win32print
import time
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

import warnings
warnings.filterwarnings("ignore")

loc = open('ProgLocation.txt','r')
original = loc.readline().strip()
loc.close()

raw = []
final = []
titl = []
reload = open('Finish_Database_Raw.txt')
finie = open('Finish_Database_Organized.txt')
count = 0
while True:
    count+=1
    line = reload.readline().strip()
    if not line:
        break
    raw.append(line.split('\t'))
while True:
    count+=1
    line = finie.readline().strip()
    if not line:
        break
    elif line[0] == "[":
        titl = set()
        titl.add(line.lower()[1:-1])
        final.append(titl)

for i in range (len(raw)):
    for k in range(len(final)):
        if raw[i][5].lower() in final[k]:
            final[k].add(raw[i][0])
reload.close()
finie.close()


pull = []
check = open('date&default.txt','r')
count = 0
while True:
    count+=1
    line = check.readline().strip()
    if not line:
        break
    pull.append(line)
check.close()

ticketsD = pull[1]
tagsD = pull[2]
outloc = open('OutputLocation.txt','r')
outputD = outloc.readline().strip()

currentprinter = win32print.GetDefaultPrinter()


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    codec = 'utf-8'
    laparams = LAParams()

    with io.StringIO() as retstr:
        with TextConverter(rsrcmgr, retstr, codec=codec,
                           laparams=laparams) as device:
            with open(path, 'rb') as fp:
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                password = ""
                maxpages = 0
                caching = True
                pagenos = set()

                for page in PDFPage.get_pages(fp,
                                              pagenos,
                                              maxpages=maxpages,
                                              password=password,
                                              caching=caching,
                                              check_extractable=True):
                    interpreter.process_page(page)
                return retstr.getvalue()


def get_full_WIN_num_list(pdftxt):
    final_list = []
    dump = open('currentList.txt','w')
    dump.writelines((pdftxt))
    dump.close()
    dump = open('currentList.txt')
    for line in dump:
        line = line.strip()
        line_edit = line[1:]
        if line_edit.isdigit() == True:
            if len(line) == 19:
                if (len(final_list) == 0) or (line != final_list[-1]):
                    final_list.append(line)
    dump.close()
    return final_list


def extractPages(nameList):
    global original
    all_tickets = PdfFileReader(ticketsD)
    all_tag = PdfFileReader(tagsD)

    for i in range(len(nameList)):
        c = canvas.Canvas("Mergeable.pdf")
        c.drawString(100,740,nameList[i])
        c.showPage()
        c.save()
        watermark = PdfFileReader("Mergeable.pdf")
        watermarkpage = watermark.getPage(0)
        pdf = PdfFileReader("Traveler.pdf")
        pdfwrite = PdfFileWriter()
        pdfpage = pdf.getPage(0)
        pdfpage.mergePage(watermarkpage)
        pdfwrite.addPage(pdfpage)
        with open(nameList[i]+"WM.pdf", 'wb') as fh:
            pdfwrite.write(fh)

    
    pgnum = all_tickets.getNumPages()
    
    
    outloc = open('OutputLocation.txt','r')
    end = outloc.readline().strip()
    outloc.close()
    dat = open('date&default.txt','r')
    prevdate = dat.readline().strip()
    dat.close()

    doubleloc = False
    if not os.path.isdir(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get()):
        dir = os.path.join(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get())
        if not os.path.exists(dir):
            os.mkdir(dir)
    else:
        for i in range(20):
            if not os.path.isdir(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get()+"("+str(i+1)+")"):
                dir = os.path.join(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get()+"("+str(i+1)+")")
                if not os.path.exists(dir):
                    os.mkdir(dir)
                    puthere = i+1
                    doubleloc = True
                    break

    for i in range(pgnum):
        cons = PdfFileReader(nameList[i]+'WM.pdf')
        curr_ticket = all_tickets.getPage(i)
        curr_tag = all_tag.getPage(i)
        constant = cons.getPage(0)
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(curr_ticket)
        pdf_writer.addPage(curr_tag)
        pdf_writer.addPage(constant)
        with Path(str(nameList[i])+".pdf").open(mode="wb") as output_file:
            pdf_writer.write(output_file)

    
    for k in range(len(final)):
        for m in range(len(nameList)):
            if nameList[m][:8] in final[k]:
                win32api.ShellExecute(0, 'print', str(nameList[m])+'.pdf', currentprinter ,'.', 0)
                nameList[m] = nameList[m]+"$"
                print("Printed "+str(nameList[m])+".pdf")
    for m in range(len(nameList)):
        if nameList[m][-1] != "$":
            pass
            win32api.ShellExecute(0, 'print', str(nameList[m])+'.pdf', currentprinter ,'.', 0)
        else:
            nameList[m] = nameList[m][:-1]
    for i in range(len(nameList)):
        os.remove(str(nameList[i])+"WM.pdf")
    isdone = "null"
    while isdone == "null":
        isdone = input("Press Enter if Printing is Done (PDFs are no longer open on PC): ")
    #time.sleep(30)
    while True:
        try:
            for i in range(pgnum):
                if doubleloc == False:
                    try:
                        shutil.move((original +"/"+str(nameList[i])+".pdf"),(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get()))
                    except shutil.Error:
                        pass
                else:
                    try:
                        shutil.move((original +"/"+str(nameList[i])+".pdf"),(end+"/"+month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get()+"("+str(puthere)+")"))
                    except shutil.Error:
                        pass
            break
        except PermissionError:
            print("Waiting for printing")

    
def getTicketsname():
    global ticketsD
    ticketsD = askopenfilename()
    ticketsPath = tk.Label(text=ticketsD)
    ticketsPath.grid(row=5,column=0,pady = 10,columnspan=5)

def getTagsname():
    global tagsD
    tagsD = askopenfilename()
    tagsPath = tk.Label(text=tagsD)
    tagsPath.grid(row=7,column=0,pady = 10,columnspan=5)

def getOutput():
    global outputD
    outputD = askdirectory()
    outputPath = tk.Label(text=outputD)
    outputPath.grid(row=9,column=0,pady = 10,columnspan=5)
    outloc = open('OutputLocation.txt','w')
    outloc.writelines(outputD)
    outloc.close()

def collectAndRun():

    extractPages(get_full_WIN_num_list(convert_pdf_to_txt(ticketsD)))

    push = []
    check = open('date&default.txt','w')
    push.append(month_entry.get()+"_"+day_entry.get()+"_"+year_entry.get())
    push.append("\n")
    push.append(ticketsD)
    push.append("\n")
    push.append(tagsD)
    push.append("\n")
    push.append(str(pull[3]))
    check.writelines((push))
    check.close()

    last = tk.Label(text="The last date used was: "+push[0])
    last.grid(row=11,column=0,columnspan=5, pady = 10)

window = tk.Tk()
window.columnconfigure([0,1,2,3,4], minsize=50, weight = 1)
window.rowconfigure([0,1,2,3,4,5,6,7,8],minsize=5, weight = 1)

date = tk.Label(text="Enter the date:")
date.grid(row=1,pady=5)
mon = tk.Label(text="Month:")
mon.grid(row=2,column = 0, pady=5)
dat = tk.Label(text="Day:")
dat.grid(row=2,column = 2, pady=5)
yr = tk.Label(text="Year:")
yr.grid(row=2,column = 4, pady=5)
month_entry = tk.Entry()
month_entry.grid(row=3,column=0, pady = 10)
first = tk.Label(text="/")
first.grid(row=3,column=1)
day_entry = tk.Entry()
day_entry.grid(row=3, column=2, pady = 10)
second = tk.Label(text="/")
second.grid(row=3,column=3)
year_entry = tk.Entry()
year_entry.grid(row=3, column=4, pady = 10)

getTicketsfile = tk.Button(text='Choose TICKETS File',command=getTicketsname)
getTicketsfile.grid(row=4,column=2)
ticketsPath = tk.Label(text=ticketsD)
ticketsPath.grid(row=5,column=0,columnspan=5, pady = 10)

getTagsfile = tk.Button(text='Choose TAGS File',command=getTagsname)
getTagsfile.grid(row=6,column=2)
tagsPath = tk.Label(text=tagsD)
tagsPath.grid(row=7,column=0,columnspan=5, pady = 10)

getOutputfile = tk.Button(text='Choose Output Directory',command=getOutput)
getOutputfile.grid(row=8,column=2)
outputPath = tk.Label(text=outputD)
outputPath.grid(row=9,column=0,columnspan=5, pady = 10)

run = tk.Button(text='Run',command=collectAndRun)
run.grid(row=10,column=2, pady = 10)


last = tk.Label(text="The last date used was: "+pull[0])
last.grid(row=11,column=0,columnspan=5, pady = 10)
here = open('ProgLocation.txt','r')
amhere = here.readline().strip()
here.close()
where = tk.Label(text="Current Location of Program: "+amhere)
where.grid(row=12,column=0,columnspan=5)


tk.mainloop()
