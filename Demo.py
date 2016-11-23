from Tkinter import *
from tkFileDialog import askopenfilename
import Tkinter as tk
import ttk
import pymsgbox
import unicodecsv as csv
import xlrd
import sys
import os
reload(sys);
sys.setdefaultencoding("utf8")
def gui():
    # Make Message
    pymsgbox.alert('Welcome To My App', 'SahandM96')
    win=tk.Tk()
    win.title("Excel to Contact")
    aLabel =ttk.Label(win,text="Convert Exel to CSV")
    aLabel.grid(column=0,row=0)
    aLabel =ttk.Label(win,text="Convert CSV to VCF")
    aLabel.grid(column=0,row=1)
    action=ttk.Button(win,text="Convert",command=Excel2CSV)
    action.grid(column=1,row=0)
    action=ttk.Button(win,text="Convert",command=CSV2VCF)
    action.grid(column=1,row=1)
    win.mainloop()

# Convert Exel to CSV
def Excel2CSV():
     pymsgbox.alert('For Best Result File Most be Have 4 Column like Name,Family,Org,Note,Phone', 'Warning')
     # Get Name And Address of File
     Tk().withdraw()
     excel_file = askopenfilename(filetypes=[
         ("Exel Files", ("*.xls", "*.xlsx")),
         ("XLSX", '*.xlsx'),
         ("XLS", '*.xls'),
         ('All', '*')
     ],)

     sheet_name = "Worksheet"
     c_s_v_file = pymsgbox.prompt('CSV Name:') + '.csv'
     workbook = xlrd.open_workbook(excel_file)
     worksheet = workbook.sheet_by_name(sheet_name)
     csv_file = open(c_s_v_file, 'wb')
     wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

     for rownum in xrange(worksheet.nrows):
         wr.writerow(
             list(x.encode('utf-8') if type(x) == type(u'') else x
                  for x in worksheet.row_values(rownum)))

     csv_file.close()
    #  pymsgbox.alert('It\s Done', 'SahandM96')/

def CSV2VCF():
    Tk().withdraw()
    somefile=askopenfilename(filetypes=[
        ("CSV Files", ("*.csv")),
        ("CSV", '*.csv'),
        ('All', '*')
    ],)
    Add2Name=pymsgbox.prompt('Anything add to Name :')
    VCF_File= pymsgbox.prompt('VCF Name:') + '.vcf'
    with open( somefile, 'r' ) as source:
        reader = csv.reader( source )
        allvcf = open(VCF_File, 'w')
        i = 0
        for row in reader:
            allvcf.write('BEGIN:VCARD' + "\n")
            allvcf.write('VERSION:3.0' + "\n")
            allvcf.write('N:' + Add2Name +';'+ row[0] + ';' + row[1] + "\n")
            allvcf.write('FN:' + row[1] + ' ' + row[0] + "\n")
            allvcf.write('ORG:' + row[2] + "\n")
            allvcf.write('TEL;CELL:' + row[4] + "\n")
            allvcf.write('NOTE:' + row[3] + "\n")
            allvcf.write('END:VCARD' + "\n")
            allvcf.write("\n")

            i += 1#counts

        allvcf.close()
        # print (str(i) + " vcf cards generated")
        # pymsgbox.alert('It\s Done', 'SahandM96')

# Excel2CSV()
gui()
