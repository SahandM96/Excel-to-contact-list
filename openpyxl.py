import openpyxl

from tkinter import filedialog
from Tkinter import *
import ttk
import  tk
import pymsgbox
reload(sys);
sys.setdefaultencoding("utf8")
def gui():
    # Make Message
    pymsgbox.alert('Welcome To My App', 'SahandM96')
    win=tk.Tk()
    win.title("Excel to Contact")
    aLabel =ttk.Label(win,text="Open Excel File :")
    aLabel.grid(column=0,row=0)
    # aLabel =ttk.Label(win,text="Convert CSV to VCF")
    # aLabel.grid(column=0,row=1)
    action=ttk.Button(win,text="Convert",command=openfile)
    action.grid(column=1,row=0)
    # action=ttk.Button(win,text="Convert",command=CSV2VCF)
    # action.grid(column=1,row=1)
    win.mainloop()

def openfile():
    Tk().withdraw()
    exaddress = askopenfilename(filetypes=[
         ("Exel Files", ("*.xls", "*.xlsx")),
         ("XLSX", '*.xlsx'),
         ("XLS", '*.xls'),
         ('All', '*')
    ],)
    exfile=openpyxl.load_workbook(exaddress)
    exfile.template = True

gui()
