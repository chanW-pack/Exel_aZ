import glob
import win32com.client
import os
import datetime
from time import sleep
import openpyxl as op 
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, GradientFill, Alignment, Border, Side
import tkinter
from tkinter import filedialog
from tkinter import *


def AzData_pull():
    root = tkinter.Tk()
    root.withdraw()
    path = filedialog.askdirectory(parent=root, initialdir="./", title="폴더를 선택 해 주세요")                        
    print("path : ", path)

    list_filepath_UI = glob.glob(path + '/*.xlsx', recursive=True)

    return list_filepath_UI

def test1():
    root = tkinter.Tk()
    root.withdraw()
    root.dirName=filedialog.askdirectory()
    print (root.dirName);

    #root.mainloop()
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def cpu_Az():
    gogo = AzData_pull()
    for i in gogo:
        print(i)
    
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 
    #list_filepath = glob.glob(r'C:\Users\*\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    for filepath in gogo:

       wb = excel.Workbooks.Open(filepath)
       wb.Worksheets("CPU_ALL").Copy(Before=wb_new.Worksheets("Sheet1"))

    path = os.getcwd()
    print(path)   

    wb_new.SaveAs("{}/gogo/CPU_SUM.xlsx".format(os.getcwd()))

    excel.Quit()

#AzData_pull()
#test1()
createFolder('./gogo')
cpu_Az()


def test():
    path = "C:/Users/{}/desktop".format(os.getlogin())  # {}부분에 사용자 이름
 
    print(path) # C:/Users/lucidyul/desktop

#disk_Az()

#test()
