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


def AzData_pull():
    root = tkinter.Tk()
    root.withdraw()
    path = filedialog.askdirectory(parent=root, initialdir="./", title="폴더를 선택 해 주세요")
    print("path : ", path)

    list_filepath_UI = glob.glob(path + '\*.xlsx', recursive=True)
 
    return [list_filepath_UI]
 

# 엑셀 총합 파일 생성
def disk_Az():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 

    #list_filepath = glob.glob(r'C:\Users\pp\Desktop\Exel_aZ\files\*.xlsx', recursive=True)
    list_filepath = AzData_pull()
  
    for filepath in list_filepath:

        print(filepath)
        #wb = excel.Workbooks.Open(filepath)
        #wb.Worksheets("DISK_SUMM").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    #wb_new.SaveAs(r"C:\Users\pp\Desktop\Exel_aZ\test\DISK_SUM.xlsx").format(os.getlogin())

    excel.Quit()

#AzData_pull()


def cpu_Az():
    path = "C:/Users/{}/desktop".format(os.getlogin())  # {}부분에 사용자 이름
    excel = win32com.client.Dispatch("Excel.Application")

    list_filepath = glob.glob(r'C:\Users\{}\Desktop\Exel_aZ\files\*.xlsx'.format(os.getlogin()), recursive=True)

    for filepath in list_filepath:

        print(filepath)
        

    excel.Quit()

def test():
    path = "C:/Users/{}/desktop".format(os.getlogin())  # {}부분에 사용자 이름
 
    print(path) # C:/Users/lucidyul/desktop

#disk_Az()
cpu_Az()
test()

