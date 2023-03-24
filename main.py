# step1.관련 모듈 및 패키지 import
import glob
import win32com.client
import os
import openpyxl as op #openpyxl 모듈 import
from openpyxl.styles import Font, PatternFill, GradientFill, Alignment, Border, Side

 # 엑셀 총합 파일을 담을 디렉터리 생성
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)
 
createFolder('./test')


# 엑셀 총합 파일 생성
def disk_Az():
    # step2.win32com(pywin32)를 이용해서 엑셀 어플리케이션 열기
    excel = win32com.client.Dispatch("Excel.Application")
    #excel.Visible = True #실제 작동하는 것을 보고 싶을 때 사용

    # step3.엑셀 어플리케이션에 새로운 Workbook 추가
    wb_new = excel.Workbooks.Add() 

    # step4.glob 모듈로 원하는 폴더 내의 모든 xlsx 파일의 경로를 리스트로 반환
    list_filepath = glob.glob(r'C:\Users\홍익과학기술\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    # step5.엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
    for filepath in list_filepath:

        # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
        wb = excel.Workbooks.Open(filepath)

        # 새로 만든 엑셀 파일에 추가
        # 추출할wb.Worksheets("추출할 시트명").Copy(Before=붙여넣을 wb.Worksheets("기준 시트명")
        wb.Worksheets("DISK_SUMM").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    # step6. 취합한 엑셀 파일을 "통합 문서"라는 이름으로 저장
    wb_new.SaveAs(r"C:\Users\홍익과학기술\Desktop\Exel_aZ\test\DISK_SUM.xlsx")

    # step7. 켜져있는 엑셀 및 어플리케이션 모두 종료
    excel.Quit()


def cpu_Az():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 
    list_filepath = glob.glob(r'C:\Users\홍익과학기술\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    for filepath in list_filepath:

        wb = excel.Workbooks.Open(filepath)
        wb.Worksheets("CPU_ALL").Copy(Before=wb_new.Worksheets("Sheet1"))

    wb_new.SaveAs(r"C:\Users\홍익과학기술\Desktop\Exel_aZ\test\CPU_SUM.xlsx")

    excel.Quit()


def mem_Az():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 
    list_filepath = glob.glob(r'C:\Users\홍익과학기술\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    for filepath in list_filepath:
        
        wb = excel.Workbooks.Open(filepath)
        wb.Worksheets("MEM").Copy(Before=wb_new.Worksheets("Sheet1"))

    wb_new.SaveAs(r"C:\Users\홍익과학기술\Desktop\Exel_aZ\test\MEM_SUM.xlsx")

    excel.Quit()



def disk_fx():
    wb = op.load_workbook(r"C:\Users\홍익과학기술\Desktop\Exel_aZ\test\DISK_SUM.xlsx") #Workbook 객체 생성
    ws = wb["Sheet1"] 

    # cell 병합
    ws.merge_cells("A1:D2")
    ws.merge_cells("E1:H2")
    ws.merge_cells("A3:D4")
    ws.merge_cells("E3:H4")
    ws.merge_cells("E10:H13")

    # font 지정
    Az_val_1 = Font(name='맑은 고딕', bold=True, size=24)
    Az_val_2 = Font(name='맑은 고딕', size=24)
    add_inq = Font(name='맑은 고딕', bold=True, size=11)

    ws['A1'].font = Az_val_1
    ws['E1'].font = Az_val_1
    ws['A3'].font = Az_val_2
    ws['E3'].font = Az_val_2
    ws['E10'].font = add_inq

    # cell 테두리 지정
    cell_box = Border(
        left=Side(border_style="medium", color='00000000'),
        right=Side(border_style='medium', color='00000000'),
        bottom=Side(border_style='medium', color='00000000'),
        top=Side(border_style='medium', color='00000000')
    )

    ws['E10'].border = cell_box
                   
    # cell 배경색 추가
    ws['A1'].fill = PatternFill(start_color='C6E0B4', fill_type = 'solid') 
    ws['E1'].fill = PatternFill(start_color='B4C6E7', fill_type = 'solid')

    # cell 정렬 적용 // wrap_text=자동 줄바꿈
    ws['E10'].alignment = Alignment(horizontal="left", vertical="top", wrap_text = True)

    # cell 값 적용
    ws["A1"].value = "Disk Read KB/s"
    ws["E1"].value = "Disk Write KB/s"
    ws["A3"].value = "=AVERAGE('DISK_SUMM:DISK_SUMM (4)'!B59)"
    ws["E3"].value = "=AVERAGE('DISK_SUMM:DISK_SUMM (4)'!C59)"
    ws["E10"].value = "문의사항                                                     TS팀 박찬우                                              tel: 010-9085-0857             chanwoo9730@naver.com"
    
    wb.save("./test/DISK_SUM.xlsx")


def cpu_fx():
    wb = op.load_workbook(r"C:\Users\홍익과학기술\Desktop\Exel_aZ\test\CPU_SUM.xlsx") #Workbook 객체 생성
    ws = wb["Sheet1"] #WorkSheet 객체 생성("무" Sheet)

    #"B1" Cell에 입력하기
    ws.cell(row=1, column=2).value = "입력테스트1"

    #"G1" Cell에 입력하기
    ws["G6"].value = "=AVERAGE('CPU_ALL:CPU_ALL (4)'!J59)"

    wb.save("./test/CPU_SUM.xlsx")
    
# 함수 실행
disk_Az()
cpu_Az()
mem_Az()

disk_fx()
#cpu_fx()
