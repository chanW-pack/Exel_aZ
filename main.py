import glob
import win32com.client
import os
import datetime
from time import sleep
import openpyxl as op 
from openpyxl.styles import Font, PatternFill, GradientFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference, BarChart, AreaChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.shapes import GraphicalProperties

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
    # win32com(pywin32)를 이용해서 엑셀 어플리케이션을 연다.
    excel = win32com.client.Dispatch("Excel.Application")
    # 실제 작동하는 것을 보고 싶을 때
    #excel.Visible = True 

    # 엑셀 어플리케이션에 새로운 Workbook 추가
    wb_new = excel.Workbooks.Add() 

    # glob 모듈로 원하는 폴더 내의 모든 xlsx 파일의 경로를 리스트로 반환
    list_filepath = glob.glob(r'C:\Users\pp\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    # 엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
    for filepath in list_filepath:

        # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
        wb = excel.Workbooks.Open(filepath)

        # 새로 만든 엑셀 파일에 추가
        # 추출할wb.Worksheets("추출할 시트명").Copy(Before=붙여넣을 wb.Worksheets("기준 시트명")
        wb.Worksheets("DISK_SUMM").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    # 취합한 엑셀 파일을 저장
    wb_new.SaveAs(r"C:\Users\pp\Desktop\Exel_aZ\test\DISK_SUM.xlsx")

    # 켜져있는 엑셀 및 어플리케이션 모두 종료
    excel.Quit()


def cpu_Az():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 
    list_filepath = glob.glob(r'C:\Users\pp\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    for filepath in list_filepath:

        wb = excel.Workbooks.Open(filepath)
        wb.Worksheets("CPU_ALL").Copy(Before=wb_new.Worksheets("Sheet1"))

    wb_new.SaveAs(r"C:\Users\pp\Desktop\Exel_aZ\test\CPU_SUM.xlsx")

    excel.Quit()


def mem_Az():
    excel = win32com.client.Dispatch("Excel.Application")
    wb_new = excel.Workbooks.Add() 
    list_filepath = glob.glob(r'C:\Users\pp\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    for filepath in list_filepath:
        
        wb = excel.Workbooks.Open(filepath)
        wb.Worksheets("MEM").Copy(Before=wb_new.Worksheets("Sheet1"))

    wb_new.SaveAs(r"C:\Users\pp\Desktop\Exel_aZ\test\MEM_SUM.xlsx")

    excel.Quit()


# font 변수
global Az_val_1, Az_val_2, add_inq, Title, hide_font

Az_val_1 = Font(name='맑은 고딕', bold=True, size=24)
Az_val_2 = Font(name='맑은 고딕', size=24)
Title = Font(name='맑은 고딕', bold=True, size=20)
add_inq = Font(name='맑은 고딕', bold=True, size=11)
hide_font = Font(color='FFFFFF')

# cell 테두리 변수
global cell_box
cell_box = Border(
    left=Side(border_style="medium", color='00000000'),
    right=Side(border_style='medium', color='00000000'),
    bottom=Side(border_style='medium', color='00000000'),
    top=Side(border_style='medium', color='00000000')
)

# 현재 날짜
global d_today
d_today = datetime.date.today()


def disk_fx():
    wb = op.load_workbook(r"C:\Users\pp\Desktop\Exel_aZ\test\DISK_SUM.xlsx") #Workbook 객체 생성
    ws = wb["Sheet1"] 

    # cell 병합
    ws.merge_cells("B2:E3")
    ws.merge_cells("F2:I3")
    ws.merge_cells("B4:E5")
    ws.merge_cells("F4:I5")

    ws.merge_cells("J2:Q5")
    ws.merge_cells("R2:U5")

    # font 지정
    ws['B2'].font = Az_val_1
    ws['F2'].font = Az_val_1
    ws['B4'].font = Az_val_2
    ws['F4'].font = Az_val_2
    ws['J2'].font = Title
    ws['R2'].font = add_inq

    # cell 테두리 지정
    #ws['A1'].border = cell_box
    #ws['E1'].border = cell_box
    #ws['E10'].border = cell_box
                   
    # cell 배경색 추가
    ws['B2'].fill = PatternFill(start_color='C6E0B4', fill_type = 'solid') 
    ws['F2'].fill = PatternFill(start_color='B4C6E7', fill_type = 'solid')
    ws['J2'].fill = PatternFill(start_color='FFE699', fill_type = 'solid')
    ws['R2'].fill = PatternFill(start_color='D6DCE4', fill_type = 'solid')

    # cell 정렬 적용 // wrap_text=자동 줄바꿈
    ws['R2'].alignment = Alignment(horizontal="left", vertical="top", wrap_text = True)
    ws['J2'].alignment = Alignment(horizontal="center", vertical="center", wrap_text = True)

    # cell 값 적용
    
    
    ws["B2"].value = "Disk Read KB/s"
    ws["F2"].value = "Disk Write KB/s"
    ws["B4"].value = "=AVERAGE('DISK_SUMM:DISK_SUMM (4)'!B59)"
    ws["F4"].value = "=AVERAGE('DISK_SUMM:DISK_SUMM (4)'!C59)"
    ws["J2"].value = f"Disk Read/Write 월 평균 계산기               {d_today}"
    ws["R2"].value = "문의사항                                                     TS팀 박찬우                                              tel: 010-9085-0857             chanwoo9730@naver.com"

    # 결과 시트 수정
    wb.move_sheet(ws, -4) 
    
    ws = wb["Sheet1"]
    ws.title = "Result"
    
    wb.save("./test/DISK_SUM.xlsx")
    sleep(3)



# disk 차트 생성
def disk_chart():
    # Workbook 객체 생성 (win32com)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(r"C:\Users\pp\Desktop\Exel_aZ\test\DISK_SUM.xlsx") 
    # 시트 데이터 변수로 저장
    disk_sum  = wb.Worksheets["DISK_SUMM"]
    disk_sum2 = wb.Worksheets["DISK_SUMM (2)"]
    disk_sum3 = wb.Worksheets["DISK_SUMM (3)"]
    disk_sum4 = wb.Worksheets["DISK_SUMM (4)"]
    ws1 = wb.Worksheets['Result']
    # 복사 및 숨기기
    disk_sum.Range("A1:D57").Copy(ws1.Range("W1:Z57"))
    disk_sum2.Range("A2:D57").Copy(ws1.Range("W58:Z114"))
    disk_sum3.Range("A2:D57").Copy(ws1.Range("W115:Z171"))
    disk_sum4.Range("A2:D57").Copy(ws1.Range("W172:Z228"))
    #ws1['W1:Z230'].font = hide_font
    excel.Quit()
    sleep(3)
    

def disk_chart_make():
    # Workbook 객체 생성 (openpyxl)
    wb_op = op.load_workbook(r"C:\Users\pp\Desktop\Exel_aZ\test\DISK_SUM.xlsx") 
    ws_op = wb_op["Result"]
     
    for i in range(228):
        ws_op.append([i])

    c1 = LineChart()
    c1.title = "Disk total KB/s" ## 차트 타이틀
    c1.style = 13 # 1~48 차트 스타일
    c1.width = 38.1 ## 차트 폭
    c1.height = 15 ## 차트 높이
    #c1.y_axis.title = 'Size' # y축 라벨
    c1.y_axis.crossAx = 500 ## y축 교차축

    ## 날짜 축 설정
    c1.x_axis = DateAxis(crossAx=100) ## x축을 날짜 축으로 설정
    c1.x_axis.title = "Date" # x축 라벨
    #c1.x_axis.number_format = 'd-mmm' ## 일 - 월 세글자만
    #c1.x_axis.majorTimeUnit = "days" ## 눈금 단위 {'months', 'years', 'days'}

    ## 그림 영역 배경색 설정 
    #props = GraphicalProperties(solidFill="999999") 
    #c1.plot_area.graphicalProperties = props 

    Disk_read = Reference(ws_op, min_row=1, max_row=228, min_col=24, max_col=25)
    #Disk_write = Reference(ws_op, min_row=1, max_row=228, min_col=25, max_col=25)
    #IO_sec = Reference(ws_op, min_row=1, max_row=228, min_col=26, max_col=26)

    #차트 객체 생성
    c1.add_data(Disk_read, titles_from_data=True)
    
    # Style the lines
    s2 = c1.series[1]
    s2.smooth = True ## 라인을 매끄럽게 만듬.
    s2.graphicalProperties.line.width = 100000 # width in EMUs
    s2.graphicalProperties.line.solidFill = "B4C6E7"
 
    s3 = c1.series[0]
    s3.smooth = True ## 라인을 매끄럽게 만듬.
    s3.graphicalProperties.line.width = 100000 # width in EMUs
    s3.graphicalProperties.line.solidFill = "C6E0B4"
    
    # 차트 추가
    ws_op.add_chart(c1, "B7")
    wb_op.save("./test/DISK_SUM.xlsx")
    sleep(3)



def cpu_fx():
    wb = op.load_workbook(r"C:\Users\pp\Desktop\Exel_aZ\test\CPU_SUM.xlsx") #Workbook 객체 생성
    ws = wb["Sheet1"] #WorkSheet 객체 생성("무" Sheet)

    #"B1" Cell에 입력하기
    ws.cell(row=1, column=2).value = "입력테스트1"

    #"G1" Cell에 입력하기
    ws["G6"].value = "=AVERAGE('CPU_ALL:CPU_ALL (4)'!J59)"

    wb.save("./test/CPU_SUM.xlsx")
    
# 함수 실행
#disk_Az()
#cpu_Az()
#mem_Az()

#disk_fx()
#disk_chart()
#disk_chart_make()
#cpu_fx()
