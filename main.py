# step1.관련 모듈 및 패키지 import
import glob
import win32com.client
import os
 
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
    list_filepath = glob.glob(r'C:\Users\chanw\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    # step5.엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
    for filepath in list_filepath:

        # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
        wb = excel.Workbooks.Open(filepath)

        # 새로 만든 엑셀 파일에 추가
        # 추출할wb.Worksheets("추출할 시트명").Copy(Before=붙여넣을 wb.Worksheets("기준 시트명")
        wb.Worksheets("DISK_SUMM").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    # step6. 취합한 엑셀 파일을 "통합 문서"라는 이름으로 저장
    wb_new.SaveAs(r"C:\Users\chanw\Desktop\Exel_aZ\test\DISK_SUM.xlsx")

    # step7. 켜져있는 엑셀 및 어플리케이션 모두 종료
    excel.Quit()


def cpu_Az():
    # step2.win32com(pywin32)를 이용해서 엑셀 어플리케이션 열기
    excel = win32com.client.Dispatch("Excel.Application")
    #excel.Visible = True #실제 작동하는 것을 보고 싶을 때 사용

    # step3.엑셀 어플리케이션에 새로운 Workbook 추가
    wb_new = excel.Workbooks.Add() 

    # step4.glob 모듈로 원하는 폴더 내의 모든 xlsx 파일의 경로를 리스트로 반환
    list_filepath = glob.glob(r'C:\Users\chanw\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    # step5.엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
    for filepath in list_filepath:

        # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
        wb = excel.Workbooks.Open(filepath)

        # 새로 만든 엑셀 파일에 추가
        # 추출할wb.Worksheets("추출할 시트명").Copy(Before=붙여넣을 wb.Worksheets("기준 시트명")
        wb.Worksheets("CPU_ALL").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    # step6. 취합한 엑셀 파일을 "통합 문서"라는 이름으로 저장
    wb_new.SaveAs(r"C:\Users\chanw\Desktop\Exel_aZ\test\CPU_SUM.xlsx")

    # step7. 켜져있는 엑셀 및 어플리케이션 모두 종료
    excel.Quit()


def mem_Az():
    # step2.win32com(pywin32)를 이용해서 엑셀 어플리케이션 열기
    excel = win32com.client.Dispatch("Excel.Application")
    #excel.Visible = True #실제 작동하는 것을 보고 싶을 때 사용

    # step3.엑셀 어플리케이션에 새로운 Workbook 추가
    wb_new = excel.Workbooks.Add() 

    # step4.glob 모듈로 원하는 폴더 내의 모든 xlsx 파일의 경로를 리스트로 반환
    list_filepath = glob.glob(r'C:\Users\chanw\Desktop\Exel_aZ\files\*.xlsx', recursive=True)

    # step5.엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
    for filepath in list_filepath:

        # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
        wb = excel.Workbooks.Open(filepath)

        # 새로 만든 엑셀 파일에 추가
        # 추출할wb.Worksheets("추출할 시트명").Copy(Before=붙여넣을 wb.Worksheets("기준 시트명")
        wb.Worksheets("MEM").Copy(Before=wb_new.Worksheets("Sheet1"))
  
    # step6. 취합한 엑셀 파일을 "통합 문서"라는 이름으로 저장
    wb_new.SaveAs(r"C:\Users\chanw\Desktop\Exel_aZ\test\MEM_SUM.xlsx")

    # step7. 켜져있는 엑셀 및 어플리케이션 모두 종료
    excel.Quit()
    
# 함수 실행
disk_Az()
cpu_Az()
mem_Az()