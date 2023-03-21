# 엑셀 함수 사용(평균 구하기)
import  openpyxl  as  op  #openpyxl 모듈 import

wb = op.load_workbook("./files/localhost_230102_0000.nmon.xlsx") #Workbook 객체 생성
ws = wb["DISK_SUMM"] #"DISK"시트 객체 생성

#print("#rows 출력")
#for  row_rng  in  ws.rows:
#    for cell in row_rng:
#       print(cell.value) #각 행에 대한 1차원 배열 출력(위치정보)


#엑셀 함수를 실제 Cell에 써보기
ws["H3"].value = "=AVERAGE(C3:C57)"

wb.save("result.xlsx")