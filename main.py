import os
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook

# 다수의 파일을 불러오기 위한 해당 디렉터리 내 파일 찾기
path = "./files"
file_list = os.listdir(path)
#print(file_list)

# 하나의 엑셀 파일, 시트 가져오기
#wb = openpyxl.load_workbook('openpyxltest.xlsx')
## ws = wb.active
#ws = wb['Sheet']

# 다수 파일 엑셀, 시트 가져오기
results = []
for file_name_raw in file_list:
    file_name = "./files/" + file_name_raw
    wb = load_workbook(filename=file_name, data_only=True)
    ws = wb['DISK_SUMM']
    result = []
    result.append(file_name_raw)
    results.append(result)


# 셀 인덱스를 반복문으로 달림
cell_range = ws['C2':'C5']
#print(cell_range) column 가져오기
for col_cell in cell_range:
    for cell in col_cell:
        print(cell.value) # 인덱스 값 확인
    

# 엑셀로 저장
#wb = Workbook()
#ws = wb.active
#for i in results:
#    ws.append(i)
#wb.save("results.xlsx")
#CPU 사용률
#Mem 사용률
#DIsk IO 월평균

#https://bebeya.tistory.com/entry/%ED%8C%8C%EC%9D%B4%EC%8D%AC-%EC%97%91%EC%85%80-%EC%BB%AC%EB%9F%BC%EA%B0%92-%EA%B0%80%EC%A0%B8%EC%98%A4%EA%B8%B0-%EB%8D%B0%EC%9D%B4%ED%84%B0-%EB%84%A3%EA%B8%B0%ED%95%9C%EC%A4%84%EC%94%A9-%EB%A6%AC%EC%8A%A4%ED%8A%B8

# 칼럼 값 더하기
#https://2toy.net/entry/python-openpyxl-%EC%85%80%EC%9D%98-%ED%95%A9%EA%B3%84-%EA%B0%92-%EC%88%98%EB%9F%89-list-%EC%B4%9D-%ED%95%A9%EA%B3%84-%EA%B5%AC%ED%95%98%EA%B8%B0