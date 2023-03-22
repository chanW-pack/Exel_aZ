import os
from openpyxl import load_workbook
from openpyxl import Workbook

## 다수의 파일 불러오기
#path = "./files"
#file_list = os.listdir(path)
##print(file_list)
#results = []
#for file_name_raw in file_list:
#    file_name = "./files/" + file_name_raw
#    wb = load_workbook(filename=file_name, data_only=True)
#    ws = wb['DISK_SUMM']
#    result = []
#    result.append(file_name_raw)
#    result.append(ws['B2'].value)
#    result.append(ws['C2'].value)
#    result.append(ws['E2'].value)
#    result.append(ws['E3'].value)
#    results.append(result)
#print(results)


print("111111111")
# 파일 단일 불러오기
file1 = './files/localhost_230102_0000.nmon.xlsx'
dan_wb = load_workbook(file1)
#dan_wb = dan_ws.active  //첫 시트
dan_ws = dan_wb['DISK_SUMM']

 
idx = []
for m in range(0,10):
    col1 = dan_ws.cell(row=m+1,column=1).value
    print(col1)
    idx.append(col1)
 
#print(sum(idx))
#print(len(idx))
print('#####')


wbb = load_workbook(file1)
wss = wbb['DISK_SUMM']
#하나에 컬럼에 대한 값을 가져오기
col_BB = wss["C1:C5"] #영어 column 만 가지고 오기
#col_B 값 출력하기

for cell_0 in col_BB:
    print(cell_0.value)  #가지고온 col_B의 인덱스 값 확인 = print(col_B)





#wb = Workbook()
#ws = wb.active
#for i in results:
#    ws.append(i)
#wb.save("results.xlsx")
#CPU 사용률
#Mem 사용률
#DIsk IO 월평균

#https://bebeya.tistory.com/entry/%ED%8C%8C%EC%9D%B4%EC%8D%AC-%EC%97%91%EC%85%80-%EC%BB%AC%EB%9F%BC%EA%B0%92-%EA%B0%80%EC%A0%B8%EC%98%A4%EA%B8%B0-%EB%8D%B0%EC%9D%B4%ED%84%B0-%EB%84%A3%EA%B8%B0%ED%95%9C%EC%A4%84%EC%94%A9-%EB%A6%AC%EC%8A%A4%ED%8A%B8