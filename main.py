import os
from openpyxl import load_workbook
from openpyxl import Workbook

path = "./files"
file_list = os.listdir(path)
print(file_list)

results = []
for file_name_raw in file_list:
    file_name = "./files/" + file_name_raw
    wb = load_workbook(filename=file_name, data_only=True)
    ws = wb['DISK_SUMM']
    result = []
    result.append(file_name_raw)
    result.append(ws['B2:B55'].value)
    result.append(ws['C2:C55'].value)
    result.append(ws['E2'].value)
    result.append(ws['E3'].value)
    results.append(result)
print(results)


#wb = Workbook()
#ws = wb.active
#for i in results:
#    ws.append(i)
#wb.save("results.xlsx")
#CPU 사용률
#Mem 사용률
#DIsk IO 월평균

#https://bebeya.tistory.com/entry/%ED%8C%8C%EC%9D%B4%EC%8D%AC-%EC%97%91%EC%85%80-%EC%BB%AC%EB%9F%BC%EA%B0%92-%EA%B0%80%EC%A0%B8%EC%98%A4%EA%B8%B0-%EB%8D%B0%EC%9D%B4%ED%84%B0-%EB%84%A3%EA%B8%B0%ED%95%9C%EC%A4%84%EC%94%A9-%EB%A6%AC%EC%8A%A4%ED%8A%B8