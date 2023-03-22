import openpyxl

# 엑셀 파일 열기'
workbook = openpyxl.load_workbook('./files/localhost_230102_0000.nmon.xlsx')

# 시트 선택하기
sheet = workbook['DISK_SUMM']

# 함수가 있는 셀 가져오기
cell = sheet['C59']

# 수식 계산하기
workbook.calculation = True

# 계산된 값 가져오기
result = cell.value

# 결과 출력하기
print(result)

# 엑셀 파일 저장하기
workbook.save('your_excel_file.xlsx')