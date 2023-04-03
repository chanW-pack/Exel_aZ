# Exel_aZ
nmon 엑셀 분석 프로그램

# 엑셀 데이터 분석 프로그램

### **개요**

---

젖소개량사업소 서버 정기점검 업무 중, 한달 리소스 사용 내역을 평균화 하는 작업이 존재하는데 이를 편리하게 개선시키기 위해서 개발을 진행하였다.

Linux 기반 server의 경우 crontab을 이용한 스케쥴링으로 nmon data를 생성하고 있고, 
Windows의 경우 성능 모니터를 이용하여 주기적으로 파일을 생성해 리소스 모니터링을 진행하고 있다.

해당 파일들은 csv 혹은 xlsx 엑셀파일로 변환이 가능한데, 기본적으로 모든 정보를 수집하기 때문에 작업인원이 손수 데이터를 분리하고 계산하여야 한다.

이를 자동화하여 필요한 데이터가 기록된 시트에서 데이터를 추출하여 계산하고 새로운 엑셀 파일로 저장하는 프로그램을 구상하고 있다.

```bash
초기 아키텍쳐 구상

엑셀파일 ------>  python에서 데이터 추출 -------> 해당 데이터 계산 -------> 엑셀파일
```

### **개발 진행 과정**

---

**개발 순서**

1. 취합  
   약 한달마다 nmon data가 4개씩 떨어진다.(일주일단위) 이를 통합하여 하나의 파일로 생성
2. 분석  
   통합된 데이터를 불러온 뒤 계산 > 새로운 파일로 값을 저장
3. 시각화  
   시간 남으면 데이터를 시각화하여 보고 메일 첨부

### `**2023.03.21~**`

데이터를 파이썬에서 가져와 계산하려다 더 좋은 방안이 생각났다.

```bash
wb = op.load_workbook(r"test.xlsx") #워크북 객체 생성
ws = wb.active #시트 객체 생성

#엑셀 함수를 실제 Cell에 써보기
ws["E11"].value = "=SUM(C:C)"

wb.save("result.xlsx")
```

위 코드는 굳이 파이썬에서 데이터를 가져와 계산하고 내보내는것이 아닌

엑셀 함수를 사용하여 즉시 원하는 셀에 계산값이 나오는것이다.

여기서 문제는 파이썬내에서 해당 셀의 정보를 가져올때 결과값이 아닌 함수 자체를 가져오는 것이다.

이는 시트 객체를 생성할 때 셀 내용을 데이터로 가져오는 방법으로 해결 가능하다.
(해당 내용은 0324~ 부분에서 확인 가능)

한 파일의 한 시트의 한 칼럼의 평균을 나타내는 방법은 찾았으니,

이제 여러 파일을 합치는 방법과 이를 나타낼 방법을 찾아야한다.

1. 파일을 합칠 시 어떻게?? (한 파일에 필요한 시트만?? 아님 한 시트에 해댱 내용들 정리??)
ㄴ 새로운 시트를 만들고 데이터를 abcd열 쭉~ 나열할까..? 이후 얘네들을 엑셀 함수로 정리하면 되지않을까

또한 이를 토대로 그래프도 만들 수 있으면 일석 이조가 아닌가?

```bash
구현 성공
> 시트 내 원하는 위치의 평균 구하기
> 새로운 엑셀 파일로 저장

조사할 내용
1. 다수의 엑셀 파일 불러오기
2. 새로운 엑셀 파일 생성하기 (완료)
3. 필요한 데이터 불러와 새로운 시트에 차례대로 저장
4. 이를 분석

(https://approximation.tistory.com/32)
해당 자료로 보면
다량의파일을 읽고, 이를 새로운 파일로 만들수 있음
필요한 자료만 불러들여, 생성이 가능한 것이다.
또한 각 정보들의 시트를 새로 만들어 저장 가능

즉, 다수의 파일을 불러오고
CPU. MEM, DISK 별 시트를 나눈다.
이후 결과값을 뱉어내고, 저장.. 끝...
```

현재 다수의 파일의 시트를 한 파일로 저장하는 기능 완성
> 근데 시트는 안합쳐짐.. ex disk disk1 disk2... 로 계속 추가 생성된다.

### `**2023.03.22~**`

문제가 발생하였다. 

python으로 함수가 사용된 셀의 정보를 가저오면 데이터가 아닌 함수 자체가 가져와진다. 
(결과값이 출력되는게 아닌 해당 함수(=SUM('A3:B2')) 형식으로 나타나게됨)

**해결방안** 
1. 그냥 다시 처음부터해서 파이썬에서 계산하고 셀에 붙이기
2. 계산할 데이터를 전부 차례대로 불러와 직접 함수 호출로 합치기 > 그나마 나아보임

정리하면,

```python
구현 성공
> 시트 내 원하는 위치의 평균 구하기
> 새로운 엑셀 파일로 저장
(추가!) > 함수를 사용하여 시트내 원하는 데이터 평군 구하기
(추가!) > 다수의 엑셀 파일을 불러와 설정한 시트로 구성된 새로운 파일 생성

조사할 내용
1. 다수의 엑셀 파일 불러오기 (완료)
2. 새로운 엑셀 파일 생성하기 (완료)
3. 필요한 데이터 불러와 새로운 시트에 차례대로 저장
4. 이를 분석

3. 에서 여러 시트에 적용되게 하고 다른 시트 참고 기능을 사용하면 구성이 가능하지 않을까?
즉, 코드는 하드코딩이지만 = AVAG DISK  DISK(1) DISK(2) 등 형식으로 진행할 수 있는지 조사해보겠다.
```

가능하다. 

(https://ko.extendoffice.com/documents/excel/2613-excel-average-cells-on-multiple-sheets.html)

그렇다면 데이터 로직은 이러하다.

```python
초기 아키텍쳐 구상

엑셀파일 ------>  python에서 데이터 추출 -------> 해당 데이터 계산 -------> 엑셀파일

현재 아키텍쳐 구상

다수의 엑셀 파일 --> DISK 시트만 새로운 엑셀로 저장 --> 해당 엑셀의 어느 셀에서 함수로 평군 계산

구성도
0. 다수의 엑셀 파일을 불러옴
1. 새로운 엑셀 파일 생성 (CPU, MEM, DISK의 3개의 시트 존재, 시트마다 해당 값들이 순서대로 나열)
2. 해당 값 어딘가에 평균값 생성
0321까지 한것.
https://approximation.tistory.com/32
https://dotsnlines.tistory.com/562
https://wikidocs.net/156319
```

분석할 파일 개수는 일정(4~5개 가량) 하기 때문에 하드코딩으로 진행하였다.
(젖소개량사업소의 작업에서는 이를 변경할 일이 없겠지만, 더 높은 완성도(낮은 위험도)와 추후 다른 기능으로 사용 가능하도록 개선이 필요함, 그리고 핵심적으로 하드코딩하면 멋이 없음)

### `**2023.03.24~**`

현재 files 디렉터리 내 xlsx 파일을 취합하여 해당 시트에 해당하는 각각의 개별 파일로 일괄 저장되는 기능을 완성하였다. 

즉, test 디렉터리에 cpu, mem, disk 등의 파일로 나눠서 저장된다.

이 파일내에 sheet1 이라는 기본 시트를 생성하는데 해당 시트에 평균 값을 저장하면 될듯

```
1. CPU의 경우
j:59 cell에 1주일 cpu 평균값이 존재함.

2. DISK R/W의 경우
B59, C59 존재

CPU_ALL ~ CPU_ALL (4) 까지의 시트에 있는 j:59 cell값의 평균을 구하면 될듯

셀 및 시트삭제 기능이 존재함.
계산완료후에는 나머지 시트 지울까??
> 함수로 계산하였기 때문에 기존 값이 삭제되면 안된다.
>>>>추가.. 계산 결과값을 데이터값으로 복사하는 기능을 찾았다.
wb = load_workbook( )로 엑셀 파일을 불러올 때 data_only 설정을 진행하면 된다나.
(주의사항으로는 즉시 적용이 안된다. 엑셀을 한번 켰다 저장해야한단다.)
ex) wb = load_workbook("formula.xlsx", data_only=True)
```

이로 현재까지는 데이터들의 평균값을 분석하는 기능 구현은 완료되었다.

이후 추가할 기능으로는,

1. 차트생성
2. PDF 저장기능 추가
3. GUI 추가
4. exe 파일로 구축

정도가 있을 듯 하다.

`!! 현재 발견된 버그/오류들..`

[Python - Openpyxl 병합셀 관련버그 해결](https://m.blog.naver.com/soldatj/221124415354)

<aside>
💡 해결해서올립니다

''' min_col, min_row, max_col, max_row =

range_boundaries(range_string)

rows = range(min_row, max_row+1)

cols = range(min_col, max_col+1)

cells = product(rows, cols)

all but the top-left cell are removed

for c in islice(cells, 1, None):

if c in self._cells:

del self._cells[c]'''

worksheet.py 파일의 위에부분을 주석처리하면 버그없어지네요 스택오버플로우에서 알아냈습니다

</aside>

`disk 차트 생성 실패 openpyxl 버전문제..`

```python
# disk 차트 생성
def disk_chart():
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
    
   
    #W1:Z228 까지의 데이터를 차트로 생성 / LineChart
    line_value = Reference(ws1, min_row=1, max_row=228, min_col=23, max_col=26)
    line_chart = LineChart() #차트 종류 설정(BarChart, LineChart, pie ...)
    line_chart.add_data(line_value) # 계열 > 영어, 수학 (제목에서 가져옴)

    #차트 제목
    line_chart.title = "Disk total KB/s " 
    line_chart.style = 20 # 미리 정의된 스타일 적용
    #line_chart.y_axis.title = "점수" #Y축의 점수
    #line_chart.x_axis.title = "번호" #X축의 번호

    ws1.add_chart(line_chart, "B7") #차트 넣을 위치

 #ws1['W1'].font = hide_font

    wb.save("./test/DISK_SUM.xlsx")
```

TypeError: expected <class 'str'> 에러로 기능 동작이 불가하다.

이는 openpyxl 버전문제인데, 정보가 너무 부족하여 다른 방법을 찾아봐야할듯

`함수 연결 이슈`

현재 차트까지 완성하였으나, 함수들을 따로 작동하면 잘 되는데 한번에 이어지지가 않음

1. 처음 xlsx 파일 생성하고 파일을 찾지 못함
> 시트 합친 파일은 직접 한번 열어서 저장을 다시 해야한다.
2. 데이터 정제하고 복사 → 차트만들기 부분에서 막힘
> win32랑 openpyxl을 사용하여 뭐가 제대로 안껴져서 겹쳐서 오류나는듯함.
하나씩 실행하면 잘 됨

### `**2023.03.25~**`

문제 해결 방안으로 win32를 사용하는 함수를 합치기로 생각하였다.

이유는 win32를 실행하고 저장하면 내가 직접 열고 저장을 다시 하지 않는 한 win32로 열지 못하는 이슈가 있다.

현재 win32를 사용하는 함수는 처음 시트 저장 함수를 제외하곤 disk_chart함수(다른 시트 cell 내용 복사)밖에 없다.

이를 openpyxl로 구현하면 해결될 듯 하다.

### `2023.03.26~`

이틀 후에 정기점검 작업이 있어 실 사용 테스트를 해보고 싶기 때문에 일단 문제가 있는 그래프화는 후순위로 미뤄두고, 기능 구현에 중점을 맞춰 진행하겠다. (목적인 평균값 계산은 달성했으니…)

그래서 오늘은

1. CPU, MEM 평균계산
2. exe 실행파일로 만들기

까지 완성해보도록 하겠다.

(추후에 차트와 GUI까지 진행.)

기존 chart 생성 부분을 담당하는 함수(def disk_chart(), def disk_chart_make())를 main.py에서 제외한 뒤, main_chart.txt 파일에 저장하였다.

CPU는 Avg이 원래 나와있으므로 즉시 완성하였고, mem의 경우 평균값을 계산해주지 않기 때문에 직접 계산이 필요하다

<aside>
💡 저장된 데이터
- **memtotal**
- hightotal
- lowtotal
- swaptotal
- **memfree**
- highfree
- lowfree
- swapfree
- memshared
- **cached**
- active
- bigfree
- **buffers**
- swapcached
- inactive

</aside>

해당 정보로 사용률 계산하는법을 간단히 설명하자면,

`(MemFree+ Buffers + Cached) / memtotal` 로 퍼센트를 확인할 수 있겠다.

각 시트별로 평균 퍼센트를 계산한 뒤 reslut 시트에서 최종 평균값을 구하면 될 듯 싶다.

`**16:42**`

기능구현 완료. 

현재 Disk, CPU. Memory 월 평균 사용률 계산이 완료됨.

일단 한 주의 데이터라도 차트화할까 생각중임. 너무 밋밋함

(주의사항)

현재 57라인까지 계산으로 되어있어 57개까지 리소스체크가 안되어있음 안됨 > 방안 모색 필요

일단 차트 이미지로 떼우기로 했는데…

### 문제발생

현재 파이썬 내 코드의 경로가

`**wb = op.load_workbook(r"C:\Users\pp*\D*esktop\Exel_aZ\test\CPU_SUM*.*xlsx") #Workbook 객체 생성**`

해당 형식으로 나타나있는데, 이는 windows pp user로 접속했을 때는 정상 작동되나,

타 컴퓨터에서 다른 user면 내용을 변경해야 하는 번거로움이 있다.

해결

해당 내용은 절대경로 설정으로 해결 가능하다.

```python
import os
 
print(format(os.getlogin()))
 
# C:\User\lucidyul\~~~~
# lucidyul

## 사용예시
import os
 
path = "C:/Users/{}/desktop".format(os.getlogin())  # {}부분에 사용자 이름
 
print(path) # C:/Users/lucidyul/desktop

>> 위 방법이 안되서 따로 구현하였음
예로,
    path = r"C:\Users\user\Desktop\Exel_aZ"
    path = r"C:\Users\*\Desktop\Exel_aZ"
이 처럼 "?" 나 "*" 로 지정하면 된다.

>>> 위위 방법 되네.. 내가 문자열에 안넣고 바보짓했음
```

기능완성하였고, 실행기능 구현중에 있음

여기서 중요한 포인트는 현재 코드는 Exel_aZ 디렉터리가 하드코딩되오있음.

exe 실행된 위치에 디렉터리가 생성되고 안에 쌓이게 변경해야됨

1. GUI로 엑셀이 있는 폴더 선택

***추가로 disk 부분 kb/s 을 mb/s 로 변경함

>> 기존 값에서 나누기 1000 넣음

추가 2 

경로 설정시 .format으로 user 부분은 수정 완료하였으나,

tkinter을 사용하여 GUI로 경로를 불러오는 부분을 적용 완료하였음.

그래서 불러오고나서 떨어지는 파일은 실행파일이 있는 디렉터리에 새로운 디렉터리를 생성하여 거기에 나옴

이슈

GUI로 경로를 불러들이는 과정에서 해당 함수가 각 시트를 불러오는 함수마다 적용되어 있어 계속 경로를 지정하는 경우가 있었는데, 이를 한번에 하기 위해 전역변수를 활용하여 해결 완료하였음

>

현재 GUI 환경까지 완성

![Untitled](https://user-images.githubusercontent.com/84123877/229450768-cf096dbd-65c6-4d76-943f-52b2844f9f07.png)

이를 하나의 파일로 만드는 것도 테스트 완료 (다른 컴퓨터에서 정상 작동^^)

문제는 현재 파일찾기와 시작, 닫기는 정상작동되나(버튼은 잘됨) 작업내역과 진행상황이 반영이 안됨.

낼 할거

이미지 관리하자

폴더에넣든 코드를바꾸든

### `**23.03.28~**`

터미널 구현 완료, 로딩창 구현 완료, 파일 다수 선택 기능 구현 완료, 메세지박스 구현 완료

문제: 터미널 및 로딩창 구현은 완료되었으나, 프로그램일 실행되면 멈췄다가 실행 완료될때 한번에나옴..

> 구현쪽에서 문제되는건 아닌거같고 프로그램이 실행중일때 UI창이 렉? 멈춤? 현상이 있는듯함.

파일 다수 선택 기능은 체크박스를 생성하여 체크하고 찾기를 누르게 설정할 예정임

![Untitled 1](https://user-images.githubusercontent.com/84123877/229450747-0c65f119-0ad7-4494-8dd5-02ec9513ce71.png)

- 문구 추가

![Untitled 2](https://user-images.githubusercontent.com/84123877/229450754-1f36e9a5-6bb8-4a3f-bc82-500daf33c3c7.png)

- 상태메시지 추가

![Untitled 3](https://user-images.githubusercontent.com/84123877/229450758-ca38c501-e66f-40a0-97ff-691e185a9364.png)

- 터미널 output 나타내기

`개발 완료 시 예상 화면(03.24)`

![Untitled 4](https://user-images.githubusercontent.com/84123877/229450763-f577cbc6-e789-4937-9ec7-700d020204f4.png)
