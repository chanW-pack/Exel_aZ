import main
#import test6
from time import sleep

def start_pro():
  # 결과파일 저장 디렉터리 생성
  sleep(1)
  #test6.loading()
  print("디렉터리 생성 중 ...")
  # 실행 함수
  main.createFolder('./cw_test')
  #test6.loading_out()

def start_pro2():
  # AzData_pull()
  sleep(1)
  #test6.loading()
  print("DISK 관련 데이터를 수집합니다.")
  sleep(0.5)
  print(".")
  sleep(0.5)
  print("...")
  sleep(0.5)
  print(".....")
  sleep(0.5)
  print(".......")
  # 실행 함수
  main.disk_Az()
  #test6.loading_out()

  sleep(1)
  #test6.loading()
  print("CPU 관련 데이터를 수집합니다.")
  sleep(0.5)
  print(".")
  sleep(0.5)
  print("...")
  sleep(0.5)
  print(".....")
  sleep(0.5)
  print(".......")
  # 실행 함수
  main.cpu_Az()
  #test6.loading_out()

  sleep(1)
  #test6.loading()
  print("Memory 관련 데이터를 수집합니다.")
  sleep(0.5)
  print(".")
  sleep(0.5)
  print("...")
  sleep(0.5)
  print(".....")
  sleep(0.5)
  print(".......")
  # 실행 함수
  main.mem_Az()
  #test6.loading_out()

  # 새 파일 계산
  sleep(1)
  #test6.loading()
  print("Disk 관련 데이터 분석을 진행합니다.")
  main.disk_fx()
  #disk_chart()
  #disk_chart_make()
  #test6.loading_out()

  sleep(1)
  #test6.loading()
  print("CPU 관련 데이터 분석을 진행합니다.")
  main.cpu_fx()
  #test6.loading_out()

  sleep(1)
  #test6.loading()
  print("Memory 관련 데이터 분석을 진행합니다.")
  main.mem_Calculation()
  main.mem_fx()
  #test6.loading_out()

  sleep(2)
  print("작업이 완료되었습니다!")
  print("작업이 완료되었습니다!")
  print("작업이 완료되었습니다!")

    