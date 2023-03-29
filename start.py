import tkinter as tk
import sys
import main
import cw_process
from tkinter import *
import tkinter.ttk as ttk
from tkinter.messagebox import *
from time import sleep
import os
from tkinter import filedialog


# 텍스트 위젯에 출력할 클래스 생성
class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")
        self.widget.yview("end")


# 지정 경로 출력
#def browse_dest_path():
#    folder_selected = filedialog.askdirectory()
    # 사용자가 취소를 누를 때
#    if folder_selected == "": 
#        print("폴더 선택 취소")
#        return
    #print(folder_selected)
#    txt_dest_path.delete(0, END)
#    txt_dest_path.insert(0, folder_selected)


def search():
    if askyesno("확인", "경로 지정을 진행하시겠습니까?"):
     showwarning("주의 사항", "분석할 xlsx 파일이 존재하는 폴더를 선택해주십시오. \n폴더 내 xlsx파일을 자동으로 인식합니다.")
     main.AzData_pull()

    else: 
     cw.quit
   

def start():
    if askyesno("확인", "데이터 추출을 진행하시겠습니까?"):
     cw_process.start_pro()
     cw_process.start_pro2()
     showinfo("작업 완료", "데이터 추출이 완료되었습니다.")

    else: 
     cw.quit

def quit_cw():
    if askyesno("확인", "chanwoojjangjjangman을 종료하시겠습니까?"):
     cw.destroy()

    else: 
     cw.quit

cw = Tk() 
cw.title('chanwoojjangjjangman') 
cw.geometry("600x470")
label = Label(cw,text='Nmon 데이터 분석기!') 

label.pack()

# 리스트 프레임
list_frame = Frame(cw)
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=15, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)

sys.stdout = TextRedirector(list_file,".")
#for i in range(5):
#    print("현재 i값은", i, "입니다.")


# 저장 경로 프레임
path_frame = LabelFrame(cw, text="저장경로")
path_frame.pack(fill="x", padx=5, pady=5, ipady=5)
txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4 ) # 높이 변경
btn_dest_path = Button(path_frame, text="찾아보기", width=10, command=search)
btn_dest_path.pack(side="right", padx=5, pady=5)


# 진행 상황 Progress Bar
#frame_progress = LabelFrame(cw, text="진행상황")
#frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)
#p_var = DoubleVar()
#progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
#progress_bar.pack(fill="x", padx=5, pady=5)


# determinate --> 프로그램이 동작중인걸 표현할 때 사용하면 좋음, 상태아이콘이 쭉~채워감
def loading():
   progressbar_deter  = ttk.Progressbar(cw, maximum=80, mode="determinate")
   progressbar_deter.start(2) # 10ms 마다 움직임
   progressbar_deter.pack(fill="x", expand=True, padx=5, pady=5, ipady=4)

def loading_out():
   progressbar_deter  = ttk.Progressbar(cw, maximum=80, mode="determinate")
   progressbar_deter.start(0) # 10ms 마다 움직임
   progressbar_deter.pack(fill="x", expand=True, padx=5, pady=5, ipady=6)

label = Label(cw, text='*주의: 동인한 데이터 형식의 xlsx파일만 분석이 가능합니다.')
label.pack()

# 실행 프레임
frame_run = Frame(cw)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=quit_cw)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

cw.resizable(False, False)



cw.mainloop()