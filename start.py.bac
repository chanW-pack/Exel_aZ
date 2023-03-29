import main
from tkinter import *
import tkinter.ttk as ttk
from tkinter.messagebox import *

cw = Tk() 
cw.title('chanwoojjangjjangman') 
cw.geometry("600x470")
label = Label(cw,text='Nmon 데이터 분석기!') 

label.pack()


def search():
        main.AzData_pull()

def start():
    if askyesno("확인", "정말 진행하시겠습니까?"):
     #main.start_pro()
     main.AzData_pull()
    else:
     cw.quit
        
# 리스트 프레임
list_frame = Frame(cw)
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=15, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)




# 저장 경로 프레임
path_frame = LabelFrame(cw, text="저장경로")
path_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path = Button(path_frame, text="찾아보기", width=10, command=search)
btn_dest_path.pack(side="right", padx=5, pady=5)


# 진행 상황 Progress Bar
frame_progress = LabelFrame(cw, text="진행상황")
frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)

p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x", padx=5, pady=5)

# 실행 프레임
frame_run = Frame(cw)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=cw.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

cw.resizable(False, False)



cw.mainloop()


# main 함수 실행
#main.start_pro()
