import tkinter as tk
import tkinter.ttk as ttk
import time

root = tk.Tk()  # tkinter root창 생성

root.title("tkinter 공부") #창 이름
root.geometry("500x500+200+200") # 창 크기, 가로 x 세로 + 창 출력 위치 좌표

# indeterminate --> 프로그램이 동작중인걸 표현할 때 사용하면 좋음, 상태아이콘이 왔다갔다
progressbar_indeter = ttk.Progressbar(root, maximum=100, mode="indeterminate")
progressbar_indeter.start(10) # 10ms 마다 움직임
progressbar_indeter.pack()

# determinate --> 프로그램이 동작중인걸 표현할 때 사용하면 좋음, 상태아이콘이 쭉~채워감
progressbar_deter  = ttk.Progressbar(root, maximum=80, mode="determinate")
progressbar_deter.start(5) # 10ms 마다 움직임
progressbar_deter.pack()

# 실제 사용자가 원하는 스타일(얼마나 진행되고 있는지 표현)
p_var2= tk.DoubleVar() # 정수, 실수도 사용하기 위해
progressbar_status = ttk.Progressbar(root, maximum=100, length=50, variable=p_var2)
progressbar_status.pack()

def btncmd_indeter():
    progressbar_indeter.stop() # 작동 중지

def btncmd_deter():
    progressbar_deter.stop() # 작동 중지

def btncmd_status():
    for i in range(1, 100): #progressbar maximum 100으로 설정했으므로
        time.sleep(0.01)

        p_var2.set(i)
        progressbar_status.update() # 프로그래스바 상태아이콘 반영
        print(p_var2.get())

btn = tk.Button(root, text="indeter 정지", command=btncmd_indeter)
btn.pack()

btn = tk.Button(root, text="deter 정지", command=btncmd_deter)
btn.pack()

btn = tk.Button(root, text="status 시작", command=btncmd_status)
btn.pack()

root.mainloop()