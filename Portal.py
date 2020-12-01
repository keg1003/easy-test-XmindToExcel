#!/usr/bin/python
# -*- coding: UTF-8 -*-

from tkinter import *
from src import DictToExcel
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import END
import webbrowser

GUI = tk.Tk()  # 创建窗口对象的背景色
try:
        path = StringVar()
        file = StringVar()

        # 操作信息
        Information = tk.LabelFrame(GUI, text="操作信息", padx=10, pady=10)  # 水平，垂直方向上的边距均为10
        Information.place(x=20, y=240)
        text = Text(Information, width=26, height=3, font=('微软雅黑', 15))
        text.grid()

        #发送指令
        Send = tk.LabelFrame(GUI, text="发送指令", padx=10, pady=5)  # 水平，垂直方向上的边距均为 10
        Send.place(x=20, y=20)

        EntrySend = tk.StringVar()
        Userin = tk.StringVar()

        def WriteData():
            global DataSend
            DataSend = EntrySend.get()
            text.insert("end", '发送指令为：' + str(DataSend) + '\n')
            text.see("end")

        def selectfile():
            path_=askdirectory()
            file.set(path_)

        def selectPath():
            path_1 = askopenfilename()
            path.set(path_1)

        def show_hand_cursor(event):
            text.config(cursor='arrow')


        def show_arrow_cursor(event):
            text.config(cursor='xterm')

        def click(event, x):
            webbrowser.open(x)

        def handlerAdaptor(fun, **kwds):
            return lambda event, fun=fun, kwds=kwds: fun(event, **kwds)

        use = ttk.Entry(Send, textvariable=Userin, width=23)
        use.grid(column=1, row=1)
        ttk.Label(Send, text="输入创建人:").grid(column=0, row=1)
        def csvto():
            global user
            if EntrySend.get()and file.get() and path.get()and Userin.get():
                vban=EntrySend.get()
                file_name = file.get()
                path_name=path.get()
                user = Userin.get()
                DictToExcel.toExcel(file_name, path_name, vban, user)
                text.delete(1.0,END)
                text.insert(END, '转换完成,文件位置：')
                text.tag_config(1, foreground='blue', underline=True)
                text.tag_bind(1, '<Enter>', show_hand_cursor)
                text.tag_bind(1, '<Leave>', show_arrow_cursor)
                text.insert(INSERT, DictToExcel.toExcel(file_name, path_name, vban, user),1)
                text.tag_bind(1, '<Button-1>', handlerAdaptor(click, x=file_name))
            else:
                messagebox.showerror('报错信息','请输入完整信息')

        ttk.Label( Send, text = "输入版本号:" ).grid( column = 0, row = 0 )
        ttk.Label( Send, text = "选择xmind:" ).grid( column = 0, row = 2 )
        ttk.Label( Send, text = "输出用例目录:" ).grid( column = 0, row = 3 )

        Send_Window = ttk.Entry(Send, textvariable=EntrySend, width=23)
        Send_Window.grid(column=1, row=0, pady=5)
        pa=ttk.Entry(Send, textvariable=path, width=23)
        pa.grid(column=1, row=2, pady=5)
        en=ttk.Entry(Send, textvariable=file, width=23)
        en.grid(column=1, row=3, pady=5)

        tk.Button(Send, text="选 择目 录", command=selectfile,width=9).grid(pady=2, sticky=tk.E,column = 2, row = 3)
        tk.Button(Send, text="选择xmind", command=selectPath,width=9).grid(pady=1, sticky=tk.E,column = 2, row = 2)
        tk.Button(Send, text="开始转换", command=csvto).grid(pady=5,column = 1, row = 4,rowspan=2)



except Exception as e:
        text.insert(END, e)

if __name__=='__main__':
    GUI.geometry("400x380")
    GUI.title('xmind转Excel脚本_fzk')
    GUI.mainloop()  # 进入消息循环
