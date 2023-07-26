import time
import pandas as pd
from pandas_datareader import data as web
import yfinance as yahoo
import tkinter as tk
from tkinter import ttk
import threading
yahoo.pdr_override()
##############################################################################
root = tk.Tk()
root.geometry("1600x750")
#請輸入股票四碼加後綴地區碼
stock_frame = tk.Frame(root)
stock_frame.pack(side=tk.TOP)
stock_label = tk.Label(stock_frame, text='股票四碼加後綴地區碼')
stock_label.pack(side=tk.LEFT)
stock_entry = tk.Entry(stock_frame)
stock_entry.pack(side=tk.LEFT)

#股票四碼加後綴地區碼範例
exid_frame = tk.Frame(root)
exid_frame.pack(side=tk.TOP)
exid_label = tk.Label(exid_frame, text='範例：1802.tw')
exid_label.pack(side=tk.LEFT)

#起
startTime_frame = tk.Frame(root)
startTime_frame.pack(side=tk.TOP)
startTime_label = tk.Label(startTime_frame, text='起：')
startTime_label.pack(side=tk.LEFT)
input_startdate_var = tk.StringVar()
startdate_widget = tk.Entry(startTime_frame,textvariable = input_startdate_var)
input_startdate_get = input_startdate_var.set(input_startdate_var.get())
startdate_widget.pack(side = tk.LEFT)

#起 範例
exstart_frame = tk.Frame(root)
exstart_frame.pack(side=tk.TOP)
exstart_label = tk.Label(exstart_frame, text='範例：xxxx-xx-xx')
exstart_label.pack(side=tk.LEFT)

#訖
finishTime_frame = tk.Frame(root)
finishTime_frame.pack(side=tk.TOP)
finishTime_label = tk.Label(finishTime_frame, text='訖：')
finishTime_label.pack(side=tk.LEFT)
input_finishdate_var = tk.StringVar()
finishdate_widget = tk.Entry(finishTime_frame,textvariable = input_finishdate_var)
input_finishdate_get = input_finishdate_var.set(input_finishdate_var.get())
finishdate_widget.pack(side = tk.LEFT)

#訖 範例
exfinish_frame = tk.Frame(root)
exfinish_frame.pack(side=tk.TOP)
exfinish_label = tk.Label(exfinish_frame, text='範例：xxxx-xx-xx')
exfinish_label.pack(side=tk.LEFT)
def run():
 stock = str(stock_entry.get())
 start_date = input_startdate_var.get()
 finish_date = input_finishdate_var.get()
 df = web.get_data_yahoo([stock],start_date,finish_date)
 file = 'temp.xlsx'
 save = df.to_excel(file)
 load = pd.read_excel(file)
 cols = list(load.columns)
 tree = ttk.Treeview(root)
 tree.pack()
 tree["columns"] = cols
 for i in cols:
  tree.column(i,anchor="center")
  tree.heading(i,text=i,anchor='center')
 for index, row in load.iterrows():
  tree.insert("",'end',text = index,values=list(row))
 tree.place(relx=0,rely=0.22,relheight=0.7,relwidth=1)
onbutton = tk.Button(finishTime_frame, text = "查詢", command = run).pack()     
root.mainloop()