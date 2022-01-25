import logging
from tkinter import *
import tkinter.filedialog
from excel import (
    ExcelOpt,ExcelOptNew  
)

class WidgetLogger(logging.Handler):
    def __init__(self, widget):
        logging.Handler.__init__(self)
        self.setLevel(logging.INFO)
        self.widget = widget
        self.widget.config(state='disabled')

    def emit(self, record):
        self.widget.config(state='normal')
        # Append message (record) to the widget
        self.widget.insert(tkinter.END, self.format(record) + '\n')
        self.widget.see(tkinter.END)  # Scroll to the bottom
        self.widget.config(state='disabled')

# 主程序
root = Tk()
# 设置标题
root.title("表格处理")
# 设置主窗口大小
root.geometry("640x480")
# 可变大小
root.resizable(width=True, height=True)

log_data_Text = Text(root, width=66, height=9) # 日志框
log_data_Text.grid(row=13, column=0, columnspan=10)
logger = WidgetLogger(log_data_Text)

# 第一排输入框 输入查询的内容
# 左边是一个标签
l1 = Label(root, text='文件路径', bg="yellow", font=(12), height=1, width=8)
l1.place(x=20, y=20)
var1 = StringVar()
input_text = Entry(root, textvariable=var1,width=40)
input_text.place(x=100, y=20)

# 第二排显示框 显示查询的结果
# 左边是一个标签
l2 = Label(root, text='查询结果', bg="yellow", font=(12), height=1, width=8)
l2.place(x=20, y=60)
var2 = StringVar()
output_text = Entry(root, textvariable=var2,width=40)
output_text.place(x=100, y=60)

def select():
    filename = tkinter.filedialog.askopenfilename()
    if len(filename) != 0:
        var1.set(filename)
    else:
        var1.set("")

btn = Button(root,text="选择文件",command=select)
btn.place(x=400, y=20)

# 创建列表框
list_itmes_select = [
    '1.提取出办事处、客户编码、订单量表（old）',
    '2.提取费用网点明细表（old）', 
    '3.提取出办事处、客户编码、订单量表（new）', 
    '4.提取费用网点明细表（new）']
list_itmes = StringVar()

list_itmes.set((
    '1.提取出办事处、客户编码、订单量表（old）',
    '2.提取费用网点明细表（old）', 
    '3.提取出办事处、客户编码、订单量表（new）', 
    '4.提取费用网点明细表（new）'))  # 设置可选项
listB = tkinter.Listbox(root, listvariable = list_itmes,width=40)
listB.place(x=100, y=80)

no_select = list_itmes_select[0]
var2.set(no_select)

def click_button():
    """
    当按钮被点击时执行该函数
    :return:
    """
    selectN = listB.curselection()
    print(len(selectN))
    if len(selectN) == 0:
        var2.set(no_select)
    else:
        text = listB.get(selectN)
        var2.set(text)

btn2 = Button(root,text="选择类型",command=click_button)
btn2.place(x=400, y=60)
def optExcel():
    print("var2： ",var1.get(),var2.get(),list_itmes_select[0])   
    if(var2.get() == list_itmes_select[0]):
        opt = ExcelOpt()
        opt.extractData(var1.get())
    elif(var2.get() == list_itmes_select[1]):
        opt = ExcelOpt()
        opt.extractDataDetail(var1.get())
    elif(var2.get() == list_itmes_select[2]):
        optNew = ExcelOptNew()
        optNew.extractDataDetail(var1.get())
    elif(var2.get() == list_itmes_select[3]):
        optNew = ExcelOptNew()
        optNew.extractDataDetail(var1.get())


btn3 = Button(root,text="执行",command=optExcel)
btn3.place(x=100, y=300)

# 运行主程序
root.mainloop()
