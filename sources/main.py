from icon import img
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import base64
import os
import sys
import xlrd
import xlwt
import csv
import webbrowser

'''
{'mode': 0, 'files': ['数据2.xls'], 'names': ['数据2'], 'open_directory': 'C:/Users/Microsoft/Desktop', 'open_filenames': ('C:/Users/Microsoft/Desktop/数据2.xls',), 'save_directory': 'C:/Users/Microsoft/Desktop', 'ready': 'yes', 'result': '', 'file_names': 1}
'''

# 获取参数，设置工作模式
def setMode(data,mode):
    data['mode']=mode
    clean(data)
    return data


# 打开文件名获取路径和去后缀的文件名
def openFiles(data):
    if data['mode']==0:
        file_names=filedialog.askopenfilenames(filetypes=[("EXCEL",("xls","xlsx"))],title="选择EXCEL",initialdir=data['open_directory'])
    elif data['mode']==1:
        file_names=filedialog.askopenfilenames(filetypes=[("CSV","csv")],title="选择CSV",initialdir=data['open_directory'])
    else:
        pass
    if file_names:
        tk_opened_file_names.set('')
        result_info.set('')
        data['files']=[]
        data['names']=[]
        for item in file_names:
            data['files'].append(os.path.basename(item))
            data['names'].append(os.path.splitext(os.path.basename(item))[0])
        tk_opened_file_names.set(','.join(data['files']))
        data['open_directory']=os.path.dirname(file_names[0])
        data['save_directory']=os.path.dirname(file_names[0])
        data['open_filenames']=file_names
        data['file_names']=len(file_names)
    check(data)
    return data


# 获取保存路径
def saveDirectory(data):
    save_directory=filedialog.askdirectory()
    if save_directory:
        tk_save_directory.set('')
        result_info.set('')
        data['save_directory']=save_directory
        tk_save_directory.set(data['save_directory'])
    check(data)
    return data

# 检查是否满足转换条件
def check(data):
    if len(tk_opened_file_names.get())>0 and len(tk_save_directory.get())>0:
        data['ready']='yes'
        tk_start_button.config(state=tk.NORMAL)
    return data
# 获取保存路径名和去后缀的文件名检查是否会有文件覆盖，有则给提示，没有则转开始转换
def starcheck(data):
    list=[]
    if data['mode']==0:
        for file_names in data['names']:
            if os.path.isfile(os.path.join(data['save_directory'],file_names+'.csv')):
                list.append(file_names+'.csv')
        info(list)
    elif data['mode']==1:
        for file_names in data['names']:
            if os.path.isfile(os.path.join(data['save_directory'],file_names+'.xls')):
                list.append(file_names+'.xls')
        info(list)
def info(files):
    if len(files)>1:
        files_check_result=messagebox.askyesno(title='文件已存在',message=files[0]+' 等'+str(len(files))+'个文件已存在\n是否继续？',default='yes')
    elif len(files)==1:
        files_check_result=messagebox.askyesno(title='文件已存在',message=files[0]+' 文件已存在\n是否继续？',default='yes')
    else:
        convert(data)
        return None
    if files_check_result:
        # 转换
        convert(data)
    else:
        # 程序没任何动静
        pass
# 获取完整路径文件名和去后缀的文件名，开始文件转换文件
def convert(data):
    if data['mode']==0:
        for key,excel_file in enumerate(data['open_filenames']):
            with open(os.path.join(data['save_directory'],data['names'][key]+'.csv'), 'w', newline='', encoding='utf-8') as csvfile:
                csv_data = csv.writer(csvfile)
                excel = xlrd.open_workbook(excel_file)
                table = excel.sheet_by_index(0)
                nrow = table.nrows
                for i in range(0, nrow):
                    # print(table.row_values(i))
                    csv_data.writerow(table.row_values(i))
        #展示结果
    elif data['mode']==1:
        for key,csv_file in enumerate(data['open_filenames']):
            with open(csv_file, "r", encoding="utf-8") as csvfile:
                csv_data=csv.reader(csvfile)
                book = xlwt.Workbook()
                sheet = book.add_sheet('Sheet1')
                for i, row in enumerate(csv_data):
                    for j, col in enumerate(row):
                        sheet.write(i, j, col)
                book.save(os.path.join(data['save_directory'],data['names'][key]+'.xls'))
    showConvertResult()
# 返回结果给界面
def showConvertResult():
    string_info=str(len(data['names']))+' 个文件转换完'
    result_info.set(string_info)
# 回调测试
def callback(param):
    print(param)
# 清理记录
def clean(data):
    tk_opened_file_names.set('')
    tk_save_directory.set('')
    result_info.set('')
    tk_start_button.config(state=tk.DISABLED)
    data['files']=[]
    data['names']=[]
    data['open_directory']=''
    data['open_filenames']=[]
    data['save_directory']=''
    data['ready']='no'
    data['result']=''
    return data
#开打浏览器
def openWeb():
    webbrowser.open_new('https://github.com/ssyatelandisi/Excel-to-CSV')
# GUI界面
tmp = open("tmp.ico","wb+")
tmp.write(base64.b64decode(img))
tmp.close()
# def main()
data={'mode':0,'files':[],'names':[],'open_directory':'','open_filenames':[],'save_directory':'','ready':'no','result':''}
root=tk.Tk()
root.title("EXCEL CSV互转工具")
menubar=tk.Menu(root)
root.geometry("500x310")
root.resizable(width=False, height=False)
#ico要填写绝对路劲
root.iconbitmap("tmp.ico")
os.remove("tmp.ico")
file_menu = tk.Menu(menubar, tearoff=False)
mode=tk.IntVar(value=0)
file_menu.add_radiobutton(label="EXCEL转CSV", variable=mode, value=0,command=lambda param=data,mode=0:setMode(data,mode))
file_menu.add_radiobutton(label="CSV转EXCEL", variable=mode, value=1,command=lambda param=data,mode=1:setMode(data,mode))
file_menu.add_separator()
file_menu.add_command(label="退出", command=root.quit)
menubar.add_cascade(label="文件", menu=file_menu)
help_menu = tk.Menu(menubar, tearoff=False)
# help_menu.add_command(label="查看帮助")
help_menu.add_command(label="关于软件",command=openWeb)
menubar.add_cascade(label="帮助", menu=help_menu)
root.config(menu=menubar)
frame_choose=tk.Frame(root,pady=5)
tk.Radiobutton(frame_choose,text="EXCEL 转 CSV",variable=mode,value=0,command=lambda param=data,mode=0:setMode(data,mode)).grid(row=0,column=0)
tk.Radiobutton(frame_choose,text="CSV 转 EXCEL",variable=mode,value=1,command=lambda param=data,mode=1:setMode(data,mode)).grid(row=0,column=1)
frame_choose.pack()
frame_main=tk.LabelFrame(root,padx=5, pady=15)
tk.Label(frame_main,text="EXCEL 转 CSV").grid(row=0,column=0)
tk.Label(frame_main,text="打开文件").grid(row=1,column=0)
tk_opened_file_names=tk.StringVar()
tk.Entry(frame_main,width=40,textvariable=tk_opened_file_names).grid(row=1,column=1,padx=10, pady=5)
tk.Button(frame_main,text="浏览", width=6,command=lambda data=data:openFiles(data)).grid(row=1,column=2)
tk.Label(frame_main,text="保存路径").grid(row=2,column=0)
tk_save_directory=tk.StringVar()
tk.Entry(frame_main,width=40,textvariable=tk_save_directory).grid(row=2,column=1, padx=10, pady=5)
tk.Button(frame_main,text="浏览", width=6,command=lambda:saveDirectory(data)).grid(row=2,column=2)
frame_main.pack()
frame_button=tk.Frame(root, pady=5)
tk_start_button=tk.Button(frame_button,text="开始转换",command=lambda param=data:starcheck(param))
tk_start_button.config(state=tk.DISABLED)
tk_start_button.grid(row=0,column=0, padx=5, pady=5)
tk.Button(frame_button,text="清空内容",command=lambda:clean(data)).grid(row=0,column=1, padx=5, pady=5)
frame_button.pack()
result_info=tk.StringVar()
tk.Label(root,textvariable=result_info,pady=5).pack()
root.mainloop()

# # 主函数执行
# main()