#打开文件夹，指定图片大小（按像素) ,指定存放路径，指定生成的文件名
#-*- coding:utf-8 -*-
from tkinter import *
from tkinter import messagebox
import tkinter as tk#要使用，先导入
import easygui
import time
from openpyxl.drawing.image import Image
from openpyxl import Workbook
import os

from PIL import JpegImagePlugin
JpegImagePlugin._getmp = lambda  x:None

selected_dict={'file_path':""} #用来存储函数中需返回的值 选定的文件夹
result_dict = {'file_path':""} #用来存储函数中需返回的值 选定的目标文件夹

def convert_path(path: str) -> str:
    return path.replace(r'\/'.replace(os.sep, ''), os.sep)
    
def selectfilepath():
     sPath = easygui.diropenbox()    
     selected_dict['file_path'] = sPath
     label_2 = tk.Label(frame1, font=('雅黑',10, 'bold '),bg='#66CDAA',justify='left',text="源文件选定的文件夹:"+selected_dict['file_path'] )
     label_2.place(relx=0, rely=0.6)
     
def resultfilepath():
     sPath = easygui.diropenbox()    
     result_dict['file_path'] = sPath  
     label_2 = tk.Label(frame3, font=('雅黑',10, 'bold '),bg='#99CDAA',justify='left',text="目标文件名"+result_dict['file_path'])
     label_2.place(relx=0, rely=0.6)
def quitapp():
     sys.exit() 
def startrun():
    #获取源文件夹名称  目标文件夹
    src_path=selected_dict['file_path']
    dest_path=result_dict['file_path']
    pic_w=(entry1.get())
    pic_h=(entry2.get())
    dest_filename=dest_path+'\图片导入'+str(time.time())+'.xlsx'
    
    wb = Workbook()
    ws = wb.active
    #初始化表头
    ws.cell(row=1, column=1).value ='sn'
    ws.cell(row=1, column=2).value ='文件名'
    ws.cell(row=1, column=3).value ='图片'
    
    file_name_list = os.listdir(src_path)

    #计数
    count=0
    for srcfile in file_name_list:     
        count=count+1
    
    sn=2         
    for srcfile in file_name_list:     
        #插入三列 第一列sn 第二列 文件名 低三列图片
        try:     
            image_file=convert_path(src_path+'\\'+srcfile)
            img = Image(image_file)
            
            image_column='C'  
            img.width, img.height = (int(pic_w), int(pic_h))
            ws.column_dimensions[image_column].width = int(pic_w)*0.15
            ws.row_dimensions[sn].height = int(pic_h)*0.8
            ws.cell(row=sn, column=1).value =sn-1
            ws.cell(row=sn, column=2).value =srcfile
            ws.add_image(img, anchor=image_column + str(sn)) 
            text1.insert(INSERT,'进度：'+str(sn-1)+' / '+str(count)+' 。 插入图片：'+srcfile+' 成功!\n')      
            text1.see("end")  
            sn=sn+1
            win.update()
        except:
            sn=sn+1
            text1.insert(INSERT,'进度：'+str(sn-1)+' / '+str(count)+image_file+'不是图片文件！\n')    
            continue
        
    wb.save(dest_filename)
    wb.close()
    text1.insert(INSERT,'程序运行结束!\n')
    messagebox.showinfo('运行成功','妹子你真好看！')
    return     

if __name__ == '__main__':     
    win = tk.Tk()#创建一个窗口，因为后面还要用到所以用window这个变量来赋值，可以自行更改
    sw = win.winfo_screenwidth()
    #得到屏幕宽度
    sh = win.winfo_screenheight()
    #得到屏幕高度
    ww = 500
    wh = 450
    #窗口宽高为500
    x = (sw-ww) / 2
    y = (sh-wh) / 2
    win.geometry("%dx%d+%d+%d" %(ww,wh,x,y))
    win.title("妹子批量插图专用！")
    win.resizable(False, False)
    # 定义第一个容器，使用 labelanchor ='w' 来设置标题的方位
    frame1 = tk.LabelFrame(win, text="1", labelanchor="w",bg='#66CDAA')
    frame1.place(relx=0, rely=0, relwidth=1, relheight=0.2)
    # 使用 place 控制 LabelFrame 的位置
    label_1 = tk.Label(frame1, font=('雅黑',12, 'bold '),bg='#66CDAA',justify='left',text="选择需要插入图片的文件夹")
    label_1.place(relx=0, rely=0.2)
    b = tk.Button(frame1, text="照片存放文件夹", justify='center',command=selectfilepath)
    b.place(relx=0.5, rely=0.15)
    
    # 定义第二个容器，使用 labelanchor ='w' 来设置标题的方位
    frame2 = tk.LabelFrame(win, text="2", labelanchor="w",bg='#88CDAA')
    # 使用 place 控制 LabelFrame 的位置
    frame2.place(relx=0, rely=0.2, relwidth=1, relheight=0.2)
    
    label_1 = tk.Label(frame2, font=('雅黑',10, 'bold '),bg='#88CDAA',justify='left',text="输入插入图片的宽度")
    label_1.place(relx=0, rely=0.15)
    
    label_2 = tk.Label(frame2, font=('雅黑',10, 'bold '),bg='#88CDAA',justify='left',text="输入插入图片的高度")
    label_2.place(relx=0, rely=0.6)
    
    entry1 = tk.Entry(frame2)
    entry1.place(relx=0.5, rely=0.15)
    entry1.delete(0, "end")
    entry1.insert(0,'150')
    
    entry2 = tk.Entry(frame2)
    entry2.place(relx=0.5, rely=0.6)
    entry2.delete(0, "end")
    entry2.insert(0,'150')
    
    # 定义第三个容器，使用 labelanchor ='w' 来设置标题的方位
    frame3 = tk.LabelFrame(win, text="3", labelanchor="w",bg='#99CDAA')
    # 使用 place 控制 LabelFrame 的位置
    frame3.place(relx=0, rely=0.4, relwidth=1, relheight=0.2)
    
    label_1 = tk.Label(frame3, font=('雅黑',10, 'bold '),bg='#99CDAA',justify='left',text="目标文件夹")
    label_1.place(relx=0, rely=0.2)
    b = tk.Button(frame3, text="照片存放文件夹", justify='center',command=resultfilepath)
    b.place(relx=0.5, rely=0.15)
    
    # 定义第四个容器，使用 labelanchor ='w' 来设置标题的方位
    frame4 = tk.LabelFrame(win, text="4", labelanchor="w",bg='#AACDAA')
    # 使用 place 控制 LabelFrame 的位置
    frame4.place(relx=0, rely=0.6, relwidth=1, relheight=0.3)
    
    label_1 = tk.Label(frame4, font=('雅黑',10, 'bold '),bg='#AACDAA',justify='left',text="输出结果")
    label_1.place(relx=0, rely=0.2)
    
    text1 = tk.Text(frame4, width=55, height=9, undo=True, autoseparators=True)
    text1.place(relx=0.15, rely=0)
    
    scroll = tk.Scrollbar()
    # 放到窗口的右侧, 填充Y竖直方向
    scroll.pack(side=tk.RIGHT,fill=tk.Y)
    # 两个控件关联
    scroll.config(command=text1.yview)
    text1.config(yscrollcommand=scroll.set)
    
    # 定义第5个容器，使用 labelanchor ='w' 来设置标题的方位
    frame5 = tk.LabelFrame(win, text="5", labelanchor="w",bg='#BBCDAA')
    # 使用 place 控制 LabelFrame 的位置
    frame5.place(relx=0, rely=0.9, relwidth=1, relheight=0.1)
    #开始插入
    b2 = tk.Button(frame5, text="开始执行", justify='center',command=startrun)
    b2.place(relx=0.3, rely=0.1)
    #结束运行
    b3= tk.Button(frame5, text="退出程序", justify='center',command=quitapp)
    b3.place(relx=0.6, rely=0.1)
    win.mainloop()

