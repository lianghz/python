# pip install pymupdf
# pip install Pillow
# pytesseract 与 easyocr 都是 ocr ，看看谁的识别率高（easyocr 没成功，所以使用 pytesseract）
# pip install pytesseract
# pip install easyocr

# pip install filetype
# https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.2.0.20220712.exe
# C:\Users\Administrator\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\LocalCache\local-packages\Python310\site-packages\pytesseract\pytesseract.py

# 识别二维码
# pip install opencv-python
# pip install pyzbar

# 使用 ttkbootstrap 做窗口应用样式
# python -m pip install ttkbootstrap

from re import I
from tkinter import messagebox
from tkinter import ttk ## combobox
import tkinter as tk
import fitz
import os
from os import path
# from PIL import Image
from PIL import ImageEnhance
import PIL
import pytesseract
# import easyocr
import cv2
import os
from pyzbar import pyzbar
from tkinter import *
import filetype
import csv  # 内置
import re

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

output_files = 'output_files'
pdf_path = 'E:/pdf/20221119/'
pdf_path = 'E:/pdf/test/'

def pdf_to_img(pdf_path):
    img_path = pdf_path + output_files + '/'
    img_path = img_path.replace('//', '/')

    files = os.listdir(pdf_path)

    zoom_x = 2.0  # 1.0=624*397
    zoom_y = 2.0

    mtx = fitz.Matrix(zoom_x, zoom_y)
    data_list = []
    for file_name in files:
        file = pdf_path+file_name
        # print('--------------------------------------')
        # print('file=', file)
        if path.isdir(file):  # 如果是目录则退出,继续遍历文件
            continue

        if get_file_type(file) != 'pdf':
            continue

        pdf = fitz.open(file)

        page_counts = pdf.page_count
        # print('总页数',page_counts)
        dts_file_pdf = ''

        for page_num in range(page_counts):
            # print('正在转换第',page_num,'页')
            page = pdf[page_num]

            # text = page.get_text()
            # print('text = ' , text)

            pix = page.get_pixmap(matrix=mtx)
            if not os.path.exists(img_path):
                os.makedirs(img_path)
            img_file = img_path + file_name[:-4] + '-' + str(page_num) + '.jpg'
            pix.save(img_file)

            image = PIL.Image.open(img_file)
            # #调整图片灰度为黑白
            # image.convert('RGB')
            # enhancer = ImageEnhance.Color(image)
            # enhancer = enhancer.enhance(0)
            # enhancer = ImageEnhance.Brightness(enhancer)
            # enhancer = enhancer.enhance(2)
            # enhancer = ImageEnhance.Contrast(enhancer)
            # enhancer = enhancer.enhance(8)
            # enhancer = ImageEnhance.Sharpness(enhancer)
            # image = enhancer.enhance(20)

            # print('img_file = ',img_file)
            text = pytesseract.image_to_string(image,'chi_sim')
            # reader = easyocr.Reader(['ch_sim','en']) 
            # text = reader.readtext('E:/pdf/test/output_files/1.jpg')
            print('text=',text)
            amt_str = text.split('写')[-1]
            amt_str = amt_str[2:7].replace(' ','')
            amt_list = re.findall(r'-?\d+\.?\d*', amt_str)
            amt = ''
            if len(amt_list) >0:
                amt = amt_list[0]
            # print('amt=',amt)
            company_name = text.split('称')[-1]
            company_idx = company_name.find('店')+1
            if company_idx == 0:
                company_idx = company_name.find('公司')+2
            
            company_name = company_name[1:company_idx].replace(' ','')

            vals = scan_img_get_value(img_file)
            format_file_name = vals[5]+'_' + vals[2]+'_'+vals[3]+'_'+vals[4]
            dts_file_jpg = img_path + format_file_name +'.jpg'
            dts_file_pdf = pdf_path + format_file_name +'.pdf'

            data_list.append([vals[5][:4]+'/'+vals[5][4:6]+'/'+vals[5][6:8], vals[2], vals[3], vals[4], amt,company_name,dts_file_pdf])
            # print(img_file)
            # print(dstFile)
            if os.path.exists(dts_file_jpg):
                os.remove(dts_file_jpg)
            os.rename(img_file, dts_file_jpg)

        pdf.close()
        os.rename(file, dts_file_pdf)
        

    print('转换完成！')
    header_list = ['日期', '发票代码', '发票号码', '金额(不含税)', '金额','销售方','文件名']
    save_csv(pdf_path, header_list, data_list)

def scan_img_get_value(img_file):
    # qrcode = cv2.imread(img_file)
    # barcode = pyzbar.decode(qrcode)
    img = cv2.imread(img_file)
    # det = cv2.QRCodeDetector()
    # barcode = det.detectAndDecode(img)
    #vals, pts, st_code = barcode
    #cv2 的二维码读取，不能对所有二维码读取，改用pyzbar比较好
    barcode = str(pyzbar.decode(img)[0].data)
    # print('barcode = ', barcode)
    # barcode =  01,10,044002100611,52573933,200.75,20221004,11969205383964570811,F5F6
    # 2：发票代码
    # 3：发票号码
    # 4：金额
    # 5：发票日期
    # 6：校验码
    return barcode.split(',')

def get_file_type(file_path):
    kind = filetype.guess(file_path)
    if kind is None:
        return 'cannot get type'
    else:
        return kind.extension

def merge_pdf(pdf_path):
    pdf = fitz.Document(filetype="pdf")

    files = os.listdir(pdf_path)

    for file_name in sorted(files,reverse=False):
        file = pdf_path+file_name

        if path.isdir(file):  # 如果是目录则退出,继续遍历文件
            continue

        if get_file_type(file) != 'pdf':
            continue
        # imagePath = f"./{i}.jpg"
        temp_pdf = fitz.Document(file)
        # imgPdf = fitz.Document("pdf",img.convert_to_pdf())
        pdf.insert_pdf(temp_pdf, 0,)

    pdf.save(pdf_path + output_files +'/invoice_list.pdf')
    pdf.close()

def save_csv(csv_path, header_list, data_list):
    with open(csv_path + output_files + "/invoice_list.csv", mode="w", encoding="utf-8-sig", newline="") as f:
        # 基于打开的文件，创建 csv.DictWriter 实例，将 header 列表作为参数传入。
        writer = csv.writer(f)
        # 写入 header
        writer.writerow(header_list)
        # 写入数据
        writer.writerows(data_list)

# pdf_to_img(pdf_path)
# merge_pdf(pdf_path)


class Application(tk.Frame):##继承Tk.Frame
    def __init__(self,root):#root 为传入的主窗体（最外层）
        super().__init__(root)##则self 继承了TK.Frame 的方法和属性
        # self.master = root
        self.pack()
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.select_type = tk.StringVar()
        root.geometry('300x200')

        self.init_win()
    def on_click_login(self):
        print('点击了登录按钮')
        username = self.username.get()
        password = self.password.get()
        select_type = self.selector.get()
        messagebox.showinfo('提示：','用户名:%s\n密码：%s\n身份：%s' %(username,password,select_type))

    def select_change(self,*args):
        self.select_type=self.selector.get()

    def init_win_bk(self):
        print('win_app_start')
        # 使用pack 来布局
        # Label(self,text='123',fg='red',bg='green',width=1).pack()
        # Entry(self,textvariable=self.envar).pack()
        # 使用grid来布局 设置所在的行和列，以及span 等
        Label(self,text='123',fg='red',bg='green',width=1).grid(row=0,column=0)
        Entry(self,textvariable=self.envar).grid(row=0,column=1)
        text = Text(self)
        text.grid(row=1,columnspan=2)
        text.insert('0.1','123')
        # text.insert('1.1','123333')

    def init_win(self):
        frame1 = Frame(self)
        Label(frame1,text='账号：').grid(row=0,column=0)
        Entry(frame1,textvariable=self.username).grid(row=0,column=1)

        frame2 = Frame(self)
        Label(frame2,text='密码：').grid(row=0,column=0)
        Entry(frame2,show='*',textvariable=self.password).grid(row=0,column=1)

        frame3 = Frame(self)
        Label(frame3,text='身份：').grid(row=0,column=0)
        self.selector =  ttk.Combobox(frame3,values=('管理员','普通成员'),width=18)
        self.selector.grid(row=0,column=2)
        self.selector.current(0)
        self.selector.bind('<<ComboboxSelected>>',self.select_change)#演示事件绑定

        frame3.grid(pady=6)
        frame1.grid(pady=10)#pady Y 轴的与上一对象的像素间隔
        frame2.grid(pady=6)
       

        ttk.Button(self,text='登录',width=15,command=self.on_click_login, bootstyle=SUCCESS).grid(pady=10)

def win_app():
    print('win_app1')

    # app = Tk()  # 建立一个窗口叫 app
    app = Application(Tk())
    app.mainloop()

    # app.title(string='PDF 发票转图片合并')
    # app.geometry('1000x500')

    # btn1 = Button(app)  # 创建一个按钮，传入参数app，放在窗口app里
    # btn1['text'] = '按钮'
    # btn1.pack()  # 调用布局管理器，把按钮合理的放到窗口上

    # # def click1(e):  # e是事件对象
    # #     messagebox.showinfo('Message', '弹出信息')

    # def click1():  # e是事件对象
    #     messagebox.showinfo('Message', '弹出信息')

    # # btn1.bind('<Button>', click1)
    # btn1['command'] = click1

    # app.mainloop()  # 调用该方法，进入事件循环，用于监听用户事件

win_app()