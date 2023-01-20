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

import datetime
import pathlib
from queue import Queue
from threading import Thread

from tkinter import messagebox
from tkinter import ttk ## combobox
import tkinter as tk
from tkinter.filedialog import askdirectory
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap import utility

output_files = 'output_files'
pdf_path = 'E:/pdf/20221119/'
pdf_path = 'E:/pdf/test/'

class FileSearchEngine(ttk.Frame):

    def pdf_to_img(self,pdf_path):
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

            if self.get_file_type(file) != 'pdf':
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
                # print('text=',text)

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

                vals = self.scan_img_get_value(img_file)
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
        self.save_csv(pdf_path, header_list, data_list)

    def scan_img_get_value(self,img_file):
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

    def get_file_type(self,file_path):
        kind = filetype.guess(file_path)
        if kind is None:
            return 'cannot get type'
        else:
            return kind.extension

    def merge_pdf(self,pdf_path):
        pdf = fitz.Document(filetype="pdf")

        files = os.listdir(pdf_path)

        for file_name in sorted(files,reverse=False):
            file = pdf_path+file_name

            if path.isdir(file):  # 如果是目录则退出,继续遍历文件
                continue

            if self.get_file_type(file) != 'pdf':
                continue
            # imagePath = f"./{i}.jpg"
            temp_pdf = fitz.Document(file)
            # imgPdf = fitz.Document("pdf",img.convert_to_pdf())
            pdf.insert_pdf(temp_pdf, 0,)

        pdf.save(pdf_path + output_files +'/invoice_list.pdf')
        pdf.close()

    def save_csv(self,csv_path, header_list, data_list):
        with open(csv_path + output_files + "/invoice_list.csv", mode="w", encoding="utf-8-sig", newline="") as f:
            # 基于打开的文件，创建 csv.DictWriter 实例，将 header 列表作为参数传入。
            writer = csv.writer(f)
            # 写入 header
            writer.writerow(header_list)
            # 写入数据
            writer.writerows(data_list)

    # pdf_to_img(pdf_path)
    # merge_pdf(pdf_path)

# queue 实现了多生产者，多消费队列，这特别适合于消息必须安全在多线程交换的线程编程。模块中的Queue实现了所需的锁定义语义。
    queue = Queue()
    searching = False

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.pack(fill=BOTH, expand=YES)

        # application variables
        _path = pathlib.Path().absolute().as_posix()
        self.path_var = ttk.StringVar(value=_path)
        self.term_var = ttk.StringVar(value='md')
        self.type_var = ttk.StringVar(value='endswidth')

        # header and labelframe option container
        option_text = "Complete the form to begin your search"
        self.option_lf = ttk.Labelframe(self, text=option_text, padding=15)
        self.option_lf.pack(fill=X, expand=YES, anchor=N)

        self.create_path_row()
        self.create_term_row()
        self.create_type_row()
        self.create_results_view()

        self.progressbar = ttk.Progressbar(
            master=self, 
            mode=INDETERMINATE, 
            bootstyle=(STRIPED, SUCCESS)
        )
        self.progressbar.pack(fill=X, expand=YES)

    def create_path_row(self):
        """Add path row to labelframe"""
        path_row = ttk.Frame(self.option_lf)
        path_row.pack(fill=X, expand=YES)
        path_lbl = ttk.Label(path_row, text="Path", width=8)
        path_lbl.pack(side=LEFT, padx=(15, 0))
        path_ent = ttk.Entry(path_row, textvariable=self.path_var)
        path_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttk.Button(
            master=path_row, 
            text="Browse", 
            command=self.on_browse, 
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5)

    def create_term_row(self):
        """Add term row to labelframe"""
        term_row = ttk.Frame(self.option_lf)
        term_row.pack(fill=X, expand=YES, pady=15)
        term_lbl = ttk.Label(term_row, text="Term", width=8)
        term_lbl.pack(side=LEFT, padx=(15, 0))
        term_ent = ttk.Entry(term_row, textvariable=self.term_var)
        term_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        search_btn = ttk.Button(
            master=term_row, 
            text="Search", 
            command=self.on_search, 
            bootstyle=OUTLINE, 
            width=8
        )
        search_btn.pack(side=LEFT, padx=5)

    def create_type_row(self):
        """Add type row to labelframe"""
        type_row = ttk.Frame(self.option_lf)
        type_row.pack(fill=X, expand=YES)
        type_lbl = ttk.Label(type_row, text="Type", width=8)
        type_lbl.pack(side=LEFT, padx=(15, 0))

        contains_opt = ttk.Radiobutton(
            master=type_row, 
            text="Contains", 
            variable=self.type_var, 
            value="contains"
        )
        contains_opt.pack(side=LEFT)

        startswith_opt = ttk.Radiobutton(
            master=type_row, 
            text="StartsWith", 
            variable=self.type_var, 
            value="startswith"
        )
        startswith_opt.pack(side=LEFT, padx=15)

        endswith_opt = ttk.Radiobutton(
            master=type_row, 
            text="EndsWith", 
            variable=self.type_var, 
            value="endswith"
        )
        endswith_opt.pack(side=LEFT)
        endswith_opt.invoke()

    def create_results_view(self):
        """Add result treeview to labelframe"""
        self.resultview = ttk.Treeview(
            master=self, 
            bootstyle=INFO, 
            columns=[0, 1, 2, 3, 4],
            show=HEADINGS
        )
        self.resultview.pack(fill=BOTH, expand=YES, pady=10)

        # setup columns and use `scale_size` to adjust for resolution
        self.resultview.heading(0, text='Name', anchor=W)
        self.resultview.heading(1, text='Modified', anchor=W)
        self.resultview.heading(2, text='Type', anchor=E)
        self.resultview.heading(3, text='Size', anchor=E)
        self.resultview.heading(4, text='Path', anchor=W)
        self.resultview.column(
            column=0, 
            anchor=W, 
            width=utility.scale_size(self, 125), 
            stretch=False
        )
        self.resultview.column(
            column=1, 
            anchor=W, 
            width=utility.scale_size(self, 140), 
            stretch=False
        )
        self.resultview.column(
            column=2, 
            anchor=E, 
            width=utility.scale_size(self, 50), 
            stretch=False
        )
        self.resultview.column(
            column=3, 
            anchor=E, 
            width=utility.scale_size(self, 50), 
            stretch=False
        )
        self.resultview.column(
            column=4, 
            anchor=W, 
            width=utility.scale_size(self, 300)
        )

    def on_browse(self):
        """Callback for directory browse"""
        path = askdirectory(title="Browse directory")
        if path:
            self.path_var.set(path)

    def on_search(self):
        """Search for a term based on the search type"""
        # search_term = self.term_var.get()
        search_path = self.path_var.get()
        print(search_path)

        if search_path == '':
            return

        # start search in another thread to prevent UI from locking
        Thread(
            target=self.pdf_to_img, 
            args=([search_path+'/'])
        ).start()
        self.progressbar.start(10)

        iid = self.resultview.insert(
            parent='', 
            index=END, 
        )
        self.resultview.item(iid, open=True)
        self.after(100, lambda: self.check_queue(iid))

    def check_queue(self, iid):
        """Check file queue and print results if not empty"""
        if all([
            FileSearchEngine.searching, 
            not FileSearchEngine.queue.empty()
        ]):
            filename = FileSearchEngine.queue.get()
            self.insert_row(filename, iid)
            self.update_idletasks()#添加数据后，刷新页面
            self.after(100, lambda: self.check_queue(iid))
        elif all([
            not FileSearchEngine.searching,
            not FileSearchEngine.queue.empty()
        ]):
            while not FileSearchEngine.queue.empty():
                filename = FileSearchEngine.queue.get()
                self.insert_row(filename, iid)
            self.update_idletasks()
            self.progressbar.stop()
        elif all([
            FileSearchEngine.searching,
            FileSearchEngine.queue.empty()
        ]):
            self.after(100, lambda: self.check_queue(iid))## after tkinter 内置函数，在指定豪秒数后执行指定函数
        else:
            self.progressbar.stop()

    def insert_row(self, file, iid):
        """Insert new row in tree search results"""
        try:
            _stats = file.stat()
            _name = file.stem
            _timestamp = datetime.datetime.fromtimestamp(_stats.st_mtime)
            _modified = _timestamp.strftime(r'%m/%d/%Y %I:%M:%S%p')
            _type = file.suffix.lower()
            _size = FileSearchEngine.convert_size(_stats.st_size)
            _path = file.as_posix()
            iid = self.resultview.insert(
                parent='', 
                index=END, 
                values=(_name, _modified, _type, _size, _path)
            )
            self.resultview.selection_set(iid)
            self.resultview.see(iid)
            print('iid=',iid)
        except OSError:
            return

# 把方法转换为静态方法，转换后方法，将不需要通过实例来调用
    @staticmethod
    def file_search(term, search_path, search_type):
        """Recursively search directory for matching files"""
        FileSearchEngine.set_searching(1)
        if search_type == 'contains':
            FileSearchEngine.find_contains(term, search_path)
        elif search_type == 'startswith':
            FileSearchEngine.find_startswith(term, search_path)
        elif search_type == 'endswith':
            FileSearchEngine.find_endswith(term, search_path)

    @staticmethod
    def find_contains(term, search_path):
        """Find all files that contain the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    if term in file:
                        record = pathlib.Path(path) / file
                        print('record=',record)
                        # 将数据添加到队列
                        FileSearchEngine.queue.put(record)
        FileSearchEngine.set_searching(False)

    @staticmethod
    def find_startswith(term, search_path):
        """Find all files that start with the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    if file.startswith(term):
                        record = pathlib.Path(path) / file
                        FileSearchEngine.queue.put(record)
        FileSearchEngine.set_searching(False)

    @staticmethod
    def find_endswith(term, search_path):
        """Find all files that end with the search term"""
        for path, _, files in pathlib.os.walk(search_path):
            if files:
                for file in files:
                    if file.endswith(term):
                        record = pathlib.Path(path) / file
                        FileSearchEngine.queue.put(record)
        FileSearchEngine.set_searching(False)

    @staticmethod
    def set_searching(state=False):
        """Set searching status"""
        FileSearchEngine.searching = state

    @staticmethod
    def convert_size(size):
        """Convert bytes to mb or kb depending on scale"""
        kb = size // 1000
        mb = round(kb / 1000, 1)
        if kb > 1000:
            return f'{mb:,.1f} MB'
        else:
            return f'{kb:,d} KB' 

def win_app():
    app = ttk.Window("File Search Engine", "journal")
    FileSearchEngine(app)
    app.mainloop()

win_app()