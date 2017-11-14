# -*- coding:utf-8 -*-
import time
import tkinter
import tkinter.filedialog
import tkinter.filedialog
import tkinter.messagebox
import tkinter.messagebox
from tkinter import *

import requests
import urllib3
import xlrd
import xlwt
from selenium import webdriver

from config.config import *

requests.packages.urllib3.disable_warnings()

browser = webdriver.PhantomJS(executable_path=DRIVER_PATH)
excel = xlrd.open_workbook(EXCEL)
sheet1 = excel.sheet_by_name(u'Sheet1')
http_client = urllib3.PoolManager()
print("总共有%d个数据" % sheet1.nrows)

# 图片列表
picture_list = []
# 商品详情列表
details_info_list = []
# 颜色 尺码 价格 子字典
color_size_list = []

lis = [{'id': 2, 'name': '2017春季新款漆皮系带厚底防水台松糕休闲女鞋潮', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39', 'price': '26.00'},
       {'id': 3, 'name': '新款英伦风女鞋2017秋冬季单鞋休闲百搭粗跟中跟韩版皮鞋厂家直销', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39',
        'price': '27.00'},
       {'id': 4, 'name': '2017秋季新款厚底女鞋坡跟休闲女单鞋深口系带圆头舒适松糕鞋女', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39',
        'price': '28.00'}]


def get_excel_data():
    browser.get("https://detail.1688.com/")
    # 表格的URL在第六列
    for i in range(1, sheet1.nrows - 1):
        row = sheet1.row_values(i)
        print("正在获取连接:", row[6])
        # 开始解析第六列的数据

        try:
            browser.get(row[6])
            # 获取商品名称
            title = browser.find_element_by_class_name("d-title").text
            # 获取价格
            price = browser.find_element_by_class_name("value").text
            # 解析出图片列表
            picture_div_ul_li = browser.find_elements_by_class_name("tab-trigger")
            # 详情在第4个DIV里面
            details = browser.find_elements_by_class_name("obj-content")[4]
            # 点击一下'加载更多'
            browser.find_elements_by_class_name("obj-expand")[1].click()
            # 构建详情列表
            temp_color_size_dict = {}
            for td in details.find_elements_by_tag_name("td"):
                if td.text is not None and len(td.text) != 0:
                    details_info_list.append(td.text)
            # 把详情列表里面的颜色 尺码提取出来
            temp_color_size_dict["id"] = i + 1
            temp_color_size_dict["name"] = title
            temp_color_size_dict["color"] = details_info_list[details_info_list.index("颜色") + 1]
            temp_color_size_dict["size"] = details_info_list[details_info_list.index("尺码") + 1]
            temp_color_size_dict["price"] = price
            color_size_list.append(
                {'id': i + 1, 'name': title, 'color': details_info_list[details_info_list.index("颜色") + 1],
                 'size': details_info_list[details_info_list.index("尺码") + 1], 'price': price})

            print("商品名称:", title, "编号:", i)
            print("细节:", temp_color_size_dict)
            print("图片:", picture_list)
            time.sleep(0.5)
        except Exception as e:
            print("出现异常:商品不存在或者已经下架!")
            print(e)

    print("最终的数据:", color_size_list)

    return color_size_list


def write_excel(self, color_size_list):
    # 写excel
    write_excel = xlwt.Workbook()  # 创建工作簿
    write_sheet1 = write_excel.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    header = [u'编号', u'名称', u'颜色', u'尺码']
    for i in range(4):
        write_sheet1.write(0, i, header[i])
        write_sheet1.write(0, i, header[i])
        write_sheet1.write(0, i, header[i])
        write_sheet1.write(0, i, header[i])

    for i in range(len(color_size_list)):
        info = color_size_list[i]
        write_sheet1.write(i + 1, 0, info["id"])
        write_sheet1.write(i + 1, 1, info["name"])
        write_sheet1.write(i + 1, 2, info["color"])
        write_sheet1.write(i + 1, 3, info["size"])

    '''
    从第四列开始插入数据
    1 2 3 4
    q w e r

    '''
    for i in range(sheet1.nrows):
        for j in range(4, sheet1.ncols):
            write_sheet1.write(i, j, sheet1.row_values(i)[j])

    write_excel.save(
        "../excel/" + str(time.asctime(time.localtime(time.time()))).replace(" ", "_").replace(":", "_") + "_data.xlsx")


class MainWindow:
    def open_file(self, event):
        filename = tkinter.filedialog.askopenfilename(filetypes=[("Excel表格", "xls"), ("Excel表格", "xlsx")])
        if filename:
            self.filename_label["text"] = filename
        else:
            self.filename_label["text"] = u"你没有选择任何文件"
            # tkinter.messagebox.showinfo("文件选择", u"你没有选择任何文件")

    def start(self, event):

        tkinter.messagebox.showinfo("文件选择", u"开始获取")
        self.console["text"] += "1111111\n"

    def exit(self, event):
        exit()

    def center_window(self):
        width = 500
        height = 200
        screenwidth = self.frame.winfo_screenwidth()
        screenheight = self.frame.winfo_screenheight()
        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.frame.geometry(size)

    def __init__(self):
        self.frame = Tk()
        self.frame.title("数据快速分类工具")
        self.center_window()

        self.scroll = Scrollbar()
        self.scroll.pack(side=RIGHT, fill=Y)
        self.console = Text(self.frame, width=200, height=200, bg="black", fg="green",
                            font=('Helvetica', '14', 'bold'), yscrollcommand=self.scroll.set)

        self.scroll.config(command=self.console.yview)

        self.filename_label = Label(self.frame)

        self.start_button = Button(self.frame, text=u"开始转换")
        self.open_file_button = Button(self.frame, text=u"打开文件")
        self.exit_button = Button(self.frame, text=u"退出程序")

        self.filename_label.pack()
        self.open_file_button.pack()
        self.start_button.pack()
        self.exit_button.pack()
        self.console.pack(side=LEFT, fill=Y)

        self.open_file_button.bind("<ButtonRelease-1>", self.open_file)
        self.exit_button.bind("<ButtonRelease-1>", self.exit)
        self.start_button.bind("<ButtonRelease-1>", self.start)

    def show(self):
        self.frame.mainloop()


if __name__ == "__main__":
    window = MainWindow()
    window.show()
