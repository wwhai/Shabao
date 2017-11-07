# -*- coding:utf-8 -*-
import os
import time

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

            # 新建文件夹
            os.mkdir('../data/' + str(i))
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

            # # 开始下载图片
            # for li in picture_div_ul_li:
            #     box_img = li.find_element_by_class_name("box-img")
            #     picture_url = box_img.find_element_by_tag_name("img").get_attribute("src").replace(".60x60.", ".")
            #     response = http_client.request("GET", picture_url)
            #     picture = response.data
            #     picture_list.append(picture_url)
            #     print("正在下载图片......")
            #     with open('../data/' + str(i) + "/" + str(uuid.uuid4()) + ".jpg", 'wb') as f:
            #         f.write(picture)
            #     print("图片下载完成!")
            # 输出具体的信息
            print("商品名称:", title, "编号:", i)
            print("细节:", temp_color_size_dict)
            print("图片:", picture_list)
            time.sleep(0.5)
        except Exception as e:
            print("出现异常:商品不存在或者已经下架!")
            print(e)

    print("最终的数据:", color_size_list)

    return color_size_list


# 写excel
excel = xlwt.Workbook()  # 创建工作簿
sheet1 = excel.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet


def write_excel(color_size_list):
    header = [u'编号', u'名称', u'颜色', u'尺码']
    sheet1.write(0, 0, header[0])
    sheet1.write(0, 1, header[1])
    sheet1.write(0, 2, header[2])
    sheet1.write(0, 3, header[3])

    for i in range(len(color_size_list)):
        info = color_size_list[i]
        sheet1.write(i + 1, 0, info["id"])
        sheet1.write(i + 1, 1, info["name"])
        sheet1.write(i + 1, 2, info["color"])
        sheet1.write(i + 1, 3, info["size"])
    excel.save("../excel/new.xls")


lis = [{'id': 2, 'name': '2017春季新款漆皮系带厚底防水台松糕休闲女鞋潮', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39', 'price': '26.00'},
       {'id': 3, 'name': '新款英伦风女鞋2017秋冬季单鞋休闲百搭粗跟中跟韩版皮鞋厂家直销', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39',
        'price': '27.00'},
       {'id': 4, 'name': '2017秋季新款厚底女鞋坡跟休闲女单鞋深口系带圆头舒适松糕鞋女', 'color': '红色,黑色,香槟', 'size': '35,36,37,38,39',
        'price': '28.00'}]

if __name__ == "__main__":
    # get_excel_data()
    write_excel(lis)
