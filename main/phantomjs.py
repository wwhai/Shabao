# -*- coding:utf-8 -*-
import os
import time
import uuid

import requests
import urllib3
import xlrd
from selenium import webdriver

requests.packages.urllib3.disable_warnings()
# # 进入浏览器设置
# options = webdriver.PhantomJS()
# # 设置中文
# options.add_argument('lang=zh_CN.UTF-8')
# # 更换头部
# options.add_argument(
#     'user-agent="Mozilla/5.0 (iPod; U; CPU iPhone OS 2_1 like Mac OS X; ja-jp) AppleWebKit/525.18.1 (KHTML, like Gecko) Version/3.1.1 Mobile/5F137 Safari/525.20"')
# options.add_argument("X-Forwarded-For=")
# options.add_argument("X-Real-IP=")

DRIVER_PATH = "../driver/phantomjs.exe"
EXCEL = '../excel/data.xlsx'
browser = webdriver.PhantomJS(executable_path=DRIVER_PATH)
excel = xlrd.open_workbook(EXCEL)
sheet1 = excel.sheet_by_name(u'Sheet1')
http_client = urllib3.PoolManager()
print("总共有%d个数据" % sheet1.nrows)


def get_excel_data():
    print("开始测试环境!")
    test = browser.get("https://detail.1688.com/")
    if test is not None:
        print("测试成功!")
    for i in range(1, sheet1.nrows - 1):
        row = sheet1.row_values(i)
        print("正在获取连接:", row[6])

        try:
            browser.get(row[6])
            title = browser.find_element_by_class_name("d-title").text
            price = browser.find_element_by_class_name("value").text
            picture_div_ul_li = browser.find_elements_by_class_name("tab-trigger")
            details = browser.find_elements_by_class_name("obj-content")[4]  # 详情在第四个DIV里面
            browser.find_elements_by_class_name("obj-expand")[1].click()  # 点击一下'加载更多'
            picture_list = []
            details_info_list = []
            os.mkdir('../data/' + str(i))  # 新建文件夹

            for td in details.find_elements_by_tag_name("td"):
                if td.text is not None and len(td.text) != 0:
                    details_info_list.append(td.text)

            print("K", details_info_list[len(details_info_list) - 2], " V:",
                  details_info_list[len(details_info_list) - 1])

            for li in picture_div_ul_li:
                box_img = li.find_element_by_class_name("box-img")
                picture_url = box_img.find_element_by_tag_name("img").get_attribute("src").replace(".60x60.", ".")
                response = http_client.request("GET", picture_url)
                picture = response.data
                picture_list.append(picture_url)
                print("正在下载图片......")
                with open('../data/' + str(i) + "/" + str(uuid.uuid4()) + ".jpg", 'wb') as f:
                    f.write(picture)
                print("图片下载完成!")

            # 输出具体的信息
            print("商品名称:", title, "编号:", i)
            print("价格:", price)
            print("详情:", details_info_list)
            print("颜色:", details_info_list[details_info_list.index("颜色") + 1])
            print("尺码:", details_info_list[details_info_list.index("尺码") + 1])

            print("图片:", picture_list)
            time.sleep(0.5)
        except Exception as e:
            print("出现异常:商品不存在或者已经下架!")
            print(e)


if __name__ == "__main__":
    get_excel_data()
