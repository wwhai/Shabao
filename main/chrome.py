# -*- coding:utf-8 -*-
import time

import xlrd
from selenium import webdriver

# # 进入浏览器设置
# options = webdriver.ChromeOptions()
# # 设置中文
# options.add_argument('lang=zh_CN.UTF-8')
# # 更换头部
# options.add_argument(
#     'user-agent="Mozilla/5.0 (iPod; U; CPU iPhone OS 2_1 like Mac OS X; ja-jp) AppleWebKit/525.18.1 (KHTML, like Gecko) Version/3.1.1 Mobile/5F137 Safari/525.20"')
# options.add_argument("X-Forwarded-For=")
# options.add_argument("X-Real-IP=")

DRIVER_PATH = u"../driver/chromedriver.exe"
EXCEL = '../excel/data.xlsx'
browser = webdriver.PhantomJS(executable_path=DRIVER_PATH)
excel = xlrd.open_workbook(EXCEL)
sheet1 = excel.sheet_by_name(u'Sheet1')
print("总共有%d个数据" % sheet1.nrows)
for i in range(1, sheet1.nrows - 1):
    row = sheet1.row_values(i)
    print("正在获取连接:", row[6])

    try:
        browser.get(row[6])
        title = browser.find_element_by_class_name("d-title").text
        price = browser.find_element_by_class_name("value").text
        picture_div_ul_li = browser.find_elements_by_class_name("tab-trigger")
        details = browser.find_elements_by_class_name("obj-content")[4]  # 详情在第四个DIV里面
        browser.find_elements_by_tag_name("obj-expand")[1].find_element_by_tag_name("a").click()  # 要点击一下详情
        picture_list = []
        details_info_list = []
        details_info_dict = {}

        for td in details.find_elements_by_tag_name("td"):
            if td.text is not None and len(td.text) != 0:
                details_info_list.append(td.text)

        # for num in range(len(details_info_list) - 1):
        #     print(details_info_list[num], details_info_list[num + 1])

        for li in picture_div_ul_li:
            box_img = li.find_element_by_class_name("box-img")
            picture_url = box_img.find_element_by_tag_name("img").get_attribute("src").replace(".60x60.", ".")
            picture_list.append(picture_url)
        print("商品名称:", title, "编号:", i)
        print("价格:", price)
        print("详情:", details_info_list)
        print("图片:", picture_list)
        time.sleep(0.5)
    except Exception as e:
        print("商品不存在或者已经下架!")
        print(e)
