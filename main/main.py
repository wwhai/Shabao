# -*- coding:utf-8 -*-
import time

import xlrd
from selenium import webdriver


# 进入浏览器设置
options = webdriver.ChromeOptions()
# 设置中文
options.add_argument('lang=zh_CN.UTF-8')
# 更换头部
options.add_argument('user-agent="Mozilla/5.0 (iPod; U; CPU iPhone OS 2_1 like Mac OS X; ja-jp) AppleWebKit/525.18.1 (KHTML, like Gecko) Version/3.1.1 Mobile/5F137 Safari/525.20"')
options.add_argument("X-Forwarded-For=")
options.add_argument("X-Real-IP=")
browser = webdriver.Chrome(executable_path="../driver/chromedriver.exe")
excel = xlrd.open_workbook('../excel/data.xlsx')
sheet1 = excel.sheet_by_name(u'Sheet1')
print("总共有%d个数据" % sheet1.nrows)
for i in range(1, sheet1.nrows - 1):
    row = sheet1.row_values(i)
    print("正在获取连接:", row[6])

    try:
        browser.get(row[6])
        title = browser.find_element_by_class_name("d-title").text
        price = browser.find_element_by_class_name("value").text
        print(title)
        print(price)
        time.sleep(0.5)
    except Exception as e:
        print("商品不存在或者已经下架!")
        print(e)
