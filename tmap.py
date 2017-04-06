# 腾讯地图：http://map.qq.com/  腾讯地图poi：http://lbs.qq.com/webservice_v1/index.html
# coding:utf-8
# github:https://github.com/tianyu8969/python-to-tmap

import json
import xlwt
from datetime import datetime
from urllib import request
from urllib.parse import quote
import sys
import time

# 获取当前日期
today = datetime.today()
# 将获取到的datetime对象仅取日期如：2017-4-6
today_date = datetime.date(today)

json_name = 'data_tmap.json'
# 腾讯地图poi：http://apis.map.qq.com/ws/place/v1/search
# 请替换为自己申请的key值：申请Web服务API类型KEY http://lbs.qq.com/key.html
# filter=category=医疗保健
url_amap = 'http://apis.map.qq.com/ws/place/v1/search?boundary=region(上海,0)&page_size=20&page_index=pageindex&keyword=卫生服务中心&filter=category=医疗保健&output=json&orderby=_distance&key=NSKBZ-P3C3W-6DVRY-OZLJ5-TYPQ5-LBFVY'
page_size = 20  # 每页条目数，最大限制为20条
page_index = r'page_index=1'  # 显示页码
total_record = 0  # 定义全局变量，总行数
# Excel表头
hkeys = ['id', 'POI类型', '医院名称', '医院类型', '医院地址', '联系电话', '北纬', '东经', '省份名称', '城市名称', '区域代码', '区域名称']
# 获取数据列
bkeys = ['id', 'type', 'title', 'category', 'address', 'tel', ['location', 'lat', 'lng'],
         ['ad_info', 'province', 'city', 'adcode', 'district']]


# 获取数据
def get_data(pageindex):
    global total_record
    # 暂停500毫秒，防止过快取不到数据
    time.sleep(0.5)
    print('解析页码： ' + str(pageindex) + ' ... ...')
    url = url_amap.replace('pageindex', str(pageindex))
    # 中文编码
    url = quote(url, safe='/:?&=')
    html = ""
    with request.urlopen(url) as f:
        html = f.read()
    rr = json.loads(html)
    if total_record == 0:
        total_record = int(rr['count'])
    return rr['data']


def getPOIdata():
    global total_record
    print('获取POI数据开始')
    josn_data = get_data(1)
    if (total_record % page_size) != 0:
        page_number = int(total_record / page_size) + 2
    else:
        page_number = int(total_record / page_size) + 1

    with open(json_name, 'w') as f:
        # 去除最后]
        f.write(json.dumps(josn_data).rstrip(']'))
        print('已保存到json文件：' + json_name)
        for each_page in range(2, page_number):
            html = json.dumps(get_data(each_page)).lstrip('[').rstrip(']')
            if html:
                html = "," + html
            f.write(html)
            print('已保存到json文件：' + json_name)
        f.write(']')
    print('获取POI数据结束')


# 写入数据到excel
def write_data_to_excel(name):
    # 从文件中读取数据
    fp = open(json_name, 'r')
    result = json.loads(fp.read())
    # 实例化一个Workbook()对象(即excel文件)
    wbk = xlwt.Workbook()
    # 新建一个名为Sheet1的excel sheet。此处的cell_overwrite_ok =True是为了能对同一个单元格重复操作。
    sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)

    # 创建表头
    # for循环访问并获取数组下标enumerate函数
    for index, hkey in enumerate(hkeys):
        sheet.write(0, index, hkey)

    # 遍历result中的每个元素。
    for i in range(len(result)):
        values = result[i]
        n = i + 1
        index = 0
        for key in bkeys:
            val = ""
            islist = type(key) == list
            if islist:
                keyv = key[0]  # 获取属性
                key = key[1:]  # 切片，从第一个开始
                for ki, kv in enumerate(key):
                    val = values[keyv][kv]
                    sheet.write(n, index, val)
                    index = index + 1
            # 判断是否存在属性key
            elif key in values.keys():
                val = values[key]
                sheet.write(n, index, val)
            if not islist:
                index = index + 1
    wbk.save(name + str(today_date) + '.xls')
    print('保存到excel文件： ' + name + str(today_date) + '.xls ！')


if __name__ == '__main__':
    # 写入数据到json文件，第二次运行可注释
    getPOIdata()
    # 读取json文件数据写入到excel
    write_data_to_excel("上海卫生服务中心-腾讯地图")
