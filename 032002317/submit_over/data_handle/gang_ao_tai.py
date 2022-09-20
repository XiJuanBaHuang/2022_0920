import cProfile
import os
import re

import xlwt
from xlwt import Workbook

# s1 = "新增确诊人数"
# s2 = "新增无症状感染人数"
get_time = re.compile(r"(\d+)月(\d+)日")
gang_ao_tai = {
    "香港": 0, "澳门": 0, "台湾": 0,
}
gang_ao_tai_list = [
    "香港", "澳门", "台湾",
]
wb = Workbook()
sheet1 = wb.add_sheet('test1')


def get_data():
    path = "D:\\Python\\gang_ao_tai"
    files = os.listdir(path)  # 得到文件夹下的所有文件名称
    data_txts = []
    for file in files:  # 遍历文件夹
        position = path + '\\' + file  # 构造绝对路径，"\\"，其中一个'\'为转义符
        with open(position, "r", encoding='utf-8') as f:  # 打开文件
            data = f.read()  # 读取文件
            data_txts.append(data)
    return data_txts  # data_txts type : still list


def get_month_date(news):  # news: string; need a regex
    temp = get_time.findall(news)  # get a list
    Month_Day = temp[0]  # get a tuple
    month = Month_Day[0]  # get a string
    day = Month_Day[1]
    return month, day


def get_per_province_cases(item, news):
    # rex = re.search(item + '.*?' + '(\d+)', news)
    p = re.compile(item + '.*?' + '(\d+)')
    rex = p.findall(news)
    number = 0
    if rex:
        if rex[0]:
            number = get_digits(rex[0])
    return number


def get_digits(string):
    p = re.compile(r"\d+")
    data = p.findall(string)
    return data[0]


def daily_sub(province, today, yesterday):
    today_number = get_per_province_cases(province, today)
    yesterday_number = get_per_province_cases(province, yesterday)
    return eval(today_number) - eval(yesterday_number)


def get_content(number):
    txts = get_data()

    for i in range(6):
        sheet1.col(i).width = 5000
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style = xlwt.XFStyle()
    style.alignment = alignment

    i = int(0)
    for key, value in gang_ao_tai.items():
        i += 1
        sheet1.write(0, i, key + '当日新增确诊人数', style)

    j = number  # int(1)
    for today_news in txts:
        if j <= 344:  # 345 - 1
            year = "2020"
        elif j <= 709:
            year = "2021"
        else:
            year = "2022"
        month, day = get_month_date(today_news)
        time = year + '/' + month + '/' + day
        sheet1.write(j, 0, time, style)
        count = int(1)
        if j == 1:
            for item in gang_ao_tai_list:
                p_number = get_per_province_cases(item, today_news)
                sheet1.write(j, count, str(p_number), style)
                count += 1
        else:
            yesterday_news = txts[j - 2]
            for item in gang_ao_tai_list:
                p_number = daily_sub(item, today_news, yesterday_news)
                sheet1.write(j, count, str(p_number), style)
                count += 1
        j += 1

    wb.save('D:\\Python\\execl_test\\gang_ao_tai_new_confirmed.xls')


def main():
    next_Number = int(1)
    get_content(next_Number)


# main()
# cProfile.run('re.compile("foo|bar")')
cProfile.run('main()')