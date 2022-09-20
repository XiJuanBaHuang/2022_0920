import cProfile
import os
import pstats
import re

import xlwt
from xlwt import Workbook

# import re
s1 = "新增确诊人数"
s2 = "新增无症状感染人数"
new_confirmed_cases = re.compile(r'新增确诊病例(\d+)')
new_unknown_cases = re.compile(r"新增无症状感染者(\d+)")
get_time = re.compile(r"(\d+)月(\d+)日")
all_provinces = {
    "黑龙江": 0, "辽宁": 0, "吉林": 0, "河北": 0, "河南": 0,
    "湖北": 0, "湖南": 0, "山东": 0, "山西": 0, "陕西": 0,
    "安徽": 0, "浙江": 0, "江苏": 0, "福建": 0, "广东": 0,
    "海南": 0, "四川": 0, "云南": 0, "贵州": 0, "青海": 0,
    "甘肃": 0, "江西": 0, "台湾": 0,
    "内蒙古": 0, "宁夏": 0, "新疆": 0, "西藏": 0, "广西": 0,
    "北京": 0, "上海": 0, "天津": 0, "重庆": 0,
    "香港": 0, "澳门": 0,
}
all_provinces_list = [
    "黑龙江", "辽宁", "吉林", "河北", "河南",
    "湖北", "湖南", "山东", "山西", "陕西",
    "安徽", "浙江", "江苏", "福建", "广东",
    "海南", "四川", "云南", "贵州", "青海",
    "甘肃", "江西", "台湾",
    "内蒙古", "宁夏", "新疆", "西藏", "广西",
    "北京", "上海", "天津", "重庆",
    "香港", "澳门",
]
wb = Workbook()
sheet1 = wb.add_sheet('test1')


def get_data(year):
    path = ""
    if year == "2020":
        path = "D:\\Python\\Data_Task1\\2020_order"  # 文件夹目录
    elif year == "2021":
        path = "D:\\Python\\Data_Task1\\2021_order"  # 文件夹目录
    elif year == "2022":
        path = "D:\\Python\\Data_Task1\\2022_order"  # 文件夹目录
    files = os.listdir(path)  # 得到文件夹下的所有文件名称
    data_txts = []
    for file in files:  # 遍历文件夹
        position = path + '\\' + file  # 构造绝对路径，"\\"，其中一个'\'为转义符
        with open(position, "r", encoding='utf-8') as f:  # 打开文件
            data = f.read()  # 读取文件
            data_txts.append(data)
    return data_txts  # data_txts type : still list


def get_month_date(news):  # news: string; need a regex
    temp = get_time.findall(news)           # get a list
    Month_Day = temp[0]                     # get a tuple
    month = Month_Day[0]                    # get a string
    day = Month_Day[1]
    return month, day


def get_new_cases_number(news):                     # news: str,
    temp = new_confirmed_cases.findall(news)        # temp: list
    if temp:
        return temp[0]                              # number: str
    else:
        return 0


def get_new_unknown_cases_number(news):             # news: str,
    temp = new_unknown_cases.findall(news)        # temp: list
    if temp:
        return temp[0]                              # number: str
    else:
        return 0


def get_content(year, number):
    txts = get_data(year)

    for i in range(4):
        sheet1.col(i).width = 5000
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style = xlwt.XFStyle()
    style.alignment = alignment

    if year == "2020":
        i = int(1)
        sheet1.write(0, i, s1, style)
        i += 1
        sheet1.write(0, i, s2, style)

    j = number  # int(1)
    for everyday_news in txts:
        month, day = get_month_date(everyday_news)
        time = year + '/' + month + '/' + day
        sheet1.write(j, 0, time, style)
        s1_number = get_new_cases_number(everyday_news)
        s2_number = get_new_unknown_cases_number(everyday_news)
        sheet1.write(j, 1, s1_number, style)
        sheet1.write(j, 2, s2_number, style)
        j += 1

    wb.save(r'D:\Python\Data_Task1\execl_test\sum_all.py')
    return j


def main():
    s = ["2020", "2021", "2022"]
    next_Number = int(1)
    for i in range(len(s)):
        next_Number = get_content(s[i], next_Number)


'''
    ncalls：表示函数调用的次数；
    tottime：表示指定函数的总的运行时间，除掉函数中调用子函数的运行时间；
    percall：（第一个percall）等于 tottime/ncalls；
    cumtime：表示该函数及其所有子函数的调用运行的时间，即函数开始调用到返回的时间；
    percall：（第二个percall）即函数运行一次的平均时间，等于 cumtime/ncalls；
    filename:lineno(function)：每个函数调用的具体信息；
'''


# cProfile.run('re.compile("foo|bar")')
# cProfile.run('main()')

# 保存在当前目录,按照时间进行排序
cProfile.run('main()', filename="result.out", sort="cumulative")

# 创建Stats对象
p = pstats.Stats("result.out")

# strip_dirs(): 去掉无关的路径信息
# sort_stats(): 排序，支持的方式和上述的一致
# print_stats(): 打印分析结果，可以指定打印前几行

# 和直接运行cProfile.run("test()")的结果是一样的
# p.strip_dirs().sort_stats(-1).print_stats()

# 按照函数名排序，只打印前3行函数的信息, 参数还可为小数,表示前百分之几的函数信息
# p.strip_dirs().sort_stats("name").print_stats(3)

# 按照运行时间和函数名进行排序
p.strip_dirs().sort_stats("cumulative", "name").print_stats(0.1)

# 如果想知道有哪些函数调用了sum_num
# p.print_callers(0.5, "sum_num")

# 查看test()函数中调用了哪些函数
# p.print_callees("test")
