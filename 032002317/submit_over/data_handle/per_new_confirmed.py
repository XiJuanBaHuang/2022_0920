import cProfile
import os
import re

import xlwt
from xlwt import Workbook

s1 = "新增确诊人数"
new_confirmed_cases = re.compile(r'新增确诊病例(\d+)')
get_time = re.compile(r"(\d+)月(\d+)日")

day_cases_province_type1 = re.compile(r"新增确诊病例(\d+)例[，。].*?境外输入病例(\d+)例（(.*?)）.*?本土病例(\d+)例（(.*?)）")
day_cases_province_type1_1 = re.compile(r"新增确诊病例(\d+)例[，。].*?(\d+)例为境外输入病例（(.*?)）.*?(\d+)例为本土病例（(.*?)）")
day_cases_province_type2 = re.compile(r"新增确诊病例(\d+)例[，。].*?境外输入病例（(.*?)）")
day_cases_province_type3 = re.compile(r"新增确诊病例(\d+)例[，。].*?本土病例（(.*?)）")

type1_sample = '''新增确诊病例43例。\
其中境外输入病例14例（广东4例，上海2例，河南2例，广西2例，北京1例，福建1例，山东1例，四川1例），\
含4例由无症状感染者转为确诊病例（河南2例，广东1例，四川1例）；\
本土病例29例（内蒙古16例，其中阿拉善盟15例、鄂尔多斯市1例；甘肃6例，均在兰州市；北京3例，均在昌平区；\
宁夏3例，其中银川市1例、吴忠市1例、中卫市1例；山东1例，在日照市）
'''
type1_1_sample = '''　　5月2日0—24时，31个省（自治区、直辖市）和新疆生产建设兵团报告新增确诊病例2例，\
其中1例为境外输入病例（在上海），1例为本土病例（在山西）；无新增死亡病例；无新增疑似病例。
　　31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者12例（境外输入2例）；
'''
type2_sample = '''6月1日0—24时，31个省（自治区、直辖市）和新疆生产建设兵团报告新增确诊病例5例，\
均为境外输入病例（四川2例，上海1例，广东1例，陕西1例）；无新增死亡病例；无新增疑似病例。
'''
type3_sample = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增确诊病例3例，均为本土病例（辽宁2例，吉林1例）
'''

all_provinces = {
    "黑龙江": 0, "辽宁": 0, "吉林": 0, "河北": 0, "河南": 0,
    "湖北": 0, "湖南": 0, "山东": 0, "山西": 0, "陕西": 0,
    "安徽": 0, "浙江": 0, "江苏": 0, "福建": 0, "广东": 0,
    "海南": 0, "四川": 0, "云南": 0, "贵州": 0, "青海": 0,
    "甘肃": 0, "江西": 0,
    "内蒙古": 0, "宁夏": 0, "新疆": 0, "西藏": 0, "广西": 0,
    "北京": 0, "上海": 0, "天津": 0, "重庆": 0,
}
all_provinces_list = [
    "黑龙江", "辽宁", "吉林", "河北", "河南",
    "湖北", "湖南", "山东", "山西", "陕西",
    "安徽", "浙江", "江苏", "福建", "广东",
    "海南", "四川", "云南", "贵州", "青海",
    "甘肃", "江西",
    "内蒙古", "宁夏", "新疆", "西藏", "广西",
    "北京", "上海", "天津", "重庆",
]
wb = Workbook()
sheet1 = wb.add_sheet('test1')


def get_data(year):
    path = ""
    if year == "2020":
        path = "D:\\Python\\2020_order"  # 文件夹目录
    elif year == "2021":
        path = "D:\\Python\\2021_order"  # 文件夹目录
    elif year == "2022":
        path = "D:\\Python\\2022_order"  # 文件夹目录
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


def get_new_cases_number(news):  # news: str,
    temp = new_confirmed_cases.findall(news)  # temp: list
    if temp:
        return temp[0]  # number: str
    else:
        return 0


def get_per_province_cases(item, news):
    temp1 = day_cases_province_type1.findall(news)
    temp1_1 = day_cases_province_type1_1.findall(news)
    temp2 = day_cases_province_type2.findall(news)
    temp3 = day_cases_province_type3.findall(news)
    p = re.compile(item + '.*?' + '(\d+)')
    p_test = re.compile('在' + item)
    number = 0
    if len(temp1) > 0:
        str0 = temp1[0]
        str1 = str0[2]
        str2 = str0[4]
        rex1 = p.findall(str1)
        rex2 = p.findall(str2)
        if len(rex1) > 0:
            number1 = rex1[0]
        else:
            rex_test = p_test.findall(str1)
            if len(rex_test) > 0:
                number1 = str0[1]
            else:
                number1 = 0
        if len(rex2) > 0:
            number2 = rex2[0]
        else:
            rex_test = p_test.findall(str2)
            if len(rex_test) > 0:
                number2 = str0[1]
            else:
                number2 = 0
        number1 = turn_to_int(number1)
        number2 = turn_to_int(number2)
        number = number1 + number2

    elif len(temp1_1) > 0:
        str1 = temp1_1[0]
        rex1 = p.findall(str1[2])
        rex2 = p.findall(str1[4])
        if len(rex1) > 0:
            number1 = rex1[0]
        else:
            rex_test = p_test.findall(str1[2])
            if len(rex_test) > 0:
                number1 = str1[1]
            else:
                number1 = 0
        if len(rex1) > 0:
            number2 = rex2[0]
        else:
            rex_test = p_test.findall(str1[4])
            if len(rex_test) > 0:
                number2 = str1[3]
            else:
                number2 = 0
        number1 = turn_to_int(number1)
        number2 = turn_to_int(number2)
        number = number1 + number2

    elif len(temp2) > 0:
        str1 = temp2[0]
        rex1 = p.findall(str1[1])
        if len(rex1) > 0:
            number = rex1[0]
        else:
            rex_test = p_test.findall(str1[1])
            if len(rex_test) > 0:
                number = str1[0]

    elif len(temp3) > 0:
        str1 = temp3[0]
        rex1 = p.findall(str1[1])
        if len(rex1) > 0:
            number = rex1[0]
        else:
            rex_test = p_test.findall(str1[1])
            if len(rex_test) > 0:
                number = str1[0]
    return number


def turn_to_int(x):
    if isinstance(x, int):
        return x
    elif isinstance(x, str):
        return eval(x)


def get_content(year, number):
    txts = get_data(year)

    for i in range(40):
        sheet1.col(i).width = 5000
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style = xlwt.XFStyle()
    style.alignment = alignment

    if year == "2020":
        i = int(1)
        sheet1.write(0, i, s1, style)
        for key, value in all_provinces.items():
            i += 1
            sheet1.write(0, i, key, style)

    j = number  # int(1)
    for everyday_news in txts:
        month, day = get_month_date(everyday_news)
        time = year + '/' + month + '/' + day
        sheet1.write(j, 0, time, style)
        s1_number = get_new_cases_number(everyday_news)
        sheet1.write(j, 1, s1_number, style)
        count = int(2)
        for item in all_provinces_list:
            p_number = get_per_province_cases(item, everyday_news)
            sheet1.write(j, count, str(p_number), style)
            count += 1
        j += 1

    wb.save('D:\\Python\\execl_test\\per_new_confirmed.xls')
    return j


def main():
    s = ["2020", "2021", "2022"]
    next_Number = int(1)
    for i in range(len(s)):
        next_Number = get_content(s[i], next_Number)


# main()
cProfile.run('re.compile("foo|bar")')
