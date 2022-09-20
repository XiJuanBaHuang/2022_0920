import cProfile
import os
import re

import xlwt
from xlwt import Workbook


s2 = "新增无症状感染人数"
new_unknown_cases = re.compile(r"新增无症状感染者(\d+)")
get_time = re.compile(r"(\d+)月(\d+)日")

day_cases_province_type1 = re.compile(r"新增无症状感染者(\d+)例[，。].*?境外输入(\d+)例.*?本土(\d+)例（(.*?)）")
day_cases_province_type0 = re.compile(r"新增无症状感染者(\d+)例（.*?境外输入.*?）")

sample0 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者29例（境外输入1例）；
'''  # 2020 / 05 / 25
sample0_1 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者4例（境外输入4例）；
'''  # 2020 / 07 / 14
sample0_2 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者34例（均为境外输入）；
'''  # 2020 / 08 / 21
sample0_3 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者11例（均为境外输入）；
'''  # 2021 / 02 / 15

sample1 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者31例，其中境外输入30例，本土1例（在江西）；
'''  # 2021 / 03 / 25
sample1_0 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者8例，其中境外输入5例，本土3例（均在云南）；
'''  # 2021 / 03 / 30
sample1_1 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者25例，其中境外输入15例，本土10例（安徽7例，辽宁3例）；
'''  # 2021 / 05 / 14
sample1_2 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者25例，其中境外输入24例，本土1例（在广东）；
'''  # 2021 / 05 / 22
sample1_3 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者35例，其中境外输入25例，本土10例（江苏7例，辽宁1例，安徽1例，广东1例）
'''  # 2021 / 07 / 22
sample1_4 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者20例，其中境外输入14例，本土6例（均在福建莆田市）；
'''  # 2021 / 09 / 16
sample1_6 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者31例，其中境外输入20例，本土11例（山东4例，均在日照市；黑龙江3例，均在黑河市；北京2例，均在昌平区；云南2例，均在德宏傣族景颇族自治州）；
'''  # 2021 / 10 / 27
sample1_5 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者37例，其中境外输入35例，本土2例（北京1例，在朝阳区；广东1例，在广州市）；
'''                             # 2022 / 01 / 18
sample1_7 = '''31个省（自治区、直辖市）和新疆生产建设兵团报告新增无症状感染者399例，其中境外输入77例，\
本土322例（山东106例，其中青岛市93例、威海市13例；吉林80例，其中吉林市65例、长春市10例、梅河口市2例、延边朝鲜族自治州2例、松原市1例；\
上海62例，其中松江区14例、闵行区11例、宝山区11例、徐汇区9例、嘉定区7例、浦东新区3例、青浦区3例、普陀区2例、长宁区1例、静安区1例；\
云南24例，其中德宏傣族景颇族自治州23例、临沧市1例；广东20例，均在东莞市；黑龙江14例，均在牡丹江市；江苏4例，均在连云港市；\
广西4例，其中防城港市2例、崇左市1例、百色市1例；山西2例，均在运城市；辽宁2例，其中沈阳市1例、丹东市1例；内蒙古1例，在阿拉善盟；\
安徽1例，在安庆市；重庆1例，在渝北区；甘肃1例，在白银市）；
'''                             # 2022 / 03 /06

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


def get_new_unknown_cases_number(news):  # news: str,
    temp = new_unknown_cases.findall(news)  # temp: list
    if temp:
        return temp[0]  # number: str
    else:
        return 0


def get_per_province_cases(item, news):
    temp1 = day_cases_province_type1.findall(news)
    p = re.compile(item + '.*?' + '(\d+)')
    p_test = re.compile('在' + item)
    number = 0
    if len(temp1) > 0:
        str0 = temp1[0]
        str1 = str0[3]
        rex1 = p.findall(str1)
        if len(rex1) > 0:
            number = rex1[0]
        else:
            rex_test = p_test.findall(str1)
            if len(rex_test) > 0:
                number = str0[2]
            else:
                number = 0
        number = turn_to_int(number)
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
        sheet1.write(0, i, s2, style)
        for key, value in all_provinces.items():
            i += 1
            sheet1.write(0, i, key, style)

    j = number  # int(1)
    for everyday_news in txts:
        month, day = get_month_date(everyday_news)
        time = year + '/' + month + '/' + day
        sheet1.write(j, 0, time, style)
        s2_number = get_new_unknown_cases_number(everyday_news)
        sheet1.write(j, 1, s2_number, style)
        count = int(2)
        for item in all_provinces_list:
            p_number = get_per_province_cases(item, everyday_news)
            sheet1.write(j, count, str(p_number), style)
            count += 1

        j += 1

    wb.save('D:\\Python\\execl_test\\per_new_unknown.xls')
    return j


def main():
    s = ["2020", "2021", "2022"]
    next_Number = int(1)
    for i in range(len(s)):
        next_Number = get_content(s[i], next_Number)


# main()
cProfile.run('re.compile("foo|bar")')
