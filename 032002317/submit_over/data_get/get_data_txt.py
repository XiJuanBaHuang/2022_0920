import asyncio
import os

from pyppeteer import launcher

# 在导入 launch 之前 把 --enable-automation 禁用 防止监测webdriver
launcher.DEFAULT_ARGS.remove("--enable-automation")

from pyppeteer import launch
from bs4 import BeautifulSoup


async def pyppteer_fetchUrl(url):
    browser = await launch({'headless': False, 'dumpio': True, 'autoClose': True})
    page = await browser.newPage()

    await page.goto(url)
    await asyncio.wait([page.waitForNavigation()])
    str = await page.content()
    await browser.close()
    return str


def fetchUrl(url):
    return asyncio.get_event_loop().run_until_complete(pyppteer_fetchUrl(url))


def getPageUrl():
    for page in range(1, 43):
        # 1. 1  ~ 14   over
        # 2. 15 ~ 28
        # 3. 29 ~ 41
        if page == 1:
            yield 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml'
        else:
            surl = 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd_' + str(page) + '.shtml'
        yield surl


def getTitleUrl(html):
    bsobj = BeautifulSoup(html, 'html.parser')
    titleList = bsobj.find('div', attrs={"class": "list"}).ul.find_all("li")
    for item in titleList:
        link = "http://www.nhc.gov.cn" + item.a["href"]
        title = item.a["title"]
        date = item.span.text
        yield title, link, date


def getContent(html):
    bsobj = BeautifulSoup(html, 'html.parser')
    cnt = bsobj.find('div', attrs={"id": "xw_box"}).find_all("p")
    s = ""
    if cnt:
        for item in cnt:
            s += item.text
        return s

    return "爬取失败！"


def saveFile(path, filename, content):
    if not os.path.exists(path):
        os.makedirs(path)
    # 保存文件
    with open(path + filename + ".txt", 'w', encoding='utf-8') as f:
        f.write(content)


# if "__main__" == __name__:
#     for url in getPageUrl():
#         s = fetchUrl(url)
#         for title, link, date in getTitleUrl(s):
#             print(title)
#             print(link)
#             print(date)
#             html = fetchUrl(link)
#             content = getContent(html)
#             print(content)
#             saveFile("D:/Python/Data_0913/", title, content)
#             print("-----" * 20)
def main():
    for url in getPageUrl():
        s = fetchUrl(url)
        for title, link, date in getTitleUrl(s):
            print(title)
            print(link)
            print(date)
            html = fetchUrl(link)
            content = getContent(html)
            print(content)
            saveFile("D:/Python/Data_0913/", title, content)
            print("-----" * 20)


# cProfile.run('re.compile("foo|bar")')

if __name__ == '__main__':
    import cProfile, pstats
    profiler = cProfile.Profile()
    profiler.enable()
    main()
    profiler.disable()
    stats = pstats.Stats(profiler).sort_stats('ncalls')
    stats.print_stats()
