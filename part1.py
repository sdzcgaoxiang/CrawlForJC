import json
import os
from time import sleep
from urllib import parse
import xlwings as xw
import requests
import random
import re

# 正则模板，匹配年度报告年份
pattern = re.compile(r'.*(\d{4})年半年度报告.*$', re.I)

USER_AGENT_LIST = [
    'MSIE (MSIE 6.0; X11; Linux; i686) Opera 7.23',
    'Opera/9.20 (Macintosh; Intel Mac OS X; U; en)',
    'Opera/9.0 (Macintosh; PPC Mac OS X; U; en)',
    'iTunes/9.0.3 (Macintosh; U; Intel Mac OS X 10_6_2; en-ca)',
    'Mozilla/4.76 [en_jp] (X11; U; SunOS 5.8 sun4u)',
    'iTunes/4.2 (Macintosh; U; PPC Mac OS X 10.2)',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:5.0) Gecko/20100101 Firefox/5.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:9.0) Gecko/20100101 Firefox/9.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:16.0) Gecko/20120813 Firefox/16.0',
    'Mozilla/4.77 [en] (X11; I; IRIX;64 6.5 IP30)',
    'Mozilla/4.8 [en] (X11; U; SunOS; 5.7 sun4u)'
]


def read_excel(excel_name):
    app = xw.App(visible=True, add_book=False) # 打开一个处理Excel的App应用
    wb = app.books.open(excel_name)             # 打开一个excel工作簿
    sht = wb.sheets["Sheet1"]                   # 打开excel里面的第一个表
    a = sht.range('C2:C105').value              # 取出c2到c105的数据，并且保存到a中
    return a                                    # 返回a


def get_adress(bank_name):
    url = "http://www.cninfo.com.cn/new/information/topSearch/detailOfQuery"
    data = {
        'keyWord': bank_name,
        'maxSecNum': 10,
        'maxListNum': 5,
    }
    hd = {
        'Host': 'www.cninfo.com.cn',
        'Origin': 'http://www.cninfo.com.cn',
        'Pragma': 'no-cache',
        'Accept-Encoding': 'gzip,deflate',
        'Connection': 'keep-alive',
        'Content-Length': '70',
        #'User-Agent': 'Mozilla/5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 75.0.3770.100Safari / 537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Accept': 'application/json,text/plain,*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    }
    # 随机生成user agent
    try:
        hd.setdefault('User-Agent', random.choice(USER_AGENT_LIST))
        r = requests.post(url, headers=hd, data=data)
        print(r.text)
        r = r.content
        m = str(r, encoding="utf-8")
        pk = json.loads(m)
        orgId = pk["keyBoardList"][0]["orgId"]  # 获取参数
        plate = pk["keyBoardList"][0]["plate"]
        code = pk["keyBoardList"][0]["code"]
        print(orgId, plate, code)
    except Exception as e:
        print("获取url出错")
        return [0,0,0]
    return orgId, plate, code


def download_PDF(url, file_name):  # 下载pdf
    url = url
    r = requests.get(url)
    f = open(bank + "/" + file_name + ".pdf", "wb")
    f.write(r.content)


def get_PDF(orgId, plate, code):
    url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
    if code[0] == '8':        #8开头股票代码结构
        data = {
            'stock': '{},{}'.format(code, orgId),
            'tabName': 'fulltext',
            'pageSize': 30,
            'pageNum': 1,
            'column': 'third',
            'category': 'category_dqgg;',
            'plate': '',
            'seDate': '',
            'searchkey': '',
            'secid': '',
            'sortName': '',
            'sortType': '',
            'isHLtitle': 'true',
        }
    else:
        data = {
            'stock': '{},{}'.format(code, orgId),
            'tabName': 'fulltext',
            'pageSize': 30,
            'pageNum': 1,
            'column': plate,
            'category': 'category_ndbg_szsh;',
            'plate': '',
            'seDate': '',
            'searchkey': '',
            'secid': '',
            'sortName': '',
            'sortType': '',
            'isHLtitle': 'true',
        }

    hd = {
        'Host': 'www.cninfo.com.cn',
        'Origin': 'http://www.cninfo.com.cn',
        'Pragma': 'no-cache',
        'Accept-Encoding': 'gzip,deflate',
        'Connection': 'keep-alive',
        # 'Content-Length': '216',
        #'User-Agent': 'User-Agent:Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Accept': 'application/json,text/plain,*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'X-Requested-With': 'XMLHttpRequest',
        # 'Cookie': cookies
    }
    # 随机生成user agent
    hd.setdefault('User-Agent', 'User-Agent:' + random.choice(USER_AGENT_LIST))

    data = parse.urlencode(data)
    print(data)
    r = requests.post(url, headers=hd, data=data)
    print(r.text)
    r = str(r.content, encoding="utf-8")
    r = json.loads(r)
    reports_list = r['announcements']
    # 年报
    try:
        for report in reports_list:
            if '摘要' in report['announcementTitle'] or "20" not in report['announcementTitle']:       #假如要爬题目中含有别的信息的年报 修改这一部分
                continue
            if 'H' in report['announcementTitle']:
                continue
            if '半年度' in report['announcementTitle'] and '2020年半年度' not in report['announcementTitle'] :
                continue
            else:  # http://static.cninfo.com.cn/finalpage/2019-03-29/1205958883.PDF
                pdf_url = "http://static.cninfo.com.cn/" + report['adjunctUrl']
                file_name = report['announcementTitle']
                print("正在下载：" + pdf_url, "存放在当前目录：/" + bank + "/" + file_name)
                download_PDF(pdf_url, file_name)
                sleep(2)
    except Exception as e:
        print(e)

    # 如果开头代码不为8，则需要爬取2020半年报
    years = {'2020'}            # 爬取年份限定
    if code[0] != '8':          # 8开头股票代码结构
        data = {
            'stock': '{},{}'.format(code, orgId),
            'tabName': 'fulltext',
            'pageSize': 30,
            'pageNum': 1,
            'column': plate,
            'category': 'category_bndbg_szsh;',
            'plate': '',
            'seDate': '',
            'searchkey': '',
            'secid': '',
            'sortName': '',
            'sortType': '',
            'isHLtitle': 'true',
        }
    hd.setdefault('User-Agent', 'User-Agent:' + random.choice(USER_AGENT_LIST))

    data = parse.urlencode(data)
    print(data)
    r = requests.post(url, headers=hd, data=data)
    print(r.text)
    r = str(r.content, encoding="utf-8")
    r = json.loads(r)
    reports_list = r['announcements']
    try:
        for report in reports_list:
            m = pattern.match(report['announcementTitle'])
            if '摘要' in report['announcementTitle'] or "20" not in report['announcementTitle']:       #假如要爬题目中含有别的信息的年报 修改这一部分
                continue
            if 'H' in report['announcementTitle']:
                continue
            # 如果该半年报不是所要求的半年报，则跳过
            if m[1] not in years:
                continue
            else:  # http://static.cninfo.com.cn/finalpage/2019-03-29/1205958883.PDF
                pdf_url = "http://static.cninfo.com.cn/" + report['adjunctUrl']
                file_name = report['announcementTitle']
                print("正在下载：" + pdf_url, "存放在当前目录：/" + bank + "/" + file_name)
                download_PDF(pdf_url, file_name)
                sleep(1)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    bank_list = ['龙的电器', ]        #单独测试
    #bank_list = read_excel(r'data.xlsx')

    for bank in bank_list:
      try:
        os.mkdir(bank)                         #创建名称为bank的文件夹
        orgId, plate, code = get_adress(bank)
        get_PDF(orgId, plate, code)
        print("下一家~")
      except Exception as e:
          print("出现错误错误编号为e")
          print(e)
    print("All done!")



