import os
import xlwings as xw
import re
import PDFAnalyse as trans

ModeSet = ["字词版", "年份版"]
Leastyear = 2010        # 起始年份
Latestyear = 2022       # 结束年份
Mode = ModeSet[1]       # 改模式

# excel新建操作
app = xw.App(visible=True, add_book=False)
dataAna = app.books.open('Analyse.xlsx')  #打开Analyse.xlsx'
sht = dataAna.sheets['sheet1']
assert isinstance(sht, xw.Sheet)  # 强制类型声明（）


# -------------------------改1-------------------------------------
if Mode == ModeSet[0]:
    # 横坐标读入（关键词）
    i = 2
    for word in trans.a:
        sht.range((1, i)).value = word  # 生成关键字坐标
        i += 1                          # 方便生成坐标
else:
    #横坐标读入（年份）
    i = 2
    for year in range(Leastyear, Latestyear + 1):
        sht.range((1, i)).value = year  # 生成关键字坐标
        i += 1                          # 方便生成坐标

# -------------------------改1-------------------------------------

# 目录数据读入
i = 2
names = [] #公司名称数组 后期做统计用

# 按列读入公司文件夹名称
for name in os.listdir(os.getcwd()):    #把所有的文件列出来，从当前目录文件夹取
    if os.path.isdir(name) & (name != '.idea') & (name != '__pycache__'):    #排除掉配置文件夹和普通文件
        # sht.range((i, 1)).value = name 不重要下面一行
        names.append(name)
        # 正则匹配文件夹内所有pdf文件名 忽略大小写 提取年份 写入pdf 调用trans
        pdfs = os.listdir(name)                                              # 获取文件夹内所有pdf文件
        pattern = re.compile(r'.*20(\d{2})年年度报告.*\.pdf$', re.I)            #正则表达式
        for pdf in pdfs:
            if pattern.match(pdf):                                          #匹配年度报告，是年度报告才进行读取
                m = pattern.match(pdf)
                year = int(m[1])+2000                                       #m[1]时提出的年份计算年份

                if year < 2009:                                             # 2009年前，跳过这一个年报pdf
                    continue

                # -------------------------改2-------------------------------------
                if Mode == ModeSet[0]:
                    sht.range((i, 1)).value = "{}{}年年报".format(name, year)      #excel写入年报名称，每行一个年份
                else:
                    sht.range((i, 1)).value = "{}年报".format(name)                   #每行一个公司
                # -------------------------改2-------------------------------------

                print("正在扫描{}{}年年报…………".format(name, year))
                word = name + '\\' + pdf                                    # 获取路径
                word_dict = trans.findTecWords(name + '\\' + pdf)           # 获取字典，调用分析的模块，对pdf进行关键词词频分析

                # -------------------------改3-------------------------------------
                if Mode == ModeSet[0]:
                    #原先的
                    j = 2
                    for word, value in word_dict.items():                           #把统计完的数据写入excel
                        sht.range((i, j)).value = value
                        j += 1                                          # 列+1
                else:
                    #2010当年总数
                    sum = 0
                    for word, value in word_dict.items():                           #把统计完的数据写入excel
                        # if value != 0:
                        #     sum = sum + 1
                        sum = sum + value
                    j = year - Leastyear + 2
                    sht.range((i, j)).value = sum
                # -------------------------改3-------------------------------------


                print(word)
                # -------------------------改4-------------------------------------
                if Mode == ModeSet[0]:
                    i += 1  # 每扫描一个pdf行+1
        if Mode == ModeSet[1]:
            i += 1                                              # 每扫描一个公司文件夹行+1
        # -------------------------改4-------------------------------------
