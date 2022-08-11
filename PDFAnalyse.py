import pdfplumber
import jieba
import os

a = {"5G", "AI", "IoT", "智能制造", "智慧办公", "智能运营", "PaaS", "生态合作", "数字化", "智能化", "人工智能", "商业智能", "智能数据分析", "智能机器人", "机器学习",
     "深度学习", "语音识别", "身份验证", "大数据", "虚拟现实", "云计算", "智能安全", "物联网", "区块链", "工业互联网", "移动互联", "电子商务", "线上支付", "第三方平台",
     "电商平台", "智能客服", "智能家居", "智能营销", "数字营销", "无人零售", "集成电路"}


def findTecWords(pdf_name):
    word_dict = dict()
    for word in a:
        word_dict[word] = 0

    count = 0
    with pdfplumber.open(pdf_name) as pdf:
        page = pdf.pages[1]  # 第一页的信息
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
        l = jieba.cut(text)
        i = 0

        for word in l:
            i += 1
            if word in word_dict:
                word_dict[word] += 1
            # if word in a:
            #     count += 1
        print(word_dict)
    return word_dict

if __name__ == '__main__':
    print(os.getcwd())
    findTecWords(r"达伦股份\2020年年度报告.pdf")
