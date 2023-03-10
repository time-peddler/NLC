import re,json,os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def NLC_parse(data,save_path='./guotu_from_sav.xlsx'):
    
    # 把数据转化为列表并删除表头表尾
    datas=data.split('\n\n')[2:-4]


    def DictToStr(list):
        if len(list) == 0:
            return None
        elif len(list) ==1:
            return list[0]
        else:
            toStr = ''
            for ele in range(0, len(list)):
                toStr = toStr + list[ele] + ';'
            return toStr

    result = []
    for data in datas:
        temp = {}
        temp["index"] = re.findall('记录号\W*(\w+)',data,re.DOTALL)[0]
        temp["title"] = re.findall('题名责任者项\s*([\S 　]*?)\[',data,re.DOTALL)[0]

        # category
        category = re.findall('\[\W*([\u4e00-\u9fa5\WA-Z]{2,}?)\]',data,re.DOTALL)
        temp["category"] = DictToStr(category)

        # CLC
        CLC = re.findall('中图分类号\s*(.+)',data,re.MULTILINE)
        temp["ClC"] = DictToStr(CLC)

        # publisher
        publisher = re.findall('[\u4e00-\u9fa5]{2,}\W*[:：]\W*[\u4e00-\u9fa5]{2,}[社出版公司]',data,re.DOTALL)
        temp["publisher"] = DictToStr(publisher)

        # key
        key = re.findall('主题\w{0,1}\s{2,}(.+)', data, re.MULTILINE)
        temp["key"] = DictToStr(key)

        # author
        author = re.findall('个人名称[等同次要]*\s*(.+)', data, re.MULTILINE)
        temp["author"] = DictToStr(author)

        # time
        time = re.findall('出版发行项.+?(19\d\d|20\d\d|2100)',data,re.MULTILINE)
        temp["time"] = DictToStr(time)
        result.append(temp)

    data = json.dumps(result)

    df1 = pd.read_json(data, encoding='utf-8')
    path = save_path
    isExists = os.path.exists(path)
    if not isExists:
        df1.to_excel(path, 'GUOTU')
    else:
        df2 = pd.read_excel(path, 'GUOTU', index_col=0)
        df = pd.concat([df1,df2])
        df.to_excel(path, 'GUOTU')
