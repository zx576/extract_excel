#-*- coding:utf-8 -*-
import xlwings as xw
from input_data import dict_all,file_info
from datetime import datetime
import json
from collections import Counter

'''
## 1. 按不同代码、不同列及对应字典，统计某“关键词”出现频率
* 第一个需求是，根据data.xls中的B行[代码]统计不同字典中每个关键词出现的总频率，不同年份出现该词的频率
* 最好可自定义时间段。

####运行前说明
1、 dict_all 中的 dictn 填写示意
    'dict1' :{
        ## 选填项：填写指定的 excel 中指定的代码名作为筛选条件，不填则默认选择所有代码
        'code_name':'',   # 或者'code_name':['00001'],

        # 选填项：以列表的形式写入需要统计词频的关键词，空列表则不对该列做统计
        'words':['公告及通告','月報表'],

        # 选填项：自定义时间段筛选，注意格式为 %Y-%M-%D ,不填则默认所有日期
        'from':'2015-01-01',
        'to':'2017-07-01',

        # 需求 3 中所需信息，本脚本不会引用
        'main_word':'月報表',
        'relative_words':[],

2、结果保存在 excel 中

结果如下：

关键词	all	2017	2016	2015	2014	2013	2012	2011	2010	2009	2008	2007
月報表	99	3	12	12	12	12	12	12	12	12	0	0
关键词	all	2017	2016	2015	2014	2013	2012	2011	2010	2009	2008	2007
通函	16	0	2	6	2	2	2	2	0	0	0	0
关键词	all	2017	2016	2015	2014	2013	2012	2011	2010	2009	2008	2007
公告及通告	0	0	0	0	0	0	0	0	0	0	0	0


同时包括了所有排名

'''


def open_file(file_path,cells,sheet=0):
    '''打开文件，返回选定的数据
    file_path - 文件路径
    cells - 选中的单元格
    sheet - 工作簿名，默认为 0
    '''
    wb = xw.Book(file_path)
    sht = wb.sheets[sheet]
    value = sht.range(cells).value
    return value

# 处理 D 行的数据
def output_data(value,input_data,col_num):
    '''返回某一张字典下关键字的词频

    value - excel 中读取的数据
    input_data - 某个含关键字的字典，对应 dict1/dict2/dict3/dict4
    col_num - 该字典对应的 excel 列


    '''
    # 生成基本字典数据
    time_limit = True
    try:
        from_time = datetime.strptime(input_data['from'],'%Y-%m-%d')
        to_time = datetime.strptime(input_data['to'],'%Y-%m-%d')
    except:
        time_limit = False

    word_fre = {}
    for word in input_data['words']:
        word_fre[word] = {
            'all':0,
        }
        if time_limit:
            for year in range(from_time.year,to_time.year+1):
                word_fre[word][str(year)] = 0

    pre_code_name = input_data['code_name']
    # print(word_fre_D)

    for col in value:
        # 处理时间
        str_s = col[0]
        date = datetime.strptime(str_s,'%d/%m/%Y %H:%M')


        # 代码名
        code_name = col[1]


        # 筛选时间
        if time_limit:
            if not from_time < date < to_time:
                continue
        # 筛选代码名，不同直接跳过
        if pre_code_name:
            if not code_name in pre_code_name:
                continue
        # 计数
        # word_fre = {}
        for word in input_data['words']:
            if word in col[col_num]:
                word_fre[word]['all'] += 1
                if time_limit:
                    word_fre[word][str(date.year)] += 1
    word_fre['code_name'] = pre_code_name
    # if time_limit:
    #     word_fre['years'] = list(range(from_time.year,to_time.year+1))
    # print(word_fre)
    return word_fre


def check_input(dict_all):
    '''对 input 文件下 dict_all 字典做有效检测，返回有效的新字典

    dict_all - input_data.py 文件下 dict_all 字典


    '''
    new_dict = {}
    for key in dict_all.keys():
        if len(dict_all[key]['words']) > 0:
            new_dict[key] = dict_all[key]

    return new_dict

def distribute_dict(value,new_dict):
    '''分发解包有效性字典，以字典的形式返回最终的结果

    value - excel 文件中多有数据
    new_dict - 有效配置字典


    '''
    output_dict = {}
    for key in new_dict.keys():
        if key == 'dict1':
            col_num = 3
        elif key == 'dict2':
            # 2.标签-后	3.标签-前	4.题目
            col_num = 4
        elif key == 'dict3':
            col_num= 5
        elif key == 'dict4':
            col_num = 6
        result = output_data(value,new_dict[key],col_num)
        output_dict[key] = result

    return output_dict

def generate_excel(output_dict):
    # print(output_dict)
    xb = xw.Book()
    # 针对第一个要求
    data_dict = {}
    for key in output_dict.keys():
        xs = xb.sheets.add(name=key)
        r_list = []
        for inner_key in output_dict[key].keys():
            f_list = []
            if inner_key == 'code_name':
                continue

            ####

            data_dict[inner_key] = output_dict[key][inner_key]['all']

            ####
            f_list.append(['关键词',inner_key])
            for m_inner_key,m_inner_value in output_dict[key][inner_key].items():

                f_list.append([m_inner_key,m_inner_value])
                # all_list.append([m_inner_key,m_inner_value])
            f_list.sort(reverse=True)
            s_list = [[],[]]
            for i in f_list:
                s_list[0].append(i[0])
                s_list[1].append(i[1])
            for i in s_list:
                r_list.append(i)
        # print(r_list)
        xs.range('A1').value = r_list
    # print(data_dict)
    xs = xb.sheets.add(name='alldata')
    # c = Counter(data_dict)
    # c_3 = c.most_common(3)
    data_list = []
    for i,j in data_dict.items():
        data_list.append([i,j])
    data_list.sort(key=lambda x: x[1],reverse=True)

    xs.range('A1').value = data_list
    xb.save()

def main():
    '''主函数'''
    file_path = file_info['filepath']
    cells = file_info['cells']
    sheet = file_info['sheet']
    value = open_file(file_path,cells,sheet=sheet)
    new_dict = check_input(dict_all)
    output_dict = distribute_dict(value,new_dict)
    generate_excel(output_dict)
    # with open('result1.txt','w')as f:
    #     f.write(json.dumps(output_dict))

    # print(output_dict)
    return output_dict

if __name__ == '__main__':
    main()
