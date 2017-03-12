#-*- coding:utf-8 -*-
import xlwings as xw
from input_data import dict_all,file_info
from datetime import datetime
import json
from collections import Counter

'''
## 3. 关联关键词
* 出现该“关键词”的条件下，希望找到有哪些关联关键词同时出现？列出出现频率排名前3的词及出现频率。
* 希望可以自定义检索4列，即4本不同的字典

####运行前说明
1、 dict_all 中的 dictn 填写示意
    'dict1' :{
        ## 选填项：填写指定的 excel 中指定的代码名作为筛选条件，不填则默认选择所有代码
        'code_name':'',   # 或者'code_name':'00001',

        # 需求 1 中所需信息，本脚本不会引用
        'words':['公告及通告','月報表'],

        # 选填项：自定义时间段筛选，注意格式为 %Y-%M-%D ,不填则默认所有日期
        'from':'2015-01-01',
        'to':'2017-07-01',

        # 填写关键词，仅包含一个，
        'main_word':'月報表',
        # 填写关联关键词，可包含多个
        'relative_words':[],

2、输出在 txt 文件中
输出结果如下：(实例)
2.1、所有的结果
{
    'dict2': {
        '關連交易': 32
    },
    'dict3': {
        '月報表': 0,
        '公告及通告': 167
    },
    'dict4': {
        '報表': 0
    }
}
2.2、有效词汇，即词频不为 0 的词汇
[('關連交易', 32), ('公告及通告', 167)]

2.3、排名前三的词汇
[('公告及通告', 167), ('關連交易', 32)]

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

def check_input(dict_all):
    '''对 input 文件下 dict_all 字典做有效检测，返回有效的新字典

    dict_all - input_data.py 文件下 dict_all 字典


    '''
    new_dict = {'main_word':{}}
    for key in dict_all.keys():
        # 提取主要关键词
        inner_dct = dict_all[key]
        # print(inner_dct)
        if inner_dct['main_word'] != '':
            new_dict['main_word']['word'] = inner_dct['main_word']
            if key == 'dict1':
                col_num = 3
            elif key == 'dict2':
                # 2.标签-后	3.标签-前	4.题目
                col_num = 4
            elif key == 'dict3':
                col_num= 5
            elif key == 'dict4':
                col_num = 6
            new_dict['main_word']['col'] = col_num
            new_dict['main_word']['code_name'] = inner_dct['code_name']
            new_dict['main_word']['from'] = inner_dct['from']
            new_dict['main_word']['to'] = inner_dct['to']

        if len(inner_dct['relative_words']) > 0:
            new_dict[key] = inner_dct['relative_words']
    # print(new_dict)
    return new_dict

# check_input(dict_all)

# 以主要关键字筛选信息
def output_data(value,dict):
    '''以主要关键字筛选所有信息，返回符合条件的数据

    value - 所有 excel 信息
    dict - 包含主要关键字的字典

    '''
    # 生成基本字典数据
    time_limit = True
    try:
        from_time = datetime.strptime(dict['from'],'%Y-%m-%d')
        to_time = datetime.strptime(dict['to'],'%Y-%m-%d')
    except:
        time_limit = False

    pre_code_name = dict['code_name']
    main_word = dict['word']
    col_num = dict['col']

    f_list = []

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
            if not pre_code_name == code_name:
                continue

        if main_word in col[col_num]:
            f_list.append(col)
    # print(len(f_list))
    return f_list

def output_data_2(f_list,word_list,col_num):
    '''从主要关键词筛选出的信息中，对关联关键词做词频统计，返回统计数据

    f_list - 主要关键词筛选出的信息
    word_list - 某一列需要筛选的关联关键词信息
    col_num - 关联关键词所在列数

    '''
    # 做数据容器
    word_dict = {}
    for word in word_list:
        word_dict[word] = 0

    for col in f_list:
        for word in word_list:
            if word in col[col_num]:
                word_dict[word] += 1

    return word_dict

def distribute_dict(value,dict):
    '''载入所有 excel 信息，和有效的设置信息，返回最终的结果

    value - 所有 excel 信息
    dict - input_data.py 文件中的 dict_all

    '''

    main_word_dict = dict['main_word']
    # print(main_word_dict)
    # 进行关键词筛选
    f_list = output_data(value,main_word_dict)

    result = {}
    for key in dict.keys():
        # result[key] = {}
        col_num = 0
        if key == 'main_word':
            continue
        if key == 'dict1':
            col_num = 3
        elif key == 'dict2':
            col_num = 4
        elif key == 'dict3':
            col_num= 5
        elif key == 'dict4':
            col_num = 6
        word_dict = output_data_2(f_list,dict[key],col_num)
        result[key] = word_dict


    return result

def handle_result(result):
    '''从最终的结果中，整理出有效的关联关键词，并提取出词频排名前三的词汇'''
    relative_words = []
    for key in result.keys():
        for inner_key in result[key].keys():
            if result[key][inner_key] > 0:
                element = (inner_key,result[key][inner_key])
                relative_words.append(element)
    c = Counter(dict(relative_words))
    common_c = c.most_common(3)

    return relative_words,common_c


def main():
    '''主函数'''
    # 获取 excel 数据
    file_path = file_info['filepath']
    cells = file_info['cells']
    sheet = file_info['sheet']
    value = open_file(file_path,cells,sheet=sheet)
    # 检查并输出配置文件
    new_dict = check_input(dict_all)
    # 得到结果
    result = distribute_dict(value,new_dict)
    # 写入文件
    relative_words,common_c = handle_result(result)
    print(relative_words,common_c)
    print(result)
    with open('result3.txt','w')as f:
        f.write(json.dumps(result))
        f.write('\n')
        f.write(str(relative_words))
        f.write('\n')
        f.write(str(common_c))

    return output_data
if __name__ == '__main__':
    main()
