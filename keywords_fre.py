#-*- coding:utf-8 -*-
import xlwings as xw
from input_data import dict_all,file_info
from datetime import datetime
import json

'''
## 1. 按不同代码、不同列及对应字典，统计某“关键词”出现频率
* 第一个需求是，根据data.xls中的B行[代码]统计不同字典中每个关键词出现的总频率，不同年份出现该词的频率
* 最好可自定义时间段。

####运行前说明
1、 dict_all 中的 dictn 填写示意
    'dict1' :{
        ## 选填项：填写指定的 excel 中指定的代码名作为筛选条件，不填则默认选择所有代码
        'code_name':'',   # 或者'code_name':'00001',

        # 选填项：以列表的形式写入需要统计词频的关键词，空列表则不对该列做统计
        'words':['公告及通告','月報表'],

        # 选填项：自定义时间段筛选，注意格式为 %Y-%M-%D ,不填则默认所有日期
        'from':'2015-01-01',
        'to':'2017-07-01',

        # 需求 3 中所需信息，本脚本不会引用
        'main_word':'月報表',
        'relative_words':[],

2、结果以如下形式
举例：以 '公告及通告' 及 '月報表' 对 D 列进行筛选
{
    'dict1': {
        '公告及通告': {
            '2017': 3,
            '2016': 17,
            '2015': 49,
            'all': 69
        },
        'code_name': '00001',
        'years': [
            2015,
            2016,
            2017
        ],
        '月報表': {
            '2017': 3,
            '2016': 12,
            '2015': 12,
            'all': 27
        }
    }
}

3、最后结果保存为 txt ,并打印输出
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
            if not pre_code_name == code_name:
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


def main():
    '''主函数'''
    file_path = file_info['filepath']
    cells = file_info['cells']
    sheet = file_info['sheet']
    value = open_file(file_path,cells,sheet=sheet)
    new_dict = check_input(dict_all)
    output_dict = distribute_dict(value,new_dict)

    with open('result1.txt','w')as f:
        f.write(json.dumps(output_dict))

    print(output_dict)
    return output_dict

if __name__ == '__main__':
    main()
