#-*- coding:utf-8 -*-
import xlwings as xw
from input_data import dict_all,file_info
from datetime import datetime

'''
## 2. 显示包含该“关键词”的那些行的信息（时间逆序排序）
* 最近的最先出现
* 希望可以自定义时间段和行信息。

####运行前说明
1、 dict_all 中的 dictn 填写示意
    'dict1' :{
        ## 选填项：填写指定的 excel 中指定的代码名作为筛选条件，不填则默认选择所有代码
        'code_name':'',   # 或者'code_name':'00001',

        # 选填项：以列表的形式写入需要筛选的关键词，空列表则不对该列做统计
        'words':['公告及通告','月報表'],

        # 选填项：自定义时间段筛选，注意格式为 %Y-%M-%D ,不填则默认所有日期
        'from':'2015-01-01',
        'to':'2017-07-01',

        # 需求 3 中所需信息，本脚本不会引用
        'main_word':'月報表',
        'relative_words':[],
2、最后结果保存在新建的 excel 表中,每个关键词占一个工作簿
第一行会显示  ####2017/3/12 19:52	月報表6#####
其中日期为运行脚本的时间，'月報表6' 为关键词+第6列数据，列数从 0 开始计算

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


def output_data(value,input_data,col_num):
    '''返回以某本字典下的关键词筛选出的所有信息

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


    pre_code_name = input_data['code_name']
    details = []
    word_list = {}
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
        if not pre_code_name == code_name and not pre_code_name:
            continue

        for word in input_data['words']:
            if not word in word_list.keys():
                word_list[word] = [[datetime.now(),word+str(col_num),None,None,None,None,None,None,None],]
            if word in col[col_num]:
                word_list[word].append(col)
    # print(word_list)
    # print('================')
    return word_list

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
    '''分发解包有效性字典，以列表的形式返回最终的结果

    value - excel 文件中多有数据
    new_dict - 有效配置字典


    '''
    output_list = []
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
        for li in result.keys():

            output_list.append(result[li])
            # print('================')
    # print(output_list)
    # print('================')
    # print(len(output_list))
    return output_list

def generate_excel(output_list):
    '''读入筛选的结果，存入新的 excel 表，每个关键字新建一张工作簿存入

    output_list - 筛选后的结果

    '''
    xb = xw.Book()
    # sheet = 0
    for output_one in output_list:
        # print(output_one)
        # xs = xb.sheets[sheet]
        xs = xb.sheets.add()
        output_one = sorted(output_one,key=lambda x:sort_list(x[0]),reverse=True)
        xs.range('A1').value = output_one
    xb.save()

def sort_list(string):
    '''对最终的结果按日期进行排序

    string - excel 表中第一列日期数据，如 '23/02/2017 20:23'

    '''
    date = datetime.now()
    try:
        date = datetime.strptime(string,'%d/%m/%Y %H:%M')
    except:
        pass
    return date
def main():
    '''主函数'''
    file_path = file_info['filepath']
    cells = file_info['cells']
    sheet = file_info['sheet']
    value = open_file(file_path,cells,sheet=sheet)

    new_dict = check_input(dict_all)

    output_list = distribute_dict(value,new_dict)
    generate_excel(output_list)
    # print(output_dict)
    return output_data

if __name__ == '__main__':
    main()
