#-*- coding:utf-8 -*-

'''
做数据输入用，在此配置信息

示例
配置 excel 信息
file_info = {
    # 文件路径
    'filepath':r'data1.xlsx',
    # excel 表名 可为工作簿名如 'sheet1' 也可直接为工作簿顺序 从 0 开始计数，默认为 0
    'sheet':0,
    # 单元格区域
    'cells':'A2:I413'
}

配置筛选信息

'dict1' :{
    ## 选填项：填写指定的 excel 中指定的代码名作为筛选条件，不填则默认选择所有代码
    'code_name':'',   # 或者'code_name':'00001',

    # 选填项：以列表的形式写入需要统计词频的关键词，空列表则不对该列做统计
    'words':['公告及通告','月報表'],

    # 选填项：自定义时间段筛选，注意格式为 %Y-%M-%D ,不填则默认所有日期
    'from':'2015-01-01',
    'to':'2017-07-01',

    # 填写关键词，仅包含一个，
    'main_word':'月報表',
    # 填写关联关键词，可包含多个
    'relative_words':[],

'''
file_info = {
    'filepath':r'data1.xlsx',  # 文件路径
    'sheet':0,   # excel 表名 可为工作簿名如 'sheet1' 也可直接为工作簿顺序 从 0 开始计数，默认为 0
    'cells':'A2:I413' # 单元格区域
}

dict_all ={
    'dict1' :{
        'code_name':'',
        'words':['公告及通告','月報表'],
        'from':'',
        'to':'2017-07-01',
        'main_word':'公告及通告',
        'relative_words':[],
    },
    'dict2' :{
        'code_name':'00001',
        'words':['關連交易','其他'],
        'from':'2015-01-01',
        'to':'2017-07-01',
        'main_word':'',
        'relative_words':['關連交易'],
    },
    'dict3' :{
        'code_name':'00001',
        'words':['公告及通告'],
        'from':'2007-01-01',
        'to':'2017-07-01',
        'main_word':'',
        'relative_words':['公告及通告','月報表'],
    },
    'dict4' :{
        'code_name':'00001',
        'words':['通函','公告及通告','月報表'],
        'from':'2007-01-01',
        'to':'2017-07-01',
        'main_word':'',
        'relative_words':['報表'],
    },
        }
