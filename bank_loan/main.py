# -*- coding: utf-8 -*-
# @Time : 2020-03-23 22:53
# @Author : icarusyu
# @FileName: main.py
# @Software: PyCharm
import pandas as pd
import os
from collections import defaultdict
import xlrd,openpyxl
from datetime import datetime

dic_type1 = {'0':'个人借','1':'农户借','2':'农户借'}
dic_type2 = {'0':'信用','1':'抵押','2':'贷款'}
# md
# 为"证据目录"等文件改名没有意义，因为是word
# "民事起诉状"中本来有涉及金额的复制粘贴部分，因为生成的是无格式的txt,拼出所有文字后还需要重调格式，建议生成后直接粘贴到留白的word中
# 依赖包：
# pandas,xlrd
# Q
# 当时说的需要计算以及给定的列，其并集不等于原"贷款利息计算表"中的所有列
# "贷款利息计算表"不明确，如果一个group有多个人，该表如何填写？目前按照group中第一人的信息进行填写。
# Attention
# 不要多次运行本程序，否则文件中会生成追加内容
# 将personal_info，和输出的文件从git版本管理中移除。否则所有信息将上传到网上公开


def create_folder(name,type1, type2, path='bank_loan/output/'):
    folder = path+ name + '(' + dic_type1[type1] + '、' + dic_type2[type2] + ')'
    is_exist = os.path.exists(folder)
    if not is_exist:
        os.makedirs(folder)
    return folder

def count_interest(folder, df): # first_id即info.lsx中"首位被告人身份证号"
    dic = defaultdict(int)
    dic['借款人'] = df.ix[0, '首位被告人']
    dic['本金']= df.ix[0,'本金']
    dic['借款起止期间0'] = df.ix[0,'借款起止期间0']
    dic['借款起止期间1'] = df.ix[0,'借款起止期间1']
    dic['月利率（‰）'] = df.ix[0,'月利率']
    dic['已清息期间0'] = df.ix[0, '已清息期间0']
    dic['已清息期间1'] = df.ix[0, '已清息期间1']
    dic['已清息'] = df.ix[0,'已清息']
    dic['欠息期间1'] = df.ix[0,'欠息期间1']
    dic['罚息期间1'] = df.ix[0,'罚息期间1']
    dic['逾期利率9.3×1.5'] = df.ix[0,'逾期利率']

    # counting
    # todo
    dic['罚息期间0'] = 0
    dic['欠息期间0'] = 0
    dic['欠息'] = 0
    dic['罚息'] = 0

    dic['欠息合计'] = dic['欠息'] +dic['罚息']
    dic['本息合计'] = dic['本金'] + dic['欠息合计']

    df_out = pd.DataFrame(dic,index=[0]) # md
    df_out.to_excel(folder + '/贷款利息计算表.xlsx')
    return df_out

def write_complaint(df,df_loan, folder):
    # 个人信息from表，由贷款利息表所得的第二部分内容，并把第二部分内容复制到下面的位置
    one_group = []
    for i in range(len(df)):
        person = []
        person.append('  被告：'+ df.ix[i,'被告人'])
        person.append('男' if int(df.ix[i,'身份证号'].astype(str)[-2])%2 else '女')
        birth = datetime.strptime(df.ix[i,'身份证号'].astype(str)[6:14],'%Y%m%d')
        person.append(str(birth.year) + '年' + str(birth.month) + '月'+ str(birth.day) +'日生')
        person.append('汉族' if df.ix[i,'民族'] ==-1 else df.ix[i,'民族'])
        person.append('户籍地' + df.ix[i,'户籍地'])
        person.append('身份证号' + df.ix[i,'身份证号'].astype(str))
        person.append('电话'+ df.ix[i,'电话'].astype(str))
        one_group.append('，'.join(person) + '。')

    str0 = '\n'.join(one_group)

    # todo 数据格式，eg.万元/元
    lst = []
    lst.append('判决被告向原告偿还借款本金'+ df_loan.ix[0,'本金'].astype(str) + '元、')
    lst.append('利息' + df_loan.ix[0,'欠息'].astype(str) + '元、')
    lst.append('罚息' + df_loan.ix[0,'罚息'].astype(str) + '元')
    t = datetime.strptime(df_loan.ix[0,'欠息期间1'].astype(str),'%Y%m%d')
    last_date = str(t.year) + '年' + str(t.month) + '月' + str(t.day) + '日'
    lst.append('（利息暂计算至'+ last_date +'止，以后的利息按照《个人借款合同》约定利率及罚息标准算至借款还清之日止），')
    lst.append('合计'+ df_loan.ix[0,'本息合计'].astype(str) +'元。')
    str1 = ''.join(lst)
    # print(str1)

    with open(folder + '/' + df_loan.ix[0,'借款人'] + '诉状','w+') as file:
        file.write(str0 + '\n' + str1)
    file.close()

def run_():
    df = pd.read_excel('bank_loan/personal_info.xlsx', '工作表1')
    n = df['group'].max()
    for i in range(n+1):
        frac = df[df['group']==i]
        folder = create_folder(frac.ix[0,'首位被告人'], frac.ix[0,'type1'].astype(str), frac.ix[0,'type2'].astype(str))
        df_loan = count_interest(folder,frac)
        write_complaint(frac,df_loan, folder)

if __name__ == '__main__':
    run_()

