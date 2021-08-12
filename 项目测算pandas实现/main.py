import pandas as pd
import numpy as np
import datetime
# 读取费率表数据和日期、回款数据
df_fee = pd.read_excel('fees.xlsx',header =2, sheet_name = 'fees')
df = pd.read_excel('fees.xlsx', sheet_name = 'dates')
# 计算计息期间
df['计算期间'] = df['计算日'].diff(1).map(lambda x:x.days)
df['计息期间'] = df['计息日'].diff(1).map(lambda x:x.days)
# 计算增值税及附加
df['增值税'] = df['利息']/1.03*0.03
df['附加税'] = df['增值税']*0.12
# 计算过程中需要用到的计算量
df['累计还本'] = df['本金'].cumsum()
df['期末本金'] = df['本金'].sum()-df['累计还本']
df['期初本金'] = df['期末本金']+df['本金']
# 计算费用
df['托管费'] = df['期初本金'] * df_fee.at[0,'费率'] * df['计息期间'] / df_fee.at[27,'费率']
df['信托报酬'] = df['期初本金'] * df_fee.at[1,'费率'] * df['计息期间'] / df_fee.at[27,'费率']
df['资产服务费'] = df['期初本金'] * df_fee.at[3,'费率'] * df['计息期间'] / df_fee.at[27,'费率'] # 本例采用工行对公项目作为例子，本例指的是后端服务费
# 添加一次性费用，本例支付日比较少，手动添加了，后续遇到其它项目支付日比较多的，再考虑使用pandas日期函数加where函数实现if-then效果
df['一次性费用'] = [0,df_fee.at[9,'费率'],0,0,0,0]
df['前端资产服务费'] = [0,df_fee.at[2,'金额'],0,0,0,0]
# 费用小计
df['税费汇总'] = df['增值税'] + df['附加税'] + df['托管费'] + df['信托报酬'] + df['资产服务费'] + df['一次性费用'] +df['前端资产服务费']
# 初始化需要用到的中间计算量，采用了np.where函数，实现if-then-else效果
df['期初优先A本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[16,'备注'],0)
df['期初优先B本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[17,'备注'],0)
df['期初次级本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[20,'备注'],0)
df['期末优先A本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[16,'备注'],0)
df['期末优先B本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[17,'备注'],0)
df['期末次级本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[20,'备注'],0)
# 第一遍计算逻辑
df['本期优先A利息'] = df['期初优先A本金']*df_fee.at[4,'费率']* df['计息期间'] / df_fee.at[27,'费率']
df['本期优先B利息'] = df['期初优先B本金']*df_fee.at[5,'费率']* df['计息期间'] / df_fee.at[27,'费率']
df['本期次级期间收益'] = df['期初次级本金']*df_fee.at[8,'费率']* df['计息期间'] / df_fee.at[27,'费率']

df['当期利息分配总额'] = df['本期优先A利息']+df['本期优先B利息']+df['本期次级期间收益']

df['剩余收益转入本金帐'] = df['利息'] - df['当期利息分配总额'] - df['税费汇总'] #正值向本金转账，负值本金向收益帐转账
df['本期可分配本金'] = df['本金'] + df['剩余收益转入本金帐']

df['本期优先A本金分配'] = np.where(df['本期可分配本金'] >= df['期初优先A本金'], df['期初优先A本金'], df['本期可分配本金'])
df['期末优先A本金'] = df['期初优先A本金'] - df['本期优先A本金分配']
df['期初优先A本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[16,'备注'], df['期末优先A本金'].shift(1))

df['本期优先B本金分配'] = np.where(df['本金'] + df['剩余收益转入本金帐'] < df['期初优先A本金'] + df['期初优先B本金'], 
                           df['本金'] + df['剩余收益转入本金帐'] - df['本期优先A本金分配'], 
                           np.where(df['计息日'] == df_fee.at[26,'费率'],df['期初优先B本金'],df['期末优先B本金'].shift(1)))
df['期末优先B本金'] = df['期初优先B本金'] - df['本期优先B本金分配']
df['期初优先B本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[17,'备注'], df['期末优先B本金'].shift(1))

df['本期次级本金分配'] = np.where(df['本金'] + df['剩余收益转入本金帐'] < df['期初优先A本金'] + df['期初优先B本金'] + df['期初次级本金'], 
                          df['本金'] + df['剩余收益转入本金帐'] - df['本期优先A本金分配'] - df['本期优先B本金分配'], 
                          np.where(df['计息日'] == df_fee.at[26,'费率'],df['期初次级本金'],df['期末次级本金'].shift(1)))
df['期末次级本金'] = df['期初次级本金'] - df['本期次级本金分配']
df['期初次级本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[20,'备注'], df['期末次级本金'].shift(1))
# 循环计算叠加
i = df.shape[0] # 读取一共有多少行数据
# 循环计算，每行计算一遍
while i:   

    df['本期优先A利息'] = df['期初优先A本金']*df_fee.at[4,'费率']* df['计息期间'] / df_fee.at[27,'费率']
    df['本期优先B利息'] = df['期初优先B本金']*df_fee.at[5,'费率']* df['计息期间'] / df_fee.at[27,'费率']
    df['本期次级期间收益'] = df['期初次级本金']*df_fee.at[8,'费率']* df['计息期间'] / df_fee.at[27,'费率']

    df['当期利息分配总额'] = df['本期优先A利息']+df['本期优先B利息']+df['本期次级期间收益']

    df['剩余收益转入本金帐'] = df['利息'] - df['当期利息分配总额'] - df['税费汇总'] #正值向本金转账，负值本金向收益帐转账
    df['本期可分配本金'] = df['本金'] + df['剩余收益转入本金帐']

    df['本期优先A本金分配'] = np.where(df['本期可分配本金'] >= df['期初优先A本金'], df['期初优先A本金'], df['本期可分配本金'])
    df['本期优先B本金分配'] = np.where(df['本金'] + df['剩余收益转入本金帐'] < df['期初优先A本金'] + df['期初优先B本金'], 
                               df['本金'] + df['剩余收益转入本金帐'] - df['本期优先A本金分配'], 
                               np.where(df['计息日'] == df_fee.at[26,'费率'],df['期初优先B本金'],df['期末优先B本金'].shift(1)))
    df['本期次级本金分配'] = np.where(df['本金'] + df['剩余收益转入本金帐'] < df['期初优先A本金'] + df['期初优先B本金'] + df['期初次级本金'], 
                              df['本金'] + df['剩余收益转入本金帐'] - df['本期优先A本金分配'] - df['本期优先B本金分配'], 
                              np.where(df['计息日'] == df_fee.at[26,'费率'],df['期初次级本金'],df['期末次级本金'].shift(1)))

    df['期末优先A本金'] = df['期初优先A本金'] - df['本期优先A本金分配']
    df['期末优先B本金'] = df['期初优先B本金'] - df['本期优先B本金分配']
    df['期末次级本金'] = df['期初次级本金'] - df['本期次级本金分配']

    df['期初优先A本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[16,'备注'], df['期末优先A本金'].shift(1))
    df['期初优先B本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[17,'备注'], df['期末优先B本金'].shift(1))
    df['期初次级本金'] = np.where(df['计息日'] == df_fee.at[26,'费率'],df_fee.at[20,'备注'], df['期末次级本金'].shift(1))
    
    i -= 1
# 次级分配完之后的超额收益
df['超额收益'] = df['本期可分配本金'] - df['本期优先A本金分配'] - df['本期优先B本金分配'] - df['本期次级本金分配']

# 各层级的现金流情况
# 优先A级现金流情况
df['A'] = np.where(df['计息日'] == df_fee.at[24,'费率'], -df_fee.at[16,'备注'],df['本期优先A利息'] +df['本期优先A本金分配'])
# df[['计息日','A']]
# 优先B级现金流情况
df['B'] = np.where(df['计息日'] == df_fee.at[24,'费率'], -df_fee.at[17,'备注'],df['本期优先B利息'] + df['本期优先B本金分配'])
# df[['计息日','B']]
# 次级现金流情况
df['次级'] = np.where(df['计息日'] == df_fee.at[24,'费率'], -df_fee.at[20,'备注'],df['本期次级期间收益'] + df['本期次级本金分配'] + df['超额收益'])
# df[['计息日','次级']]
# 计算次级xirr
# 定义xirr函数https://www.twblogs.net/a/5caca20abd9eee2dd0f29604/?lang=zh-cn
def xirr(cashflows):
    years = [(ta[0] - cashflows[0][0]).days / 365. for ta in cashflows]
    residual = 1.0
    step = 0.05
    guess = 0.05
    epsilon = 0.0001
    limit = 10000
    while abs(residual) > epsilon and limit > 0:
        limit -= 1
        residual = 0.0
        for i, trans in enumerate(cashflows):
            residual += trans[1] / pow(guess, years[i])
        if abs(residual) > epsilon:
            if residual > 0:
                guess += step
            else:
                guess -= step
                step /= 2.0
    return guess - 1
# 测试
# data = [(datetime.date(2006, 1, 24), -39967), (datetime.date(2008, 2, 6), -19866), (datetime.date(2010, 10, 18), 245706), (datetime.date(2013, 9, 14), 52142)]
# xirr(data)
# 次级时间、现金流
# 将次级的日期以及次级的现金流先转换组合成一个字典
data = df[['计息日','次级']].set_index('计息日').to_dict()['次级']
# 将字典转换为符合data格式要求的列表
data = list(data.items())
# 计算次级xirr
# xirr(data)
# to do:格式化,定义一个字典，每列的数字格式
#用 hide_columns() 方法可以选择隐藏一列或者多列，代码如下：
# df.style.hide_index().hide_columns(['计算日','基金经理','上任日期','2021'])
#在设置数据格式之前，需要注意下，所在列的数值的数据类型应该为数字格式，
#如果包含字符串、时间或者其他非数字格式，则会报错。
#可以用 DataFrame.dtypes 属性来查看数据格式。这个仅可以设置pandas格式，输出excel的时候需要另外设置
format_dict = {'本金':'￥{0:.2f}',
               '利息':'￥{0:.2f}',
               '计息日':lambda x: "{}".format(x.strftime('%Y%m%d')),
               '增值税':'￥{0:.2f}',
               '附加税':'￥{0:.2f}',
               '托管费':'￥{0:.2f}',
               '信托报酬':'￥{0:.2f}',
               '资产服务费':'￥{0:.2f}',
               '一次性费用':'￥{0:.2f}',
               '前端资产服务费':'￥{0:.2f}',
               '税费汇总':'￥{0:.2f}',
               '期初优先A本金':'￥{0:.2f}',
               '本期优先A利息':'￥{0:.2f}',
               '本期优先A本金分配':'￥{0:.2f}',
               '期末优先A本金':'￥{0:.2f}',
               '期初优先B本金':'￥{0:.2f}',
               '本期优先B利息':'￥{0:.2f}',
               '本期优先B本金分配':'￥{0:.2f}',
               '期末优先B本金':'￥{0:.2f}',
               '期初次级本金':'￥{0:.2f}',
               '本期次级期间收益':'￥{0:.2f}',
               '本期次级本金分配':'￥{0:.2f}',
               '期末次级本金':'￥{0:.2f}',
               '超额收益':'￥{0:.2f}'               
                }
df_save = df[['本金','利息','计息日','增值税','附加税','托管费','信托报酬','资产服务费',
              '一次性费用','前端资产服务费','税费汇总','期初优先A本金','本期优先A利息',
              '本期优先A本金分配','期末优先A本金','期初优先B本金','本期优先B利息','本期优先B本金分配',
              '期末优先B本金','期初次级本金','本期次级期间收益','本期次级本金分配','期末次级本金',
              '超额收益']].dropna().style.format(format_dict,na_rep='-')#空值用-代替
# to do：输出保存，选择特定列进行输出
writer = pd.ExcelWriter("工行测算.xlsx",
                        datetime_format='yyyyMMdd',
#                         datetime_format='mmm d yyyy hh:mm:ss',
                        date_format='mmmm dd yyyy')
df_save.to_excel(writer, sheet_name='Sheet1', index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

format1 = workbook.add_format({'num_format': '#,##0.00'})
# format2 = workbook.add_format({'num_format': '0%'})
worksheet.set_column('A:B', 11, format1)
worksheet.set_column('D:X', 11, format1)


writer.save()
