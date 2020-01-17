from pprint import pprint

import pandas as pd
'''
谈到这个模块,我羞愧了,之前我是很想使用模块的
因为之前听到大数据和数据分析的时候,我就知道这个模块了,但是一直没用过
今天我终于要拿起我的刀了,嘿嘿。对了,这还可以操作csv文件哦(有幸操作过一回)
'''

# 读取Excel文件的第一种方法,默认读取第一个sheet:
# df= pd.read_excel('../切换数据.xlsx')
# 默认读取前5行数据,传入参数:读取10行,得到一个二维矩阵
# data=df.head(10)

# 方法2:通过sheet名的方式读取,注意大小写
df=pd.read_excel('book.xlsx',sheet_name='Sheet',header=[0,1])
# data=df.head()

# 方法3:通过sheet索引来访问
# 这就牛逼了,sheet_name传入多个sheet名,相当于定位多个sheet
# dg=pd.read_excel('book.xlsx',sheet_name=['Sheet'])
# 传入多个sheet索引
# dg=pd.read_excel('../切换数据.xlsx',sheet_name=[1,2])

# 获取所有的数据
# data=df.values
# 读取第一行,这里的ix已经弃用了,可以使用iloc,loc
data=df.iloc[0].values
data2=df.loc['彤彤和库斗':'雨中跘']
pprint(data2)