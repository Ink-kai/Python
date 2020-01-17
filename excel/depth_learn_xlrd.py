import xlrd

'''
不说把xlrd模块玩烂,但是使用xlrd玩转Excel还是不成问题的
xlrd是读取xlsx、xls文件的哦,只是读
'''

# 文件
file= '../切换数据.xlsx'
# 打开文件,这里需要注意:xlrd默认编码是ASCII,如果不是UnicodeDecodeError会引发异常,可以自己设定编码,设置参数encoding_override='编码'
data= xlrd.open_workbook(file)

# 方法1：根据sheet下标(也叫索引)获取,0开始
sheet=data.sheet_by_index(0)
# 方法2：知道sheet名字
# sheet=data.sheet_by_name('sheet名')
# 方法3：也是通过索引获取
# sheet=data.sheets()[0]
# 获取sheet名,sheet_names()获取所有的sheet名,[0]取第一个sheet名
# sheet=data.sheet_names()[0]

# 检查某个sheet是否导入完毕,参数可以是sheet名,也可以是索引
data.sheet_loaded(0)

# 最大有效行(空的不会计算)
max_row=sheet.nrows
# 最大有效列
max_col=sheet.ncols

# xlrd获取Excel的单元格(行列)。这里需要注意一下,range(start,end,step),start默认0,end是传入的数字-1,step在这里...我没用过
# 这里为什么从1开始:兄弟,这里是操作Excel,你见过谁用Excel计数从0开始的
for col in range(1,max_col):
    # 行的话,从2开始,1是单元格标题
    for row in range(2,max_row):
        # 获得单元格对象(还可以进行一些其他操作),获取单元格类型不就是咯
        # sheet.cell(行,列).value是获取值 .ctype是单元格类型(等会会详解)
        cell=sheet.cell(row,col)
        # 直接获取值
        cell_value=sheet.cell_value(row,col)

# 获取整行、整列,都是传入参数索引,从0开始
rows=sheet.row_values(3)
cols=sheet.col_values(3)
# 字面意思呗,行的长度,不就是列嘛...不知道有什么用,还是把方法讲一下
row_len=sheet.row_len(3)

# 获取单元格值的方法,循环中的2种,这里再讲几种(0开始):
# col_values(列,开始行,结束行):常用的这三种参数,第1列的第4行到第22行
# 注意:传入参数必须有start_row,结束行不写默认最大行
col_value=sheet.col_values(1,4,22)
# 跟col_values相反:row_value(行,开始列,结束列),注意跟col_value一样
# 第1行第5列到第7列
row_value=sheet.row_values(1,5,7)

'''
xlrd单元格的数据类型
0   empty(英文空好像,英语有点差)
1   string
2   number
3   date
4   boolean(布尔)
5   error(错误)
'''