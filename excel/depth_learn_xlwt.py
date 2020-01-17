import xlwt
'''
xlwt是对xlsx文件写入数据操作的,
其实我对xlwt和xlrd模块用的少,为什么我敢讲,其实把一个模块玩得差不多了,就会发现其实都差不多,主要还是看需求

xlwt操作基本流程:
1.创建文件、sheet
2.写入数据
3.看需要对单元格进行格式设置(就是字体、颜色啥的,太骚了,不会)
4.保存文件
'''

# 创建工作薄,编码还是总结的时候讲吧
wb= xlwt.Workbook()
# 创建一个sheet,可指定名称。cell_overwrite_ok文档是说如果为True,则如果多次写入,则添加的工作表中的单元格不会引发异常()
sheet=wb.add_sheet('sheet2',cell_overwrite_ok=True)

# 写入数据(参数:行、列、值),这是不带入样式,下面讲一下样式设置
sheet.write(0,0,'值')

# 初始化样式
style=xlwt.XFStyle()
# 为样式创建字体(就是初始化字体,特么的,差点把我搞傻了)
font=xlwt.Font()
# 开始设置字体
font.name='宋体'
# blod粗体
font.bold=True
# 下划线
font.underline=True
# 斜体字
font.italic=True
# 设置字体样式
style.font=font
for col in range(1,6):
    for row in range(1,10000):
        if col==0:
            sheet.write(row, col, '第{}列{}行的值'.format(col,row), style)
        elif col==1:
            sheet.write(row, col, '第{}列{}行的值'.format(col,row), style)
        elif col == 2:
            sheet.write(row, col, '第{}列{}行的值'.format(col,row), style)
        elif col == 3:
            sheet.write(row, col, '第{}列{}行的值'.format(col,row), style)
        else:
            sheet.write(row, col, '第{}列{}行的值'.format(col,row), style)

# 保存文件,参数是文件名
wb.save('xlwt文件.xls')

'''
这里只讲一下写入和简单的样式
我看出来了,xlwt专门处理xls的写入,更牛逼的就是设置格式,个人来讲,不是很需要
'''