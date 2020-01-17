import xlsxwriter
'''
XlsxWriter是一个Python模块，用于以Excel 2007+ XLSX文件格式编写文件。
它可用于将文本，数字和公式写入多个工作表，它支持格式化，图像，图表，页面设置，自动过滤，条件格式等功能。
缺点就是只能创建文件,不能对已有文件进行读、写
如果是创建文件名存在,那么原有文件的内容不会有任何改变,只能是追加
'''

# 创建workbook对象
wb= xlsxwriter.Workbook('xlsxwriter文件.xlsx')
# 添加sheet,不存在获取操作sheet的说法好吧
sheet=wb.add_worksheet()

# 设置表头
headings = ['Number','testA','testB']
data = [
    ['2017-9-1','2017-9-2','2017-9-3','2017-9-4','2017-9-5','2017-9-6'],
    [10,40,50,20,10,50],
    [30,60,70,50,40,30],
]

# 写入表头数据,按行写入
sheet.write_row('A1',headings)
# write_column是按列写入,那么的list就不要长度不一哦
sheet.write_column('A2',data[0])
sheet.write_column('B2',data[1])
sheet.write_column('C2',data[2])
# 第一行第一列单元格
# sheet.write(0,0,'标题')

# 插入折线图,果然还是有点东西
chart_col=wb.add_chart({'type':'line'})
# 图表设置格式,填充内容
chart_col.add_series(
    {
    'name':'=sheet1!$B$1',
    'categories':'=sheet1!$A$2:$A$7',
    'values':'=sheet1!$B$2:$B$7',
    'line': {'color': 'red'},
    }
)
# 设置图表表头及坐标轴
chart_col.set_title({'name':'测试'})
chart_col.set_x_axis({'name':"x轴"})
chart_col.set_y_axis({'name':'y轴'})
chart_col.set_style(1)
# 放置图表位置
sheet.insert_chart("B10",chart_col,{'x_offset':25,'y_offset':10})

# 这里写完数据,不是保存数据,而是关闭文件
wb.close()

'''
xlsxwriter模块虽然只能写,但是写的功能也蛮强大的,插入表格(折线、柱状、面积图等等)、设置单元格格式
弱水三千,只取一瓢。这些模块都牛逼,好好利用吧
'''