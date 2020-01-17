import xlrd
from xlutils.copy import copy

'''
xlutils(Excel utilities)是一个提供了许多操作修改Excel文件方法的库。
xlrd库仅用于读取excel文件中的数据，xlwt库则用于将数据写入excel文件。
但是对于已有的excel文件，想要追加或者修改，这两个库则没有办法完成。
xlutils库也仅仅是通过复制一个副本进行操作后保存一个新文件，xlutils库就像是xlrd库和xlwt库之间的一座桥梁。
因此，xlutils库是依赖于xlrd和xlwt两个库的

看到这,你就应该知道xlrd、xlwt和xlutils的关系了吧
xlrd    只读
xlwt    只写
xlutils 修改(可以覆盖原单元格的值)
'''

rb= xlrd.open_workbook('xlwt文件.xls')
# 这行代码至关重要,这里的copy是xlutils.copy模块,它的的作用:将xlrd.book(就是rb)对象拷贝为一个xlwt.Workbook对象,现在是不是将只读文件对象拷贝成了一个只写对象了
wb= copy(rb)
# 获取第一个sheet
wb_sheet=wb.get_sheet(0)
# 现在的wb对象是一个xlwt.Workbook对象了,可以写了
wb_sheet.write(5,8,'xlutils模块的魅力')
wb.save('xlwt文件.xls')

'''
总的来说,xlutils.copy模块就是把xlrd文件对象拷贝成xlwt文件对象
xlutils肯定不止如此,具体请看文档,这里只是把copy模块讲了
'''