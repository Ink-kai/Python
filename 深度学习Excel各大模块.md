[xlrd]:https://github.com/Ink-kai/Python/blob/master/excel/depth_learn_xlrd.py
[xlwt]:https://github.com/Ink-kai/Python/blob/master/excel/depth_learn_xlwt.py
[xlutils]:https://github.com/Ink-kai/Python/blob/master/excel/depth_learn_xlutils.py
[openpyxl]:https://github.com/Ink-kai/Python/blob/master/excel/depth_learn_openpyxl.py
[xlsxwriter]:https://github.com/Ink-kai/Python/blob/master/excel/depth_learn_xlutils.py
#### 今天,详解一下处理Excel各大模块,算是对知识的一个总结吧
##### 来列举一下各个模块以及之间的性能

- xlwings:可结合VBA实现对Excel编程,强大的数据输入分析能力,同时拥有丰富的的接口,结合`pandas`/`numpy`/`matplotlib`轻松应对 Excel 数据处理工作
- openpyxl:简单易用,功能广泛,单元格格式/图片/表格/公式/筛选/批注/文件保护等等功能应有尽有,图表功能是其一大亮点,缺点是对 VBA 支持的不够好(其实就很爽了)
- pandas:使用需要结合其它库,数据处理是`pandas`立身之本(原谅我,竟然不熟)
- win23com:是一个win应用的扩展,不仅仅是Excel,可以处理office。需要注意,该库不单独存在,可通过安装`pypiwin32`或者`pywin32`获取(特点倒是吹的牛逼)
- xlsxwriter:看模块名字就知道是处理xlsx纯写入的,不过,它拥有丰富的特性,支持图片/表格/图表/筛选/格式/公式等,功能与`openpyxl`相似,优点是相比`openpyxl`还支持VBA文件导入,迷你图等功能
- dataNitro:作为插件内嵌到excel中,可完全替代VBA,在excel中优雅的使用python(就相当于在Excel中使用Python代码),不扯了,付费的哈
- xlutils:结合`xlrd`/`xlwt`,老牌py包了,但是必须同时安装这三个库,而且仅支持xls文件

|         | 打开文档 | 新建文档 | 修改文档 | 保存文档 |
| :-------: | :-------: | :-------: | :-------: | :-------: |
| `win32com` | true | true | true | true |
| `xlwings` | true | true | true | true|
| `xlsxwriter` | false | true | false | true|
| `pandas` | true | false | true | true|
| `openpyxl` | true | true | true | true|
| `xlutils` | true | true | true | true|

1.xlrd&xlwt&xlutils(读&写&修改excel)
> + 官方文档：http://www.python-excel.org/
> + xlrd官方介绍：https://pypi.python.org/pypi/xlrd/1.0.0
> + xlwt官方介绍：https://pypi.python.org/pypi/xlwt/1.1.2
> + xlutils官方介绍：https://pypi.python.org/pypi/xlutils

[xlrd例子][xlrd]  
[xlwt例子][xlrd]  
[xlutils例子][xlrd]

2.openpyxl(可读可写,本人常用模块,但是看了前面仨,不敢说最爱了)
> * openpyxl文档:https://openpyxl.readthedocs.io/en/stable/

[openpyxl例子][openpyxl]

3.xlsxwriter(主要用来生成excel表格，插入数据、插入图标等表格操作)
> + 英文文档:https://xlsxwriter.readthedocs.io/index.html

[xlsxwriter例子][xlsxwriter]

4.pandas(详细介绍一下这位大哥大级别的"人物")
Pandas是一个强大的分析结构化数据的工具集；它的使用基础是Numpy（提供高性能的矩阵运算）；用于数据挖掘和数据分析，同时也提供数据清洗功能
利器之一：DataFrame
DataFrame是Pandas中的一个表格型的数据结构，包含有一组有序的列，每列可以是不同的值类型(数值、字符串、布尔型等)，DataFrame即有行索引也有列索引，可以被看做是由Series组成的字典。
利器之一：Series
它是一种类似于一维数组的对象，是由一组数据(各种NumPy数据类型)以及一组与之相关的数据标签(即索引)组成。仅由一组数据也可产生简单的Series对象

__后续会补上的...__