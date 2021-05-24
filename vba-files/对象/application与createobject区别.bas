方法一 ：
Dim Ex As Excel.Application, Exw As Excel.workbook
Set Ex = New Excel.Application
Set Exw = ex.workbooks.open(filename)
……

方法二 、
Dim Ex As Object, Exw As Object
Set Ex = Create(" Excel.Application")
Set Exw = ex.workbooks.open(filename)

方法一 ：前期绑定, 好处是在对象后输入句点可以给出快速提示 ，因为需要先引用对象
，所以容易出现版本兼容问题

方法二 ：后期绑定 ，没有提示 ，根据运行代码机器上对象的版本创建对象 ，兼容性好

提示 ：有时二者有较大区别 ，可论坛搜索字典对象 ，建议编写代码时使用前期绑定
，发布时使用后期绑定
