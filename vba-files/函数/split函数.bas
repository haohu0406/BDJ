Split("test.xls", ".")(0) '返回test，即返回前半部分
Split("test.xls", ".")(1) '返回xls，即返回后半部分
'分离为组，存在数组里

'分离得到""中间的数值
b1 = "bus", "school", "student"
Sub t()
    Dim i
    i = Split([b1], """")(0) '返回空值
    i = Split([b1], """")(1) '返回bus
    i = Split([b1], """")(2) '返回，
    i = Split([b1], """")(3) '返回school
    MsgBox i
End Sub