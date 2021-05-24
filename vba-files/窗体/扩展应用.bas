Sub test5()
    Dim rg As range
    Set rg = application.InputBox("请选择单元格区域", "选区提示", , , , , , , 8)
    '8表示range对象
    MsgBox rg.parent.Name & "|" & rg.address
End Sub


Sub test6()
    Dim rg
    rg = application.InputBox("请选择单元格区域", "选取提示", , , , , , 8)
    '上面取值是一个区域，这里取值是一个数组
    MsgBox rg(2, 1)
End Sub

Sub test7()
    Dim r
    r = application.InputBox("请输入公式", "输入提示", , , , , , 0) '0公式
    MsgBox r
End Sub


Sub test8()
    Dim r
    r = application.InputBox("请输入公式", "输入提示", , , , , , 1) '数字
    MsgBox r
End Sub


Sub test9()
    Dim r
    r = application.InputBox("请输入公式", "输入提示", , , , , , 2) '可以输入字符
    MsgBox TypeName(r)
End Sub

Sub test10()
    Dim r
    r = application.InputBox("请输入公式", "输入提示", , , , , , 64) '数组
    MsgBox r(2, 1) '当输入一位数组时会提示出错，因为此处是取二维数组的值
End Sub

Sub test1()
    Dim sr
    sr = InputBox("输入测试", "测试", 100)
    MsgBox sr
    sr = application.InputBox("输入测试", "测试", 100)
    MsgBox sr
End Sub

'不输入内容分别返回什么
Sub test2()
    Dim sr
    sr = InputBox("输入测试", "测试", 100) '返回空值
    MsgBox sr
    sr = application.InputBox("输入测试", "测试", 100) '返回空值
    MsgBox sr
End Sub

'点击取消，返回什么
Sub test2()
    Dim sr
    sr = InputBox("输入测试", "测试", 100) '返回空值
    MsgBox sr
    sr = application.InputBox("输入测试", "测试", 100) '返回false
    MsgBox sr
End Sub