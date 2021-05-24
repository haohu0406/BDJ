d.keys(1) '此方放只有 dim d as new dictionary才可以用，只有引用法可以用
d("李四") '此方法对应上述，只适用于创建法
d.keys是集合相当于一维数组上标为0
d.count

'区分大小写的字典
Sub t5()
    Dim d As New dictionary
    Dim x
    d.compare = binarycompare '区分大小写
    For x = 1 To 5
        d(cells(x, 1).value) = ""
    Next x
    Stop
End Sub


Sub 提取不重复值()
    Dim d As New dictionary
    Dim arr, x
    arr = range("a2:a12")
    For x = 1 To UBound(arr)
        d(arr(x, 1)) = ""
    Next x
    range("c2").resize(d.count) = application.transpose(d.keys)
End Sub

Sub 汇总()
    Dim d As New dictionary
    Dim arr, x
    arr = range("a2:a12")
    For x = 1 To UBound(arr)
        d(arr(x, 1)) = d(arr(x, 1)) + arr(x, 2)
    Next x
    range("c2").resize(d.count) = application.transpose(d.keys)
    range("d2").resize(d.count) = application.transpose(d.items)
End Sub

Sub 查找()
    Dim d As New dictionary
    Dim arr, x
    For x = 3 To 5
        arr = sheets(x).range("a2").resize(sheets(x).range("a65535").End(xlup).row - 1, 2)
        For y = 1 To UBound(arr)
            d(arr(y, 1)) = arr(y, 2) '此处字典添加数据用此方法
            d(arr(y, 2)) = arr(y, 1) '正反查找
        Next y
    Next x
    MsgBox d("c1")
    MsgBox d("无情")
End Sub

'判断字典关键字是否存在
d.exists()
'也可以用以下形式代替
r = d()
If r <> 0 Then
    
End If

arr = d.keys '获得的数组是一个0开始的一维数组