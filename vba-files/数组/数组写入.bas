'为社么可以提高运行速度
'调用数据的两种方式
'1调用内存中数据，2调用单元格的数据

'调用内存中数据比较快，数组存放在内存中，所以数组能提高运行速度

'指针型数据不连续，数组数据连续，故删除添加操作指针型比较快，访问数据数组较快


'数组可以存放数字，文本，对象

1).常量数组
Array （1 ，2 ）
Array （Array(1, 2, 4), Array("a", "b", "c"))
2).静态数组
arr() 有5个位置
arr(1 To 10) 有10个位置
arr （1 To 10 ，1 To 2 ）10 行2列的二维数组
3).动态数组
arr() 不知道多少行列


1.一维数组写入
Dim x As Integer
Dim arr(1 To 10)
For x = 1 To 7
    arr(x) = x * 10
Next x
Stop
End Sub

二维数组写入
Dim x, y As Integer
Dim arr(1 To 5, 1 To 4))
For x = 1 To 5
    For y = 1 To 4
        arr(x, y) = cells(x, y)
    Next y
Next x
End Sub

2.动态数组
Dim arr() '需要两步来生成（redim）
Dim row
row = sheets("sheet2").range("a65535").End(xlup).row - 1
ReDim arr(1 To row)
For x = 1 To row
    arr(x) = cells(x, 1)
Next x
End Sub

3.批量写入
Dim arr()
arr = Array(1, 2, 3, "a")
End Sub

Dim arr
arr = range("a1:d5")
End Sub

数组中字符串连接用 “ + ”
arr(i) = arr(i) & "+" & "论坛"
'返回的是arr本身的字符合并论坛两个字

ReDim Preserve 语句只能改变数组最末维的大小