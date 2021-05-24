在内存中读取数组

Dim arr, arr1()
Dim x, k, m As Integer
arr = range("a1:a10")
m = application.countif(range("a1:a10"), ">10")
ReDim arr1(1 To m)
For x = 1 To 10
    If arr(x, 1 ） > 10 Then
        k = k + 1
        arr1(k) = arr(x, 1)
    End If
Next x
End Sub

2.
Dim arr, arr1(1 To 5, 1 To 1)
Dim x As Integer
arr = range("b2:c6")
For x = 1 To 5
    arr1(x, 1) = arr(x, 1) * arr(x, 2)
Next x
range("d2").resize(5) = arr1 '等同于range("d2:d6")=arr1,等同于range("d2").resize(5,1)
End Sub

3.application.transpose(arr1) '转置数组
一维数组通过转置成为多行一列的二位数组
通过selection （一列 ），range （一列 ）赋值的数组是二维数组
Sub t()
    Dim arr(1 To 3)
    arr(1) = 1
    arr(2) = 2
    arr(3) = 3
    brr = Application.Transpose(arr)
    i = UBound(Application.Transpose(arr))
End Sub
