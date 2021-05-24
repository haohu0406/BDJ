'数组的大小

Sub d3()
    Dim arr
    Dim k, m
    arr = range("a2:d5")
    For x = 1 To UBound(arr, 1)
        
    Next
End Sub

'动态数组的动态扩容
ReDim Preserve arr()

Sub d7()
    Dim arr, arr1()
    
End Sub

'一维数组的动态扩充
Private Sub combobox1_gotfocus()
    Dim arr(), , arr1, k
    arr1 = range("a1:a10")
    For x = 1 To UBound(arr1)
        If arr1(x, 1) > 10 Then
            k = k + 1
            ReDim Preserve arr(1 To k)
            arr(k) = arr1(x, 1)
        End If
    Next
    combobox1.list = arr
End Sub


'二维数组的动态扩充，只能扩充列
Sub d7()
    Dim arr, arr1()
    Dim x, k
    For x = 1 To UBound(arr)
        If arr(x, 1) = "b" Then
            k = k + 1
            ReDim Preserve arr1(1 To 4, 1 To k)
            arr1(1, k) = arr(x, 1)
            arr1(2, k) = arr(x, 2)
            arr1(3, k) = arr(x, 3)
            arr1(4, k) = arr(x, 4)
        End If
    Next
    range("a8").resize(k, 4) = application.transpose(arr1)
End Sub

'取数组一部份赋值给单元格区域
Sub d8()
    Dim arr, arr1(1 To 100000, 1 To 4)    
    arr = range("a1:d6")
    Dim x, k
    For x = 1 To UBound(arr)
        If arr(x, 1) = "b" Then
            k = k + 1
            arr1(k, 1) = arr(x, 1)
            arr1(k, 2) = arr(x, 2)
            arr1(k, 3) = arr(x, 3)
            arr1(k, 4) = arr(x, 4)
        End If
    Next
    range("a15").resize(k, 4) = arr1
End Sub

Erase arr '清空数组
