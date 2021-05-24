Sub 冒泡排序()
    Dim i, j As Integer
    Dim arr, temp
    arr = [a1].currentregion
    For i = 1 To UBound(arr)
        For j = UBound(arr) To i + 1
            If arr(j, 1) < arr(j - 1, 1) Then
                temp = arr(j, 1)
                arr(j, 1) = arr(j - 1)
                arr(j - 1, 1) = temp
            End If
        Next j
    Next i
End Sub
'总结：从下往上，比较相邻两个数，较小的上浮
'		第一趟找出最小的值	
'外层循环次数：每循环一次只能确定一个最小值，故要找N次最小值。
'内层循环次数；因每次排序确定一个最小值，故没下一轮循环减少一个最小值。所以
'内层循环次数是总循环次数减外层循环次数。