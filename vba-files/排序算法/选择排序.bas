Sub 选择排序()
    Dim i, j, k, arr, temp
    arr = [a1].currentregion
    For i = 1 To UBound(arr) - 1
        temp = arr(i, 1)
        k = i
        For j = i + 1 To UBound(arr)
            If temp > arr(j, 1) Then
                temp = arr(j, 1)
                k = j
            End If
        Next j
        arr(k, 1) = arr(i, 1)
        arr(i, 1) = temp
    Next i
    range("f1"), resize(UBound(arr)) = arr
End Sub