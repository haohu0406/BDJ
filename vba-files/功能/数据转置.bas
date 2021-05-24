Sub zh()
    Dim arr, brr, crr
    arr = sheet1.[a1].currentregion
    ReDim crr(1 To UBound(arr) * UBound(arr, 2), 1 To 3)
    For i = 2 To UBound(arr)
        For j = 2 To UBound(arr, 2)
            k = k + 1
            crr(k, 1) = arr(i, 1)
            crr(k, 3) = arr(1, j)
            crr(k, 2) = arr(i, j)
        Next
    Next
    sheet2.range("a1") = "简称"
    sheet2.range("b1") = "销售"
    sheet2.range("c1") = "类型"    
    sheet2.range("a2:c2").resize(k) = crr
End Sub
'数据转置