' 子函数3：扩展数组第一维
Public Function redim_arrs_row(arr, x)
    Dim brr, crr
    brr = Application.Transpose(arr)
    ReDim Preserve brr(1 To UBound(brr), 1 To x)
    crr = Application.Transpose(brr)
    redim_arrs_row = crr
    Erase brr, crr
End Function