Sub test_a3()
    Dim arr, i, j, d1, d2, t1, t2, a, n    
    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = CreateObject("Scripting.Dictionary")
    arr = range(sheet1.[a1], sheet1.[d65535].End(3))
    For i = 2 To UBound(arr)
        t1 = arr(i, 2): t2 = t1 & arr(i, 4)
        d1(t1) = "": d2(t2) = d2(t2) & "、" & arr(i, 4)
    Next
    arr = sheet2.[a1:h1].resize(d1.count + 1)
    For Each a In d1.keys
        n = n + 1:arr(n, 1) = a
        For j = 2 To 8:arr(n, j) = Mid(d2(a & arr(1, j)), 2): Next
    Next
    sheet2.[a1].resize(n, 8) = arr
End Sub