Sub test_a1()
    Dim arr, brr, crr, a, xD, i&, j%, n&(5), t$, u&
    Set xD = CreateObject("Scripting.Dictionary")
    arr = range(sheet2.[d1], sheet2.[a65535].End(3))
    For i = 2 To UBound(arr)
        xD(arr(i, 1) & "|" & arr(i, 2) & "|" & arr(i, 3)) = i
    Next i
    brr = range(sheet1.[c1], sheet1.[a65535].End(3))
    crr = brr
    For i = 2 To UBound(brr)
        t = brr(i, 1) & "|" & brr(i, 2) & "|" & brr(i, 3)
        If xd.exists(t) Then
            n(3) = n(3) + 1:xD.remove t
            For j = 1 To 3:crr(n(3) + 1, j) = brr(i, j): Next
        Else
            n(4) = n(4) + 1
            For j = 1 To 3:brr(n(4) + 1, j) = brr(i, j): Next            
        End If
    Next i
    For Each a In xD.items
        n(5) = n(5) + 1
        For j = 1 To 4 ：arr(n(5) + 1, j) = arr(a, j): Next
    Next
    xD(3) = crr: xD(4) = brr: xD(5) = arr
    For j = 3 To 5
        sheets(j).usedrange.clearcontents
        sheets(j).[a1].resize(n(j) + 1, IIf(j = 5, 4, 3)) = xD(j)
    Next j
End Sub



Sub tt()
    Dim arr, brr, crr, i&, j&, t1$, d1, a, n&(5)
    Set d1 = CreateObject("Scripting.Dictionary")
    brr = Sheet2.UsedRange
    For i = 2 To UBound(brr)
        t1 = brr(i, 1) & "|" & brr(i, 2)
        d1(t1) = i
    Next
    arr = Sheet1.UsedRange
    crr = arr
    For i = 2 To UBound(arr)
        t1 = arr(i, 1) & "|" & arr(i, 2)
        If d1.Exists(t1) Then
            n(3) = n(3) + 1: For j = 1 To 3:crr(n(3) + 1, j) = arr(i, j): Next
        d1.Remove t1
    Else
        n(4) = n(4) + 1: For j = 1 To 3:arr(n(4) + 1, j) = arr(i, j): Next
End If
Next
For Each a In d1.Items
    n(5) = n(5) + 1: For j = 1 To 4:brr(n(5) + 1, j) = brr(a, j): Next
Next
d1(3) = crr: d1(4) = arr: d1(5) = brr
For j = 3 To 5
    Sheets(j).UsedRange.ClearContents
    Sheets(j).[a1].Resize(n(j) + 1, IIf(j = 5, 4, 3)) = d1(j)
Next
End Sub