Sub hz()
    Set xD = CreateObject("Scripting.Dictionary")
    Dim i, j, arr, sr
    arr = [a1].currentregion
    
    For i = 2 To UBound(arr)
        sr = arr(i, 1) & arr(i, 2) & arr(i, 3)
        d(sr) = d(sr) + arr(i, 4)
    Next
    
    
End Sub


Sub TEST_A1()
    Dim Arr, xD, i&, j%, T1$, T2$, N&, R&
    Set xD = CreateObject("Scripting.Dictionary")
    Arr = range([d1], [a65535].End(3)): N = 1
    For i = 2 To UBound(Arr)
        T1 = arr(i, 1) & "|" arr(i, 2) & "|" & arr(i, 3)
        T2 = T1 & "|" & Arr(i, 4)
        xD(T2) = xD(T2) + 1:If xD(T2) > 1 Then GoTo i01
        R = xD(T1)
        If R > 0 Then Arr(R, 4) = Arr(R, 4) + Arr(i, 4): GoTo i01
        N = N + 1:R = N: xD(T1) = N
        For j = 1 To 4:Arr(R, j) = Arr(i, j): Next
i01:    Next i
    With [j1].resize(N, 4)
         .value = Arr
        For j = 3 To 1 Step -1
             .sort key1:=.item(j), order1: xlascending, header:=xlyes
        Next j
    End With
End Sub


Sub TEST_A1()
    Dim Arr, xD, i&, j%, T1$, T2$, N&, R&
    Set xD = CreateObject("Scripting.Dictionary")
    Arr = range([d1], [a65535].End(3))
    For i = 2 To UBound(Arr)
        T1 = Arr(i, 1) & "|" & Arr(i, 2) & "|" & Arr(i, 3)
        T2 = T1 & Arr(i, 4)
        xD(T2) = xD(T2) + 1:If xD(T2) > 1 Then GoTo i01
        R = xD(T1)
        If R > 0 Then Arr(R, 4) = Arr(R, 4) + Arr(i, 4): GoTo i01
        N = N + 1:R = N: xD(T1) = N
        For j = 1 To 3 ：Arr(R, j) = Arr(i, j): Next j
        i01 ：Next i
    With [j1].resize(N, 4)
         .value = Arr
        For j = 3 To 1 Step -1
             .sort key1:=.item(j), order1: xlascendding, header:=xlyes
        Next j
    End With
End Sub

Sub TEST_A1()
    Dim Arr, T1$, T2$, R&, N&, i&, j&, xD
    Set xD = CreateObject("Scripting.Dictionary")
    Arr = range([d1], [a65535].End(3))
    For i = 2 To UBound(Arr)
        T1 = Arr(i, 1) & "|" & Arr(i, 2) & "|" & Arr(i, 3) & "|"
        T2 = T1 & Arr(i, 4)
        xD(T2) = xD(T2) + 1:If xD(T2) > 1 Then GoTo i01
        R = xD(T1)
        If R > 0 Then Arr(R, 4) = Arr(R, 4) + Arr(i, 4): GoTo i01
        N = N + 1:R = N: xD(T1) = N
        For j = 1 To 3:Arr(R, j) = Arr(i, j): Next j
i01:    Next i
    With [j1].resize(N, 4)
         .value = Arr
        For j = 3 To 1 Step -1
             .sort key1:=.item(j), order: xlascending, header:=xlyes
        Next j
    End With
End Sub