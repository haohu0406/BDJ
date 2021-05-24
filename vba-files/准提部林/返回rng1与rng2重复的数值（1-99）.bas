Function check(rng1, rng2) As String
    Dim Ar%(1 To 99), A
    For Each A In rng1
        Ar(A) = 1
    Next
    For Each A In rng2
        Ar(A) = Ar(A) + 2
        If Ar(A) = 3 Then check = Trim(check & " " & A) 'trim去除字符串前后空格
    Next
End Function

Sub tq()
    Dim arr, brr, i&, j&, m&, k&, s&, n&
    arr = sheets("sheet1").usedrange
    ReDim brr(1 To 2000, 1 To 4)
    n = 1
    For i = 2 To UBound(arr)
        For j = 2 To 4
            m = arr(i, j): brr(1, j) = arr(1, j)
            If m = "" Then GoTo 101
            k = (m - 1) \ 150 + 1
            For s = 1 To k
                n = n + 1:brr(n, 1) = arr(i, 1): brr(n, j) = IIf(m > 150, 150, m)
                m = m - 150
            Next s
101            :
        Next j
    Next i
    brr(1, 1) = arr(1, 1)
    [sheet2! a1].resize(n, 4) = brr
End Sub