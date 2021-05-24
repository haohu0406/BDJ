Sub qfw()
    Dim t As Double
    t = Timer
    Dim arr, i, j, sr, sr1, sr2, brr
    arr = [g17:g35]
    ReDim brr(1 To UBound(arr))
    For i = 17 To 35
        sr = Split(arr(i - 16, 1), ".")(0)
        l = Len(sr)
        s = (l - 1) \ 3
        k = l - s * 3
        If k > 0 Then
            If arr(i - 16, 1) Like "*.*" Then
                sr1 = Mid(sr, 1, k) & ","
            Else
                sr1 = Mid(sr, 1, k)
            End If
        End If
        For j = k + 1 To l Step 3
            If j < l - 3 Then
                sr1 = sr1 & Mid(sr, j, 3) & ","
            Else
                sr1 = sr1 & Mid(sr, j, 3)
            End If
            k = 3
        Next
        If arr(i - 16, 1) Like "*.*" Then
            brr(i - 16) = sr1 & "." & Split(arr(i - 16, 1), ".")(1)
        Else
            brr(i - 16) = sr1
        End If
        sr1 = ""
    Next
    [j17:j35] = Application.Transpose(brr)
    Debug.Print Timer - t
End Sub