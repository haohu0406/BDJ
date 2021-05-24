Sub qff()
    Dim i, j, k, d, n, m, y As Integer
    Dim hm(1 To 19), arr, brr(1 To 19), qm(1 To 19)
    arr = Range("h17:h35")
    With Sheet1
        For i = 17 To 35 '给17-35行单元格的数据加千分符
            y = i - 16
            If  .Cells(i, 8) Like "*.*" Then
                hm(y) = Split(.Cells(i, 8), ".")(1)
                qm(y) = Split(.Cells(i, 8), ".")(0)
            Else
                hm(y) = "0"
                qm(y) =  .Cells(i, 8)
            End If
            k = (Len(qm(y)) - 1) \ 3
            d = Len(qm(y)) Mod 3
            If k > 0 Then
                If d > 0 Then
                    n = 1
                    brr(y) = brr(y) & Mid(qm(y), 1, d) & ","
                Else
                    n = 0
                End If
                
                For j = 1 To k - n
                    m = d + (j - 1) * 3 + 1
                    brr(y) = brr(y) & Mid(qm(y), m, 3) & ","
                Next
                
                brr(y) = brr(y) & Mid(qm(y), d + (k - n) * 3 + 1, 3)
            End If
            If  .Cells(i, 8) Like "*.*" Then
                brr(y) = brr(y) & "." & hm(y)
            Else
                If  .Cells(i, 8) = 0 Then
                    brr(y) = brr(y) & hm(y)
                End If
            End If
        Next
         .Range("i17").Resize(UBound(brr)) = Application.Transpose(brr)
    End With
End Sub