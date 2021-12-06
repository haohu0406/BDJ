
Sub 科目明细表按月分析()
    Dim arr, i&, j&, erow&, k&, m&,fx&
    With ActiveWorkbook.ActiveSheet
    i=MsgBox "取借方数选是，取贷方数选否",vbyesno,"请选择" 
    if i=6 then fx=7
    fx=8
    arr = .UsedRange
    erow = UBound(arr) - 2
    Const sr1 = "本月合计"

    k = 2
    ReDim Preserve arr(1 To UBound(arr), 1 To 12)
    For j = 2 To 12
        arr(1, j) = j - 1
    Next j
    For i = 3 To UBound(arr)
        If arr(i, 5) = sr1 And arr(i - 1, 2) <= arr(erow, 2) Then
            If arr(i - 1, 1) <> arr(k, 1) Then
                If k = 2 Then
                    arr(k, 1) = arr(i - 1, 1)
                Else
                    arr(k + 1, 1) = arr(i - 1, 1): k = k + 1
                End If
            End If
            arr(k, arr(i - 1, 2) + 1) = arr(i, fx)
        End If
    Next i
    .UsedRange.Clear
    .Range("a1").Resize(k, 12) = arr
    End With
End Sub


    