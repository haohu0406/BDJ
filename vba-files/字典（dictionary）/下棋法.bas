'分类汇总
Sub 分类汇总()
    Dim qp(1 To 10000, 1 To 3)
    Dim row
    Dim arr, x, k
    Dim d As New dictionary
    arr = range("a2:c" & range("a65535").End(xlup).row)
    For x = 1 To UBound(arr)
        If d.exists(arr(x, 1)) Then
            row = d(arr(x, 1))
            qp(row, 2) = qp(row, 2) + arr(x, 2)
            qp(row, 3) = qp(row, 3) + arr(x, 3)
        Else
            k = k + 1
            d(arr(x, 1) = k
            qp(k, 1) = arr(x, 1)
            qp(k, 2) = arr(x, 2)
            qp(k, 3) = arr(x, 3)
        End If
    Next x
    range("f2").resize(k, 3) = qp
End Sub

Sub 多条件汇总()
    Dim qp(1 To 10000, 1 To 3)
    Dim row
    Dim arr, x, k, sr
    Dim d As New dictionary
    arr = range("a2:c" & range("a65535").End(xlup).row)
    For x = 1 To UBound(arr)
        sr = arr(x, 1) & "-" & arr(x, 2)
        If d.exists(sr) Then
            row = d(sr)
            qp(row, 3) = arr(x, 3) + qp(row, 3)
            qp(row, 4) = arr(x, 4) + qp(row, 4)
        Else
            k = k + 1
            d(sr) = k
            qp(k, 1) = arr(x, 1)
            qp(k, 2) = arr(x, 2)
            qp(k, 3) = arr(x, 3)
            qp(k, 4) = arr(x, 4)
        End If
    Next x
    range("g2").resize(k, 4) = qp
End Sub