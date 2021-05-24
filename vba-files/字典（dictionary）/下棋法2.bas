Sub 下棋法数据透视表式汇总()
    Dim d As New dictionary
    Dim qp(1 To 10000, 7)
    Dim arr, sr, k, x, cl, ro
    For x = 1 To UBound(arr)
        cl = (InStr("1月2月3月4月5月6月", arr(x, 2)) + 1) / 2 + 1
        If d.exists(arr(x, 1)) Then
            ro = d(arr(x, 1))
            qp(ro, cl) = qp(ro, cl) + arr(x, 3)
        Else
            k = k + 1
            d(arr(x, 1)) = k
            qp(k, cl) = arr(x, 3)
        End If
    Next x
    range("f2").resize(k, 7) = qp
End Sub
