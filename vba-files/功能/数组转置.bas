Function Trans(ByRef mrr)
    Dim i&, j&, nrr
    On Error Resume Next
    ReDim nrr(LBound(mrr, 2) To UBound(mrr, 2), LBound(mrr, 1) To UBound(mrr, 1))
    If Err.Number <> 0 Then Exit Function
    For i = LBound(mrr, 1) To UBound(mrr, 1)
        For j = LBound(mrr, 2) To UBound(mrr, 2)
            nrr(j, i) = mrr(i, j)
        Next
    Next
    mrr = nrr
End Function