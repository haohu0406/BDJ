Sub fontcolor()
    Dim arr, i&, d1, t&, k$
    Dim rng As range: Dim rng1 As range
    Set d1 = CreateObject("Scripting.Dictionary")
    arr = range([m1], [m65535].End(3))
    For i = 2 To UBound(arr)
        k = arr(i, 1): If k = "" Then GoTo 101
        t = d1(k): If t = 0 Then d1(k) = i: GoTo 101
        Set rng = range("m" & i)
        If t > 0 Then Set rng = union(rng, range("m" & t)): d1(k) = -1
        If rng1 Is Nothing Then Set rng1 = rng Else Set rng1 = union(rng1, rng)
101        :
    Next i
    If Not rng1 Is Nothing Then rng1.font.color = vbRed
End Sub