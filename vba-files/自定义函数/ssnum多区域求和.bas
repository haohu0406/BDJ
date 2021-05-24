Function ssnum(ParamArray n())
    Dim num, k
    k = 0
    For Each num In n
        k = k + num
    Next num
    ssnum = k
End Function
'sum的高级形式