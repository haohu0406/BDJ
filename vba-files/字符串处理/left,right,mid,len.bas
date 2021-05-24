Sub z1()
    Dim sr
    sr = "excel精英培训网"
    debug.Print Left(sr, 5)
    debug.Print Right(sr, 5)
    debug.Print Mid(sr, 3, 5)
    debug.Print Left(sr, Len(sr) - 1)
End Sub

Sub z2()
    Dim sr, arr
    
    Sub z3()
        Dim sr
        sr = "89.90美元"
        debug.Print Val(sr)
    End Sub
    '只能取得数字开头的字符串的数字
    
    Sub a4()
        Dim sr, arr
        sr = "excel-精英-培训网"
        arr = Split(sr, "-")
        debug.Print Join(arr, "+")
    End Sub
    'excel+精英+培训网
    
    Mid(要处理的字符串 ，开始的字符 【包含本身 】，要提取的字符个数)
    Sub tt1()
        k = Mid(Range("h17"), 1, 0)
        MsgBox k
    End Sub
    '返回的是空