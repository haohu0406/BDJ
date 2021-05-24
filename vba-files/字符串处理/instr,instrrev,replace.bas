字符串的查找与替换

Sub c1()
    Dim sr
    sr = "excel精英培训"
    debug.Print InStr(sr, "精英") '和find相反的用法
End Sub
'返回第一个字符的位置，6

Sub c2()
    Dim sr
    sr = "excel精英培训论坛"
    debug.Print InStrRev(sr, "培")
End Sub
'从后往前查找

Sub c3()
    Dim sr
    sr = "excel精英培训网"
    sr = Replace(sr, "培训网", "论坛")
    debug.Print sr
End Sub

Sub c4()
    Dim sr
    sr = "excel精英培训网"
    Mid(sr, 8, 2) = "论坛" '当mid(sr,8,3)="论坛",替换的也是两个字符，而不是3个字符
    debug.Print sr
End Sub
