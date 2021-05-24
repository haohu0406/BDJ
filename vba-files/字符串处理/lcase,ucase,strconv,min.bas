Sub z1()
    debug.Print LCase("abc") '第一个字符转为小写
End Sub

Sub z2()
    debug.Print UCase("abc") '第一个字符转为大写
End Sub

strconv函数
常数 值 说明
vbuppercase 1 将字符串文字转换成大写
vblowercase 2 将字符串文字转换成小写
vbpropercase 3 将字符串中每个字的开头字母转成大写

Sub 转换()
    debug.Print VBA.StrConv("wHo ARE you?", vbpropercase)
End Sub
'返回Who Are You?