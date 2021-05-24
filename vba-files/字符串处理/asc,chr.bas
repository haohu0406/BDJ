asc返回一个integer, 达标字符串中首字母的字符代码 ，ansi
chr返回一个string ，其中包含有与制定的字符代码相关的支付
Sub z4()
    debug.Print Asc("z") '列标转换成数字
    debug.Print Chr(90) '数字转换成列标
End Sub


Sub z5()
    debug.Print "a" & Space(10) & "b" '10个空格
    debug.Print "c" & String(10, "*") & "d" '10个*
End Sub