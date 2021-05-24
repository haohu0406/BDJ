对比字符串 语法:
字符串1 like 字符串2

Sub t1()
    debug.Print "abc" like "Abc"
End Sub
'返回false

'通配符？代表1个字符串
Sub t2()
    debug.Print "BA" like "?A"
End Sub

'通配符*，代表任意个字符串
Sub t3()
    debug.Print "excel精英培训" like "*cel*"
End Sub
'返回true

Sub t5()
    debug.Print "qab" like "q?b"
    debug.Print "q?b" like "q?b" '返回值为true，但实际无法判断
    debug.Print "q?b" like "q[?]b" '返回值为true
End Sub

'通配符#，代表一个数字
Sub t6()
    debug.Print 559 like "###"
End Sub
'返回true

Sub t6()
    debug.Print "q" like[A - Za - z] 'true
    debug.Print "H" like[A - GM - Z] 'false
    debug.Print 7 like[2 - 6] 'false
End Sub

Sub t8()
    debug.Print "a" like "[abcdefg]" 'true
End Sub

'通配符! 相反的
Sub t0()
    debug.Print "a" like "[!c-z]" '不在这个区间
End Sub

'匹配的模式      匹配的字符床
'?                      任何一个字符
' *                     零个或者多个字符
'#                     任何一个数字
'[字符串列表]   任何一个在字符串列表中的字符
'[!字符串列表]  任何一个不在字符串列表中的字符
'[a - z]             字母a到字母z之间的任何一个字母
'[A - Z]            字母A到字母Z之间的任何一个字母