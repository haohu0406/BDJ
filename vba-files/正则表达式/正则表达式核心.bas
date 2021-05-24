'特殊符号
' \号
'1.放在不便书写的字符前面 ，如换行符 （ \ r), 回车符( \ t),  \ 自身( \ \)
'2.放在有特殊意义字符的前面 ，表示它自身 ，"\$", "\^", "\."
'3.放在可以匹配多个字符的前面
' \ d 0 - 9 的数字
' \ w 任意一个字母或数字或下划线 ，也就是A - Z, a - z, 0 - 9,  _ 中任意一个
' \ s 包括空格 、制表符 、换页符等空白字符的其中任意一个

'以上改为大写时 ，为相反的意思 ，如 \ D表示非数字类型

Sub t1()
    Dim regx As New regexp
    Dim sr
    sr = "ae45b646c"
    With regx
         .global = True
         .pattern = "\D" '排除非数字
        debug.Print  .Replace(sr, "")
    End With
End Sub


'.（点）
'可以匹配除换行符以外的所有字符
'.也是一个占位符


'+号
'表示一个字符可以有任意多个重复的
Sub t11()
    Dim regx As New regexp
    Dim sr
    sr = "a234ca7a"
    With regx
         .global = True
         .pattern = "A\d+"
        debug.Print  .Replace(sr, "")
    End With    
End Sub

'{}号
'可以设置重复次数
'{n}重复n次
Sub t16()
    Dim regx As New regexp
    Dim sr
    sr = "a2234123"
    With regx
         .global = True
         .pattern = "\d{5}" '连续5个数字
        debug.Print  .Replace(sr, "")
    End With
End Sub

'{m,n}最小重复M次，最多重复n次
Sub t22()
    Dim regx As New regexp
    Dim sr
    sr = "a2348ca7a67"
    With regx
         .global = True
         .pattern = "\d{2,3}" '连续两个数字或者连续三个数字
        debug.Print  .Replace(sr, "")
    End With
End Sub

'{m，}最少重复m次，相当于+
Sub t23()
    Dim regx As New regexp
    Dim sr
    sr = "a2348ca7a67"
    With regx
         .global = True
         .pattern = "\d{2,}" ' 连续两个数组或连续三个数字
        debug.Print  .Replace(sr, "")
    End With
End Sub

'* 可以出现0等任意次数
'相当于{0, }, 比如: "\^*b" 可以匹配 "b", "^^^b"

'?  匹配表达式0次或者1次，相当于{0,1},比如："a[cd]?"可以匹配"a","ac","ad"
Sub t24()
    Dim regx As New regexp
    Dim sr
    sr = "a23.48ca7a6..7"
    With regx
         .global = True
         .pattern = "\d+\.?\d+" '最多连续出现1次
        debug.Print  .Replace(sr, "")
    End With    
End Sub
'利用+?的格式可以分段匹配
Sub t36()
    Dim regx As New regexp
    Dim sr, mat, m
    sr = "<td><p>aa</p></td> <td><p>bb</p></td>"
    With regx
         .global = True
         .pattern = "<td>.*?</td>"
        Set mat =  .excute(sr)
        For Each m In mat
            debug.Print m
        Next
    End With
End Sub


'我想查看数组中哪一个字符串dog的前面是big
'说一下 ? =
'这个元字符组合表示当所要搜索的字符串匹配了模式开头部分时就接着匹配后面这一部分
'按你的要求 ：
'你开头部分是要big ，后面要是dog
'那就把big放在开头 ，匹配模式放在后面 ，中间用 ? = 连接就可以了 。
'就变成big(? = dog)
'这就是所谓的向前查找 。

'再举个例子 ："windows (?=95|98|NT)" 这个模式能匹配 "windows 95" 中window ，而不能匹配 "windows vista“中的windows

'而你后面所提的要求 ，属于向后查找 ，据我所知 ，VBS不支持向后查找 。