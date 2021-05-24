固定区间返回不重复随机数
Function shuju(maxnum, geshu) 'maxnum区间最大的数，geshu返回多少个不重复的数
    Dim d As New dictionary
    Dim num
    application.volatitle '异适应函数自动更新
    Do
        num = int(Rnd() * maxnum + 1)
        d(num) = ""
    Loop Until d.count = geshu
    shuju = application.transpose(d.keys)
End Function
'此函数为数组，需按数组函数按键


参数值默认和参数缺省 ，一个函数某一个参数可以省略
Function shuju(maxnum, geshu, optional qo As Integer)
    Dim d As New dictionary
    Dim num
    application.volatitle
    Do
        num = int(Rnd() * maxnum + 1)
        If qo = o Then
            d(num) = ""
        ElseIf qo Mod 2 = 0 Then
            d(num) = ""
        ElseIf qo Mod 1 = 1 Then
            d(num) = ""
        End If
    Loop Until d.count = geshu
    shuju = application.transpose(d.keys)
End Sub
'当qo为0时，不区分奇偶数，1为取奇数，2为取偶数

Function shuju(maxnum, geshu, optional qo As Integer = 0)
    Dim d As New dictionary
    Dim num
    application.volatitle
    Do
        num = int(Rnd() * maxnum + 1)
        If qo = 0 Then
            d(num) = ""
        ElseIf qo Mod 2 = 0 Then
            d(num) = ""
        ElseIf qo Mod 1 = 0 Then
            d(num) = ""
        End If
    Loop Until d.count = geshu
    shuju = application.transpose(d.keys)
End Sub
'与上面的区别是qo有一个默认值，可以省略
