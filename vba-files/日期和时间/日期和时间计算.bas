1.计算相隔天数 ，月数 ，年数 ，小时 ，分钟 ，秒
Sub tt1()
    Dim d1, d2 As Date
    d1 = #11 / 21 / 2011#
    d2 = #12 / 1 / 2011#
    Debug.Print "相隔" & (d2 - d1) & "天"    
    Debug.Print dateiff("d", d1, d2) & "天"
    Debug.Print dateiff("m", d1, d2) & "月"
    Debug.Print dateiff("yyyy", d1, d2) & "年"
    Debug.Print dateiff("q", d1, d2) & "季度"
    Debug.Print dateiff("w", d1, d2) & "周"
    Debug.Print dateiff("h", d1, d2) & "小时"
    Debug.Print dateiff("n", d1, d2) & "分钟"
    Debug.Print dateiff("s", d1, d2) & "秒"
End Sub

2.计算两个时间的差
Sub tt2()    
    Dim t As Date
    t = Timer
    For i = 10000        
    Next    
    Debug.Print Timer - t
End Sub

3.日期时间加减
Sub tt3()
    Dim d1, d2 As Date    
    d1 = "2001-10-1 00:00:00"
    Debug.Print VBA.DateAdd("d", 10, d1) '加上10天
    Debug.Print VBA.DateAdd("m", 10, d1) '加上10月
    Debug.Print VBA.DateAdd("yyyy", 10, d1) '加上10年
    Debug.Print VBA.DateAdd("yyyy",  - 10, d1) '减去10年
    Debug.Print VBA.DateAdd("h", 10, d1) '加上10小时
    Debug.Print VBA.DateAdd("n", 10, d1) '加上10分钟
    Debug.Print VBA.DateAdd("s", 10, d1) '加上10秒
End Sub

4.获取当前月份的天数

Sub tt6()
    Dim y, m, d
    y = Year(Now)
    m = Month(Now)
    MsgBox Day(DateSerial(y, m + 1, 0))
End Sub


Dim myNow As Date, BL As Integer '定义myNow为日期型;定义BL为长整型 
myNow = Now '把当前的时间赋给变量myNow
Do '开始循环语句Do
    BL = Second(Now) - Second(myNow) '循环中不断检查变量BL的值
    If BL = 30 Then GoTo Cl '当BL=30即跳转到CL
    DoEvents '转让控制权,以便sheets可继续操作
Loop Until BL > 30 '当BL>30即跳出循环
Exit Sub
