1.返回当前日期 、时间
Sub t1()
    debug.Print Date '返回当前日期
    debug.Print Time '返回当前时间
    Debug.Print Now '返回当前日期+时间    
End Sub

2.格式化显示日期
Sub t2()
    Debug.Print Format(Now, "yyyy-mm-dd")    
    debug.Print Format(Now, "yyyy年mm月dd日")
    Debug.Print Format(Now, "yyyy年mm月dd日 hh:mm:ss")
    debug.Print Format(Now, "d-mmm-yy") '英文月份3-dec-11
    debug.Print Format(Now, "d-mmmm-yy") '3-december-11
    debug.Print Format(Now, "aaaa") '中文星期
    debug.Print Format(Now, "ddd") '英文星期前三个字母
    Debug.Print Format(Now, "dddd") '英文星期完整显示
End Sub

3.根据年月返回日期
Sub t1()
    debug.Print VBA.DateSerial(2011, 10, 1)    
End Sub

4.根据小时分钟返回时间
Sub t2()
    debug.Print VBA.TimeSerial(1, 2, 1)
End Sub

5.返回年月日小时分钟
Sub t5()
    d = "2011-10-28 01:10:03"
    debug.Print Year(d) & "年"
    debug.Print Month(d) & "月"
    debug.Print Day(d) & "日"
    debug.Print Hour(d) & "时"
    debug.Print VBA.Minute(d) & "分"
    debug.Print Second(d) & "秒"
End Sub

6.获取当年年份, 月份
Year(Now)
Month(Now)