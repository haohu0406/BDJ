Mid(要处理的字符串 ，开始的字符 【包含本身 】，要提取的字符个数)
Sub tt1()
    k = Mid(Range("h17"), 1, 0)
    MsgBox k
End Sub
'返回的是空
Mid获得的数据类型是string需要通过0 + mid转换成数值