使用方法
1、引用法
vbscript regular expressions
Dim regex As New regexp
2、直接创建发
Dim regex As Object
Set regex = CreateObject("vbscript.regexp")


属性
1、global属性
如果为true, 则搜索全部字符
如果为false ，则搜索到第一个即停止
例如 ：
Sub t3()
    Dim reg As New regexp
    Dim sr
    With reg
         .global = True
         .patter = "a"
        debug.Print  .Replace(sr, "")
    End With
End Sub

2、ignorecase属性
区分大小写 ，为false(缺省值) 分 ，True不分

3、pattern属性
一个字符串 ，用来定义正则表达式 ，缺省值(不设置) 为空文本

4、multiline属性 ，字符串是不是使用了多行 ，如果是多行 ，$适用于每一行的最后一个
多行则$会判断每行的结尾

5、execute方法
返回一个matchcollection对象 ，该对象包含每个成功匹配的match对象
返回的信息包括
firstindex ：开始位置
length ：长度
value ：值
Sub t5()
    Dim reg As New regexp
    Dim sr, matc
    With reg
         .global = True
         .pattern = "a\d+"
        Set matc =  .excute(sr)
    End With    
    Stop
End Sub

Function ns(rg)
    Dim reg As New regexp
    Dim sr, ma, s, m, x
    With reg
         .global = True
         .pattern = "\d*\.?\d*"
        Set ma =  .execute(rg)
        For Each m In ma
            s = s + Val(m)
        Next
    End With    
    ns = s
End Function

text方法
返回一个布尔值 ，该值指示正则表达式是否与字符串成功匹配 ，其实就是判断两个字符是否
匹配成功
Sub t7()
    Dim reg As New regexp
    Dim sr
    sr = "bcr6ea"
    With reg
         .global = True
         .pattern = "\d+"
        If  .test(sr) Then MsgBox "字符串中含有数字"
    End With    
End Sub
