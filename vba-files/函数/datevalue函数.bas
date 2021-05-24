[已解决后总结] ：原来VBA内置的DATEVALUE() 也可以很强大 
，但是不如EXCEL内置的DATEVALUE() 识别的文本型日期范围广比如 ：
1.）在VBA里面
Dim x As String,
x = “16 SEP2014 ”
y = DateValue(x)
会报错 “Run - Time Error '13': Type mismatch”
但是如果用 y = DateValue(Left(x, 2) & "-" & Mid(x, 3, 3) & "-" & Mid(x, 6, 4)) 则不会报错 ；说明VBA的DateValue可以识别 “16 - SEP - 2014 ”这种文本型日期 ，但是无法识别 “16 SEP2014 ”这种 。


2.）在Excel里面
[B1] = “16 SEP2014 ”
C1里面输入公式 = VALUE(B1) 或者 = DateValue(B1) 都可以得到值 41898
在C1里面输入公式 = VALUE(Left(B1, 2) & "-" & Mid(B1, 3, 3) & "-" & Mid(B1, 6, 4)) 或者 = DateValue(Left(B1, 2) & "-" & Mid(B1, 3, 3) & "-" & Mid(B1, 6, 4))
同样都可以得到值 41898
http: / /club.excelhome.net / thread - 1170800 - 1 - 1. html