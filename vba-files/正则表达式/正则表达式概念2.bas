vba正则表达式入门

1、什么是正则表达式
正则表达式是一个天才创建的用于快速检索匹配字符串 ，通过简单的表达式匹配文本 。



2、正则表达式的组成
正则表达式也是一个字符串 ，包括元字符 、限定符和正常意义的字符 。正则表达式强大的地方就在元字符和限定符 。



3、限定符
很多人讲这个都是先讲元字符 。其实先讲限定符更加容易吸收 。限定符是表示前面字符或元字符出现的次数 。主要限定符如下 ：

限定符 含义
 * 表示前面字符或元字符出现0次或多次 。例如 ：zo * m ，可以匹配zm ，zom ，zoom
 + 表示前面字符或元字符出现1次或多次 。例如 ：zo + m ，可以匹配zom ，zoom
？表示前面字符或元字符出现0次或1次 。例如 ：zo ?m ，可以匹配zm ，zom
{n} 表示前面字符或元字符出现n次 。例如 ：zo{3} m ，可以匹配zooom
{n, m} 表示前面字符或元字符出现n到m次 。例如 ：zo{1, 3} m ，可以匹配zom ，zoom ，zooom
{n, } 表示前面字符或元字符至少出现n次 。例如 ：zo{2, } m ，可以匹配zoom ，zooom等等


当然 ，限定符用法不止这些 。正则表达式还有个规则叫贪婪与吝啬 。有些人也叫贪婪与懒惰 。这个 “贪婪与吝啬 ”是正则表达式难点和重点之一 。

例如 ，字符串 “n123n456n789n ”。那么我们怎么匹配获取 “n123n ”和 “n123n456n789n ”呢 。

使用正则表达式 ，首先要找规律 。很明显我们要获取的内容开头和结尾都有一个字母n ，中间是数字或字母 。我们先学一个元字符 。元字符是可以代表一定含义或规律的字符 。可以匹配除了换行符之外的任意字符是小数点 。

那么我们的表达式可以这么写 ：n. + n

两个字母n ，中间夹着1个或多个任意字符 。

但这样只能匹配得到一个结果 ：n123n456n789n 。在我们没有对其任何限制的情况下 ，正则表达式会尽可能多匹配符合条件的结果 。从头到尾整个都符合 ，所以都匹配了 。这个称之为贪婪匹配 。

那如何做到尽可能少的匹配 。这个就需要加个 ?进行限制 。

例如 ，表达式 ：n. + ?n

这个表达式尽可能少匹配 ，也就是说碰到一次符合条件的就立马返回结果 。结果可以匹配到 “n123n ”、“n789n ”。

这种规则叫做吝啬匹配 。只要在限定符后面再加个问号即可 。



4、元字符
元字符是用于匹配字符串 ，可以代表一定含义或规律的字符串 。主要的元字符如下 ：

元字符 含义
.小数点 ，代表除了换行符以外的任意字符
 \ 转义 ，若我想匹配一些被正则表达式占用的字符 ，例如小数点 ，可以用 \ .
[abc] 匹配中括号内的字符 ，例如[a - zA - Z] ，可以匹配到大小写字母
[ ^ abc] 不匹配中括号内的字符 ，例如[ ^ a - z] ，表示不匹配小写字母
 \ w 可以匹配字母 、下划线和数字 ，相当于[a - zA - Z0 - 9 _ ]
         \ W 大写的W是小写的w相反情况 ，也就是不匹配字母 、下划线和数字 。相当于[ ^ a - zA - Z0 - 9 _ ]
         \ s 匹配任意空白符 ，相当于[ \ f \ n \ r \ t \ v]
 \ S 匹配任意非空符 ，相当于[ ^ \f \ n \ r \ t \ v] 或[ ^ \s]
 \ d 匹配数字 ，相当于[0 - 9]
 \ D 匹配非数字 ，相当于[ ^ 0 - 9]
 \ b 匹配单词的边界 。这个匹配英文单词特别有用 。例如 \ b[ \ w ']+?\b就可以匹配任意单词了
 \ f 匹配换页符
 \ n 匹配换行符
 \ r 匹配回车符
 \ t 匹配tab制表符
 \ v 匹配垂直制表符
 ^ 不在中括号内的 ^ ，表示从字符串的开头开始匹配
$表示匹配到字符串的结尾
x |y 匹配x或y
(表达式) 元组 ，用小括号括起来的表达式当作一个元组 ，可以当作一个整体 ，也可以被 \ 1 \ 2 \ 3 这样类似索引获取 。


元字符比较多 ，这里就建议大家先收藏 ，需要用的时候再查阅 。多用几次就自然记住了 。

这里还有个小技巧 ，若我想匹配全部任意字符 ，包括换行符 。可以用一组相反的元字符 ，例如[ \ s \ S] ，就可以匹配全部任意字符 。



5、常见的正则表达式

说了这么多 ，晕了没 ？看一些实例 ：

1）匹配邮编 ，邮编是6位数字 。正则表达式 ： \ d{6}

2）匹配手机 ，手机号是11位数字 。正则表达式 ： \ d{11}

3）匹配电话 ，电话是区号 - 号码组成 ，区号有3到4位 ，号码有6到9位 。正则表达式 ： \ d{3, 4} - \d{6, 9}

4）匹配日期 ，日期格式如1992 - 5 - 30 ，明显数字加横线组成 。正则表达式 ： \ d{4} - \d{1, 2} - \d{1, 2}

5）匹配汉字 ，汉字需要通过编码转义 ，汉字都unicode编码中都在一个范围内 。正则表达式 ：[ \ u4e00 - \u9fa5]



6、vba中使用正则表达式

若只是上面这些内容 ，那么还是纸上谈兵 ，需要应用到实际中 。看看如何在vba中使用正则表达式 。

vba使用正则表达式需要用到一个RegExp对象 。

该对象可以通过引用Microsoft VBScript Regular Expressions 5.5 。再声明定义 ：

Dim reg As New RegExp
还可以直接用CreateObject方法创建 ：

Dim reg As Object
Set reg = CreateObject("VBScript.Regexp")


创建RegExp对象之后 ，看看它的相关属性和方法 。

属性 ：

1）Global ，是否全局匹配 ，若为False ，匹配到一个结果之后 ，就不再匹配 。默认False ，建议设为True ；

2）IgnoreCase ，是否忽略大小写 ，默认False ，建议设为False ，这个会影响到正常表达式匹配 ；

3）Multiline ，是否跨行匹配 ，默认False ，建议设为False ，这个会影响到正常表达式匹配 ；

4）Pattern ，获取或设置正则表达式 。



方法 ：

1）Execute ，执行匹配

2）Replace ，根据正确表达式全部替换

3）Test ，测试正则表达式能否匹配到内容



举一些典型的例子 ：

1）判断是否存在数字

Public Function CheckNumber(Str As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
    
    Dim is_exist As Boolean    
    With reg
         .Global = True
         .Pattern = "\d"        
        is_exist =  .Test(Str)        
    End With
    CheckNumber = is_exist
End Function
用Test方法 ，判断能否匹配到数字 。



2）获取所有编号

Public Sub GetCode()
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
    
    Dim Str As String
    Str = "编号：ABC123155 日期：2016-01-11" &  _
            "编号：ABD134215 日期：2016-02-21" &  _
            "编号：CBC134216 日期：2016-01-15"
    
    reg.Global = True reg.Pattern = "[A-Z]{3}\d+" '获取匹配结果'
    Dim matches As Object, match As Object
    Set matches = reg.Execute(Str)
    
    '遍历所有匹配到的结果'
    For Each match In matches
        '测试输出到立即窗口'
        Debug.Print match
    Next
End Sub
因为这个编号是3个大写字母和多个数字组成 。可以利用代码中的表达式匹配到3个结果 ：ABC123155 、ABD134215和CBC134216 。



3）去掉字符串中的数字

Public Function ClearNumber(Str As String) As String
    Dim reg As Object    
    Set reg = CreateObject("VBScript.Regexp")    
    
    reg.Global = True    
    reg.Pattern = "\d"    
    
    '把所有数字替换成空'    
    ClearNumber = reg.Replace(Str, "")
End Function
执行ClearNumber函数 ，即可去掉数字 。例如ClearNumber("你342好234啊") ，可得到 "你好啊" 。



4）获取子字符串

例如想获取某些字符串中的部分数据 ，可以匹配完成之后 ，再用字符串函数处理 。但其实不用 ，用元组可以一次性搞定 。

Public Sub GetHref()    
    Dim reg As Object    
    Set reg = CreateObject("VBScript.Regexp")    
    
    Dim Str As String    
    Str = "xxx1xxx2"    
    
    reg.Global = True    
    '获取a标签中href的属性值'    
    reg.Pattern = "href='(.+?)'"    
    
    '获取匹配结果'    
    Dim matches As Object, match As Object    
    Set matches = reg.Execute(Str)    
    
    '遍历所有匹配到的结果'    
    For Each match In matches        
        '测试输出子表达式到立即窗口'        
        Debug.Print match.SubMatches(0)
    Next
End Sub
这里 ，可以通过match的SubMatches集合获取元组里面的内容 。轻松得到xxx1和xxx2 。
