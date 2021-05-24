RegExp属性
Global
IgnoreCase
Pattern

RegExp方法
Execute - -对指定的字符串执行正则表达式搜索 。
Replace - -替换正则表达式搜索到的字符 。
Test

RegExp对象
Match
Match属性
FirstIndex
Length
Value

RegExp集合
Matches
SubMatches

说明 ：

正则表达式搜索的设计模式是通过 RegExp 对象的 Pattern 来设置的 。

Execute 方法返回一个 Matches 集合 ，其中包含了在 String 中找到的每一个匹配的 Match 对象 。如果未找到匹配 ，Execute 将返回空的 Matches 集合 。

举例 ：

Function RegExpTest(patrn, strng)
    　　Dim regEx, Match, Matches ‘建立变量
    　　Set regEx = New RegExp ‘建立正则表达式
    　　regEx.Pattern = patrn ‘设置搜索方法
    　　regEx.IgnoreCase = True ‘设置是否区分大小写
    　　regEx.Global = True ‘设置全程可用性
    　　Set Matches = regEx.Execute(strng) ‘根据正则表达式规则执行搜索字符串
    　　For Each Match In Matches ‘遍历Matches集合
        　　RetStr = RetStr & Match.Value & "," ‘显示符合正则表达式规则的字符 ，此句也可写为 ：RetStr = RetStr & Match        
        　　Next
    　　RegExpTest = RetStr
End Function
Str = RegExpTest("\d+", "xxafaf12dfasf3433432xx你好")
MsgBox Str
显示 ：12, 3433432

View Code
　　在这个示例中我们可以看到方法Execute和对象Match是使用在集合Matches中的 ，Match和Matches是作为变量来使用的 ，为了我们很容易得看懂它 ，我们没有必要将他们换名字 。关于对象Match的属性 ，我在上面的示例中都做了注释了 。

RegExp的Replace方法介绍 ：

Replace - -替换在正则表达式搜索到的文本 。

Object.Replace(string1, string2)

参数 ：

Object

Required.Always the Name of a RegExp Object.

string1

Required.String1 Is the text String In which the text replacement Is To occur.

string2

Required.String2 Is the replacement text String.

说明 ：

被替换的文本的实际模式是通过 RegExp 对象的 Pattern 属性设置的 。

Replace 方法返回 string1 的副本 ，其中的 RegExp.Pattern 文本已经被替换为 string2 。如果没有找到匹配的文本 ，将返回原来的 string1 的副本 。

下面的例子说明了Replace方法的用法 。


Function ReplaceTest(patrn, replStr)
    Dim regEx, str1
    str1 = “the quick browm fox jumps over the lazy dog.”
    ‘create regular expression
    Set regEx = New RegExp
    regEx.Pattern = patrn
    regEx.IgnoreCase = True
    ‘make replacement.
    ReplaceTest = regEx.Replace(str1, replStr)
End Function
MsgBox(ReplaceTest(“fox ”, ”cat ”)) ‘Replace ‘fox ’With ‘cat ’.

RegExp的Test方法介绍:
    
    Test 方法
    
    对指定的字符串执行一个正则表达式搜索 ，并返回一个 Boolean 值指示是否找到匹配的模式 。
    
    Object.Test(String)
    
    参数
    
    Object
    
    必选项 。总是一个 RegExp 对象的名称 。
    
    String
    
    必选项 。要执行正则表达式搜索的文本字符串 。
    
    说明
    
    正则表达式搜索的实际模式是通过RegExp对象的Pattern属性来设置的 。RegExp.Global属性对Test方法没有影响 。
    
    如果找到了匹配的模式 ，Test方法返回True ；否则返回False 。
    
    下面的代码说明了Test 方法的用法 。

    Function RegExpTest(patrn, strng)
        Dim regEx, retVal ' 建立变量。
        Set regEx = New RegExp ' 建立正则表达式。
        regEx.Pattern = patrn ' 设置模式。
        regEx.IgnoreCase = False ' 设置是否区分大小写。
        retVal = regEx.Test(strng) ' 执行搜索测试。
        If retVal Then
            RegExpTest = "找到一个或多个匹配。"
        Else
            RegExpTest = "未找到匹配。"
        End If
    End Function
    MsgBox(RegExpTest("is.", "IS1 is2 IS3 is4"))

    SubMatches 集合
    
    正则表达式子匹配字符串的集合 。
    
    说明
    
    SubMatches 集合包含了单个的子匹配字符串 ，只能用 RegExp 对象的 Execute 方法创建 。SubMatches 集合的属性是只读的 。
    
    运行一个正则表达式时 ，当圆括号中捕捉到子表达式时可以有零个或多个子匹配 。SubMatches 集合中的每一项是由正则表达式找到并捕获的的字符串 。
    
    下面的代码演示了如何从一个正则表达式获得一个 SubMatches 集合以及如何它的专有成员 ：
    
    
    Function SubMatchTest(inpStr)
        　　Dim oRe, oMatch, oMatches
        　　Set oRe = New RegExp
        　　'查找一个电子邮件地址（不是一个理想的 RegExp）
        　　oRe.Pattern = "(\w+)@(\w+)\.(\w+)"
        ‘得到 Matches 集合
        Set oMatches = oRe.Execute(inpStr)
        ‘得到 Matches 集合中的第一项
        Set oMatch = oMatches(0)
        ‘创建结果字符串 。
        Match 对象是完整匹配 —dragon @ xyzzy.com
        retStr = "电子邮件地址是： " & oMatch & vbNewLine
        ‘得到地址的子匹配部分 。
        retStr = retStr & "电子邮件别名是： " & oMatch.SubMatches(0) ' dragon
        　　retStr = retStr & vbNewLine
        　　retStr = retStr & "组织是： " & oMatch.SubMatches(1) ' xyzzy
        　　SubMatchTest = retStr
    End Function
