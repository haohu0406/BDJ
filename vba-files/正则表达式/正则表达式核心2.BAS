RegExp对象是VBScript中用于提供简单地正则表达式支持的对象 。VBScript中所有和正则表达式有关的属性和方法都有这个对象有关联 。

一 、RegExp对象的属性和方法 （三个属性 ，三个方法 ）

属性 ：

1. Global属性: True Or False, 指明模式是匹配整个字符串中所有与之相符的地方还是只匹配第一次出现的地方 。默认值是False(表示只匹配第一次出现的地方) 。

2. IgnoreCase属性: True Or False, 指明模式匹配是否大小写敏感 。默认是False(表示搜索为大小写敏感) 。

3. Pattern属性 ：设置或返回用于搜索的正则表达式 。

方法 ：

1. Execute方法 ：将正则表达式应用到字符串上并返回Matches集合 。

2. Replace方法 ：替换在正则表达式搜索中找到的文本 。

3. Test方法 ：对字符串执行正则表达式搜索 ，并返回一个布尔值说明匹配是否成功(True:匹配成功, False:匹配不成功) 。

二 、VBScript中的正则表达式字符

字符 描述
 \ 表示下一个字符是特殊字符或文字常量
 ^ 匹配输入的开头
$匹配输入的结尾
 * 匹配前一个字符零次或多次
 + 匹配前一个字符一次或多次
?匹配前一个字符零次或一次
.匹配除换行符以外的任何单个字符
(pattern)
匹配并记住这个模式 。可以用[0] …[n] 从结果的matches集合中获取匹配到的字符串 。要匹配括号本身 ，在前面加上斜杠 - 用 "\(" 或 "\)"

(?:pattern) 匹配但不捕获模式 ，也就是不会存储匹配结果供以后使用 。这可以用于使用 "or" 字符(|) 合并模式的不同部分 。例如 ，"anomal(?:y|ies)" 比 "anomaly|anomalies" 要划算得多
(? = pattern) 当所要搜索的字符串匹配了模式的开头部分时就接着匹配这一部分 。非捕获匹配 ，也就是说不会保存匹配结果供以后使用 。例如 ，"Windows (?=95|98|NT|2000|XP|Vista)" 能匹配 "Windows Vista" 中的Windows而不能匹配 "Windows 3.1" 中的Windows
(?!pattern) 与上一个相反 ，这会匹配模式中没有出现的内容 。非捕获匹配 。例如 ，"Windows (?！95|98|NT|2000|XP|Vista)" 能匹配 "Windows 3.1" 中的Windows而不能匹配 "Windows Vista" 中的Windows
x |y 匹配x或y
{n} 准确地匹配n次(n必须是一个非负整数)
{n, } 至少匹配n次(n必须是一个非负整数 - 注意结尾的逗号)
{n, m} 至少匹配n次 ，最多匹配m次(m和n必须都是非负整数)
[xyz] 匹配其中包括的任一个字符(xyz表示一个字符集)
[ ^ xyz] 匹配其中不包括的字符( ^ xyz表示一个字符集的补集)
[a - z] 匹配指定范围内的字符(a - z表示字符的范围)
[ ^ m - z] 匹配指定范围以外的字符( ^ m - z表示指定范围的补集)
 \ b 匹配一个单词边界 ，这个位置在单词和空格之间
 \ B 匹配一个非单词边界
 \ d 匹配数字 。等价于[0 - 9]
 \ D 匹配非数字 。等价于[ ^ 0 - 9]
 \ f 匹配换页符
 \ n 匹配换行符
 \ r 匹配回车符
 \ s 匹配空白 ，包括空格 、制表符 、换页符等 。等价于 "[\f\n\r\t\v]"
 \ S 匹配非空白的字符 。等价于 "[^\f\n\r\t\v]"
 \ t 匹配制表符
 \ v 匹配纵向制表符
 \ w 匹配字母 、数字 ，以及下划线 。等价于 "[A-Za-z0-9_]"
 \ W 匹配非字符数字 。等价于 "[^A-Za-z0-9\_]"
 \ .匹配.
 \ |匹配 |
 \ {匹配{
 \ } 匹配}
 \ [匹配[
 \ ] 匹配]
 \ (匹配(
 \ ) 匹配)
$num 匹配num ，其中num是正整数 。返回匹配结果的引用
 \ n 匹配n ，其中n是八进制转义符 。八进制转义符的长度应为1 、2 或3
 \ uxxxx 匹配UNICODE形式的ASCII字符
 \ xn 匹配n ，其中n是十六进制转义符 。十六进制转义符必须是两位长度
 \ \匹配 \
三 、实例

1、验证电话号码输入 ：

Dim re, s, objMatch, colMatches
Set re = New RegExp
re.Pattern = "\([0-9]{3}\)[0-9]{4}-[0-9]{4}"
re.Global = True
re.IgnoreCase = True
s = InputBox("Enter your phone number in the following Format (XXX)XXXX-XXXX:")
If re.Test(s) Then
    MsgBox "Thank you!"
Else
    MsgBox "Sorry but that number is not in a valid format."
End If
2、分解URI ：将URI分解成协议 （如ftp 、http等 ）、域名地址 、页面和路径四个部分 。

Dim re, s
Set re = New RegExp
re.Pattern = "(\w+):\/\/([^/:]+)(:\d*)?([^#]*)"
re.Global = True
re.IgnoreCase = True
s = "http://www.wrox.com:80/misc-pages/support.shtml"
MsgBox re.Replace(s, "$1")
MsgBox re.Replace(s, "$2")
MsgBox re.Replace(s, "$3")
MsgBox re.Replace(s, "$4")
3. 检查HTML元素 ："<(.*)>.*<\/\1>"

4. 匹配空白 ："^[\t]*$"

5. 匹配HTML注释标签 ："^.*<!--.*—>.*$"

将上述匹配模式稍作修改 ，匹配HTML注释标签并清除HTML注释 ：

Dim re, s
Set re = New RegExp
re.Pattern = "(^.*)(<!--.*-->)(.*$)"
re.Global = True
s = "<title>A Title</title><!-- a title tag -->"
If re.Test(s) Then
    MsgBox "HTML comment tags found."
Else
    MsgBox "No HTML comment tags found."
End If
MsgBox re.Replace(s, "$1" & "$3")