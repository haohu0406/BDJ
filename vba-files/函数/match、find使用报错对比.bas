VBA ，使用find() 和 match() 进行查找时 ，可能出现的各种错误 （我犯的各种错误总结 ）


1 正确的代码
应该也有多种写法

Sub test5032()
    Dim a As Range
    in1 = InputBox("请输入一个电影名")
    Set a = Range("a1:a15").Find(in1)
    If a Is Nothing Then
        Range("A" & 1 + Range("a65536").End(xlUp).Row).Value = in1
    Else
        MsgBox "内容重复了"
    End If
End Sub


2 典型错误1 ：判断 range().find() 返回值乱用null, err这些
2.1 使用 range().If () 查找时 ，对其返回值不合适的 判断函数
2.1    .1 这3种用法的结果不同
    range().find()
    如果查得到 ，返回的是rang()
    如果查不到 ，返回的是nothing()
    
    
    判断条件只能是 ：    
    Dim a As range
    Set a = range().find()    
    If a Is Nothing        
    语法不能写成 If a = Nothing
a.value Is Nothing 也报对象变量错误 ，因为nothing 是专门给对象变量用的


错误的判断条件 ，查不到返回 Nothing 而不是 Err, 用 a = err等判断 ，始终都会为否
If a = Err Then
    If IsError(a) Then
    
    
    错误的判断条件 ，不明白为啥 Nothing 不是 Null 或 Empty ?用isnull() ispempty() 判断始终都为否
    都是判断变量的值的函数
    If IsNull() Then
If IsEmpty() Then    
    
    
2.1    .2 错误代码举例 ：如果用 IsNull() 和 IsEmpty(), 下面的if 始终只会判断为否
    错误原因
    
    Empty 有效的空值 ，如0 ""
    Null 无效的控制 ，比如二选一之外的空
    Nothing 对象变量的空值
    
    
    Sub test50411()
Dim a As Range
in1 = InputBox("请您输入1个电影名")
Set a = Range("a1:a15").Find(in1)

If IsNull(a) Then
    Range("A" & 1 + Range("a65536").End(xlUp).Row) = in1
Else
    MsgBox "内容重复了"
End If
    End Sub
    
    
2.1    .3 Empty Null Nothing 三者的区别
    参考文章 ：https: / /www.cnblogs.com / lys_013 / archive / 2013 / 04 / 16 / 3024229. html
    Empty 、Null和Nothing都可为Variant变量赋值
    声明时系统会设Variant变量为Empty
    如果要将Variant变量设为无效数据可用Null
    如果不再使用对象变量就应尽快将之设成Nothing以利系统释放资源 。
    
    
    3 典型错误 ：使用range().find() 必须先做错误处理 ，再处理正常流程
    3.1 range().find() 如果查不到 ，或者查找的类型不匹配会报错
    因为 range().find() 不事先处理报错导致的报错
    典型错误 ，没有先做异常处理会导致报错 ，想到find() 函数的应该先进行出错处理 ，再处理正常流程
    
    
3.1    .1 思路的学习 ，以后用 find() 和if 要考虑全面点 ，先考虑出错可能和如何处理出错
    新思路 ：应该先考虑异常逻辑和进行异常处理 ，然后再处理正常逻辑 。
    
    分析出错的原因 ：
    我的老思路 ：用 find() 来查找 ，先判断是否重复 ，如果重复就报错 ，如果不重复就新增1条记录
    问题是
    range().find() 如果查找不到会报错 。
    所以应该先判断 ，是否会出错 ，然后马上解决出错后如何处理 。然后再去处理正常逻辑 。
    
    
    以后使用range().find() 必须先做错误处理 ，否则可能会先跳出报错而中断
    先要考虑出错的可能
    还要想到如何处理出错的手段
    而不是上来开始正常的逻辑过程
    
    
3.1    .2 中间变量挺重要
    中间变量很重要
    不光是把一些数据存起来
    而且防止了传送过程中数据发生变化
    
    
3.2    .1 错误代码
    Sub test50323() '这是一个错误代码，举例
Dim a As Range
in1 = InputBox("请输入一个电影名")
Set a = Range("a1:a15").Find(in1)
If Range("a1:a15").Find(in1) = in1 Then '一旦这个找不到就会报错
    MsgBox "内容重复了"
Else
    Range("A" & 1 + Range("a65536").End(xlUp).Row).Value = in1
End If
    End Sub
    
    
    
    
    
    
3.2    .2 修正代码 - --还是错的 - --不知道为什么还是只走一个分支
    加了 On Error Resume Next
    If 先判断了find() 到
结果还是只走上面的分支 ？
Sub test50323()
    Dim a As Range
    in1 = InputBox("请输入一个电影名")
    Set a = Range("a1:a15").Find(in1)
    
    On Error Resume Next
    
    Debug.Print "Range("" a1: a15 "").Find(in1)="
    Debug.Print "Range("" a1: a15 "").Find(in1)=" & Range("a1:a15").Find(in1).Address '因为出错，整句都不会打印？
    
    If Range("a1:a15").Find(in1) = in1 Then '不知道为什么这么改了，为什么还只有一个分支？
MsgBox "内容重复了"
    Else
Range("A" & 1 + Range("a65536").End(xlUp).Row).Value = in1
    End If
End Sub


4 典型低级错误
4.1 使用语法错误 ：设置变量对象时 ，赋值语法不对 ，没有用set
对象变量没设置好也会报错 ，也是报错 ，对象变量或with变量未设置
虽然报错一样 ，但是和 range().find() 方法查不到而报错 ，本质不一样




Sub test50311()
    
    ' 典型错误，没有保存变量的
    Dim a As Range
    Set a = Range("a1:a15").Find(InputBox("请输入一个电影名"))
    If a Is Nothing Then
Range("A" & 1 + Range("a65536").End(xlUp).Row).Value = InputBox("请输入一个电影名") 'TextBox1 不在窗体里报错，不应该在initial做歌词初始化
    Else
MsgBox "内容重复了"
    End If
    
End Sub


4.2 典型错误 ：多用中间变量保存数据 ，然后再传递和使用 ，而不是每次临时处理 。
我这又犯了新手常犯的错误
输入的数据应该用变量保存下来 ，装到变量盒子里
典型错误 ，没有保存变量值 ，多次使用inputbox ，这样是不对的
Sub test50311()
    
    ' 典型错误，没有使用中间变量保存和传递数据
    ' 错误的使用2次inputbox() ，第2个input和第一次木关系了
    ' 典型新手会犯的错误
    
    Dim a As Range
    Set a = Range("a1:a15").Find(InputBox("请输入一个电影名"))
    If a Is Nothing Then
Range("A" & 1 + Range("a65536").End(xlUp).Row).Value = InputBox("请输入一个电影名") 'TextBox1 不在窗体里报错，不应该在initial做歌词初始化
    Else
MsgBox "内容重复了"
    End If
    
End Sub


4.3 典型错误 ：用 分支判断时 ，谁的值 ？
If Err Then ’典型错误 ，这种纯粹SX的错误 If 谁 ? = Err ? Then
    If a = Err Then '记得应该是判断某个变量，或者 表达式的值 为err 
另外 False <> Err
If True Then '这是故意的写法，就是故意只执行false的分支！
    
    
    不要随便混用
    True False '这个是bool值
    Error '错误值
    
    
    5 不合适的功能函数 ，使用application.find() 导致的错误
    5.1 application.find() 和worksheetfunction.find() 本质是工作表函数 - --查找的范围主要还是 字符串内string
    如果使用的功能函数 ，是工作表函数 ，会出问题
    工作表函数 ，application.find() 返回的是 查找内容在 字符串 String 内的 第几个位置 。
    
    
    application.find() 返回值 ，用 IsEmpty() IsError() = Null = Err 判断了都不行 ，不知道怎么搞
    Sub test504() - --错误代码 ，举例

On Error Resume Next
in1 = InputBox("请您输入1个电影名")
a = Application.Find(in1, Range("a1:a15"))

If IsError(a) Then '尝试过 isempty()  =err  =null 等都不行
    Debug.Print "没找到"
    Range("A" & 1 + Range("a65536").End(xlUp).Row) = in1
Else
    Debug.Print "找到了"
    MsgBox "内容重复了"
End If

    End Sub
    
    
    6 使用match() 函数也会遇到类似的问题
    application.match() 本质也是工作表函数
    和 range().find() 不一样 ，rang().find(）返回的是 range() 对象 ，find不到位空对象变量是 Is Nothing
    而application.match() 查到的是一个index 对应的数 。如果查不到会报错 ，而报错是err = 0
    6.1 错误的写法
    错误的思路 ：没有先考虑出错处理
    application.match() 查不到会报错 ，无法取得 match 属性
    
    
    
    
    正确代码
    使用查找 ，匹配函数时 ，要先进行错误处理
    但同时要对错误值和空值有区分 ，
    如果需要处理错误值 ，则要可能要先加一个语句 On Error Resume Next
    如果要处理空值 ，得注意是对象变量是 Is Nothing, 其他变量 则可能是 Empty 或者 Null 等
    Sub aa31()
On Error Resume Next
in1 = Int(InputBox("请输入1个要查的数字"))
a = WorksheetFunction.Match(in1, Array(1, 2, 3, 4, 5), 0)

If Err = 0 Then
    Debug.Print a
Else
    Debug.Print "没找到!"
End If
    End Sub
    err的具体值
    最初猜想 ：err是布尔值 ？Err = 0 表示没错误 ，Err = 1 表示是出错了
    测试代码 ，换ERR = 1 ，且调换IF判断次序 ，结果找不到就不反馈 ，有问题
    Sub aa32() - --测试代码 ，换ERR = 1 ，且调换IF判断次序 ，结果找不到就不反馈 ，有问题
On Error Resume Next
in1 = Int(InputBox("请输入1个要查的数字"))
a = WorksheetFunction.Match(in1, Array(1, 2, 3, 4, 5), 0)

If Err = 1 Then
    Debug.Print "没找到!"
Else
    Debug.Print a
End If
    End Sub
    
    
    Err 其实是 Err.number
    Err = 0 表示 Err.number = 0 没出出错
    Err.number有很多出错数字