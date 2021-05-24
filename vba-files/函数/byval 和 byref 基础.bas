ByVal 和 ByRef 基础
在定义过程或函数时 ，如果需要传递变量 ，则每个参数需要指定传递类型 。
传递类型有 2 种 ，分别是 ByVal 和 ByRef 。

'ByVal 传递类型
Sub TestSub1(ByVal msg As String)
    
End Sub

'ByRef 传递类型
Sub TestSub2(ByRef msg As String)
    
End Sub
针对基础数据类型 ，例如数字 、文本等 ，两种传递类型的说明和区别如下 ：

ByVal ：传递变量时 ，复制一份该变量 ，传入过程或函数 。在过程和函数内部
对该变量进行修改 ，只对该副本有效 ，对上一级过程 （父过程 ）的变量没有影响 。
ByRef ：传递变量时 ，将该变量的引用地址传入过程或函数 。传入引用地址意味着 ，
在过程或函数内部对其修改时 ，也会影响上一级过程 （父过程 ）中的变量的值 。
ByVal 实例
通过以下代码测试 ByVal 类型 ：

Sub Test()
    
    Dim msg As String
    msg = "main"
    
    TestSub1 msg
    
    MsgBox msg
    
End Sub

'ByVal 传递类型
Sub TestSub1(ByVal msg As String)
    msg = "val"
End Sub
首先定义一个 msg 变量 ，赋值 main ，然后调用 TestSub1 过程 ，传入 msg 变量 ，
在过程内部对 msg 重新赋值 Val 。最后返回上一个过程 ，显示 msg 变量 。结果如下 ，
msg 变量的值没有改变 。



ByRef 实例
通过以下代码测试 ByVal 类型 ：

Sub Test()
    
    Dim msg As String
    msg = "main"
    
    TestSub2 msg
    
    MsgBox msg
    
End Sub

'ByRef 传递类型
Sub TestSub2(ByRef msg As String)
    msg = "ref"
End Sub
首先定义一个 msg 变量 ，赋值 main ，然后调用 TestSub2 过程 ，传入 msg 变量 ，
在过程内部对 msg 重新赋值 ref 。最后返回上一个过程 ，显示 msg 变量 。结果如下 ，
msg 变量的值已改变 。



省略传递类型
默认情况下 ，当省略传递类型时 ，默认值是 ByVal ，因此以下两种写法是等效的 。

'指定 ByVal 传递类型
Sub TestSub1(ByVal msg As String)
    
End Sub

'省略传递类型
Sub TestSub1(msg As String)
    
End Sub
使用 ByVal 和 ByRef 传递对象
在上述介绍中说道 ，以上机制适用于传递基础类型变量 ，例如数字 、文本 、逻辑值等 。

使用 ByVal 和 ByRef 传递对象时 ，情况有些不同 。具体用法和不同点将在介绍对象
时详细说明 。

使用 ByVal 和 ByRef 传递数组
过程或函数传递数组时 ，只能以引用形式传递 ，即以 ByRef 形式 。如果尝试用 ByVal
传递数组 ，VBA 会提示错误 。详细的用法将在介绍数组时详细