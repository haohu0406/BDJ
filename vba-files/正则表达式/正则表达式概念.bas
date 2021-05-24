对比字符串 语法:
字符串1 like 字符串2

Sub t1()
    debug.Print "abc" like "Abc"
End Sub
'返回false

'通配符？代表1个字符串
Sub t2()
    debug.Print "BA" like "?A"
End Sub

正则表达式
1.完成复杂的字符串判断
2.在字符串判断时 ，可以最大限度避开循环 ，从而提高运行效率

使用方法

1、引用法
2、直接他建法
