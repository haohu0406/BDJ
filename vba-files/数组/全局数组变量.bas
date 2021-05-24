'在一个模块的最上端 ：
Public Arr
'此全局数组才可被每个过程识别 ，不必另外在过程中定义arr

'thisWorkbook下定义的PUBLIC是做为Thisworkbook的一个属性来引用的
'在模块里引用时用 ：
ThisWorkbook.Arr

'动态数组传递 ：
'在同一个模块 （模块名为ctt ）中
'模块级声明 （可以加上Option Explicit ）
Public Zn As Integer, Mn As Integer, Alf As Single, Hv As Single, Cv As Single, Xv As Single, Arr
'一个过程
Sub gcurve()
    Narr = 0 '手工修改，看效果
    Start = 0 '手工修改，看效果
    If Start = 0 Then
        Call Point(Narr)
        MsgBox ctt.Arr(1) + ctt.Arr(3)
    Else
        Narr = Narr + 2
        Call Point(Narr)
        MsgBox ctt.Arr(1) + ctt.Arr(3)
    End If
End Sub
'一个子过程
Public Sub Point(n)
    Dim i As Integer
    ReDim Arr(1 To n) As Single
    For i = 1 To n
        Arr(i) = i * 0.1
    Next i
End Sub
'【注意 】
'需要在模块级声明公共变量 Arr
'需要在子过程Point中声明redim Arr()
'在过程gcurve中声明redim Arr() ，而在子过程中不声明redim Arr() ，则msgbox将出问题