'可以更改标题，内容。添加按钮等
基本语法
MsgBox(窗口中显示的内容 ，按钮和图示类别 ，窗口标题 ，
相关的帮助文件 ，帮助文件上下文的编号 ）

窗口显示的内容




2）换行显示
Chr(10) 可以生成换行符
Chr （13 ）可以生成回车符
VbCrLf 换行符和回车符
vbLf 等同于chr(10)
vbCr 等同于chr(13)
3）表格显示
char(9) 制表符
Sub test4()
    MsgBox "姓名" & Chr(9) & "职业" & Chr(10) & "张三" & Chr(9) & "工程师" _
             & Chr(10) & "于上伟" & Chr(9) & "教师"
End Sub

Sub test5()
    Dim sr, x, y
    For x = 1 To 5
        For y = 1 To 3
            If VBA.IsNumeric(cells(x, y)) Then
                k = 12 - Len(cells(x, y))
            Else
                k = 12 - Len(cells(x, y)) * 2
            End If
            sr = sr & cells(x, y) & Space(k)
        Next y
        sr = sr & cells(x, y) & Chr(10)
    Next x
    MsgBox sr
End Sub

'标题更改
Sub test7()
    MsgBox "核对关系出错", , "系统提示"
End Sub