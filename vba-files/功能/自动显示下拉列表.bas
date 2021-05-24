Private Sub worksheet_selectionChange(ByVal Target As Range)
    If Target.count > 1 Then Exit Sub
    If Target.column = 2 And Target = "" Then
        Application.SendKeys "%{down}"
        '这种方式有一个bug就是小键盘灯会熄灭
        'Set x = CreateObject(" wscript.Shell ") '用这种方式
        '就不会有上面的bugx.SendKeys " % {down} "
        'set x = Nothing
    End If
End Sub