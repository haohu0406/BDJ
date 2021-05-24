Dim k

Private Sub CommandButton1_Click()
    k = 0
    For i = 1 To 100000000
        DoEvents
        If k = 1 Then Exit For
    Next i
    MsgBox i
End Sub

Private Sub CommandButton2_Click()
    k = 1
End Sub
'使用了DoEvents ，按钮1运行后 ，可用按钮2方便地使循环中途停止 。
Application.EnableEvents = False '避免引起其他事件 
Application.EnableEvents = True '可触发其他事件