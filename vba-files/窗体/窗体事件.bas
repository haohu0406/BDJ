1.窗体加载事件
'用于窗体出现之前（加载变量等）
Private Sub userform_initialize()
    MsgBox Me.caption
End Sub

2.窗体关闭事件
Private Sub userform_queryclose(cancel As Integer, closemode As Integer)
    If closemode = 0 Then
        cancel = 1 '取消关闭    
    End If    
End Sub
'cancel值为>0时，关闭
'closemode关闭模式，0为点击关闭按钮，1是使用unloadme关闭

3.窗体关闭后事件
Private Sub userform_terminate()
    
End Sub

4.窗体活动和非活动事件
Private Sub userform_deactivate()
    
End Sub

5.窗体的单击和双击事件
Private Sub userform_click()
    
End Sub

6.键盘事件
keydown, keyup, keypress...

7.鼠标事件
