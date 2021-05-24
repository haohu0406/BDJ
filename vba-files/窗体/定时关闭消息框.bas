Sub aa()
    Dim wshshell As Object
    Set wshshell = CreateObject("wscript.shell")
    wshshell.popup "1秒后关闭", 1, "提示", 64
End Sub
