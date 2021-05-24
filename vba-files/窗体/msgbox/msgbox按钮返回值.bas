Sub test11()
    k = MsgBox("测试返回值", vbYesNoCancel)
    MsgBox "你点击了按钮:" & Choose(k, "确定", "取消", "终止", "重试", "忽略", "是", "否")
End Sub


'如果要返回值必须给msgbox后的代码加（）
Sub test12()
    If MsgBox("你确定要删除第15行吗", vbQuestion + vbYesNo, "删除提示") = vbYes Then
        rows(15).delete
        MsgBox "删除成功”
    Else
        MsgBox "你取消了删除"
    End If
End Sub

常数 值 说明
vbOK 1 确定
vbCancel 2 取消
vbAbort 3 终止
vbRetry 4 重试
vbIgnore 5 忽略
vbYes 6 是
vbNo 7 否

Sub test13()
    Dim x
    x = MsgBox("测试添加帮助的效果", vbOKCancel + vbmsgboxhelpbutton, "测试帮助!", "d:/a.chm", 0)
End Sub