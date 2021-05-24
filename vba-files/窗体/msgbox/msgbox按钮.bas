Sub test8()
    MsgBox "test", vbYesNo + vbQuestion
End Sub

vbOKOnly 0 只显示确定按钮
vbOKCancel 1 显示确定及取消按钮
vbAbortRetryIgnore 2 显示放弃 ，重试及忽略按钮
vbYesNoCancel 3 显示是否及取消按钮
vbYesNo 4 显示是否按钮
vbRetryCancel 5 显示重试及取消按钮
vbCritical 16 危险图标
vbQuestion 32 询问图标
vbExclamation 48 警告图标
vbInformation 64 信息图标
vbDefaultButton1 0 第一个按钮是缺省值
vbDefaultButton2 256 第二个按钮是缺省值
vbDefaultButton3 512 第三个按钮是缺省值
vbDefaultButton4 768 第四个按钮是缺省值
vbApplicationModal 0 应用程序强制返回 ，应用程序一直被挂起 ，直到用户响应
vbSystemModal 4096 系统强制返回 ，全部应用该程序被挂起 ，直到用户响应
vbmsgboxhelpbutton 16384 将help按钮添加到消息框
vbmsgboxsetforeground 65536 制定消息框窗口最为前景窗口 ，就是显示在窗口的最上层
vbmsgboxright 524288 文本为右对齐
vbmsgboxrtlreading 1048576 制定文本应为希伯来和阿拉伯语系统中的从右到左显示