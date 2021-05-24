getsavefilename语法
getsavefilename(默认显示的文件名 ，帅选条件 ，多给我帅选类型时显示第几个 ，标题)
'注：该窗口也会有实质性的保存操作，只作为返回文件名的一个途径

Sub t1()
    Dim f
    f = application.getsavefilename("示例.xls", "excel表格,*.xls", , "保存示例")
    MsgBox f
End Sub
