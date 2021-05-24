getopenfilename不能打开文件 ，不能返回文件夹

Sub f1()
    Dim f
    Dim dig As Object
    Set dig = application.filedialog(msofiledialogopen)
    With application.filedialog(msofiledialogopen)
         .allowmultiselect = True '允许选择多个
         .filters.add "excel文件", "*.xls", 1 'filters过滤，帅选
         .initialfilename = thisworkbook.fullname 'initial最初的，开始的，第一个
         .initialview = msofiledialogviewdetails '显示方式
         .title = "对话框测试"
        '.execute 此方法可以打开文件   execute执行，处死
         .show
        MsgBox .show
        For Each f In  .selecteditems
            MsgBox f
        Next f
    End With
    Set dig = Nothing
End Sub

选择并返回文件夹
Sub f2()
    Dim dig As Object
    Set dig = application.filedialog(msofiledialogfolderpicker)
    'msofiledialogfolderpicker文件夹常量
    With dig
         .initialfilename = "d:\"
         .show
        MsgBox .selecteditems(1)
    End With
    Set dig = Nothing
End Sub