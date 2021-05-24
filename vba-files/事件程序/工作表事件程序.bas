工作簿事件 工作簿 （thisworkbook ）
工作表事件 工作表
程序事件 工作簿 （thisworkbook ）或类模块

功能 ：适用于单个工作表中的多个程序功能
事件程序时放在EXCEL工作表中的 ，会在打开工作表自动运行 ，放在模块中不会自动运行
1.selectionchange
Private Sub worksheet_selectionchange(ByVal target As range)
    MsgBox target.address
End Sub
worksheet_selectionchange '对象加动作（事件）
target传递变量
程序中出现target传递你正在选取的单元格

只能选取a1
Private Sub worksheet_selectionchange(ByVal target As range)
    If target.address <> "$a$1" Then
        [a1].Select
    End If
End Sub

2.change事件 '单元格编辑
Private Sub worksheet_change(ByVal target As range)
    MsgBox "你已经改单元格的值"
End Sub
Private Sub worksheet_change(ByVal target As range)
    application.enableevents = False '停止触发下一次事件
    target = target * 2 '没条件的时候会一直*2
    application.enableevents = True
End Sub

3.activate事件
Private Sub worksheet_activate() '工作表激活
    If activesheet.Name = "sheet2" Then
        sheets(1).Select
    End If
End Sub '程序促使无法选中sheet2


4.beforedoubleclick
'双击单元格触发程序运行

5.beforerightclick
'右键单击出发

6.calculate
Private Sub worksheet_calculate() '工作表是否发生了计算
    MsgBox "公式"
End Sub

7.deactivate
Private Sub worksheet_deactivate() '离开工作表时触发
    MsgBox "谢谢使用sheet3"
End Sub

8.followhyperlink
Private Sub worksheet_followhyperlink(ByVal target As hyperlink)
    MsgBox target.address
End Sub
'点开超链接显示超链接地址

9.pivottableupdate '数据透视表更新
Private Sub worksheet_pivottableupdate(ByVal target As pivottable)
    
End Sub