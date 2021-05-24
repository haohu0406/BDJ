功能 ：适用于单个工作簿中的多个程序功能
workbook
1.open ;打开事件
Private workbook_open()
userform1.show
End Sub

2.beforeclose '关闭前事件
Private Sub workbook_beforeclose(cancel As Boolean) 'cancel
    cancel = True '禁止关闭
End Sub

3.beforeprint '打印前事件
Private workbook_beforeprint(cancel As Boolean)
cancel = True '禁止打印
End Sub

4.newsheet '工作表插入事件
Private workbook_newsheet(ByVal sh As Object)
sh delete
End Sub
'插入工作表就执行删除

5.beforesave '保存前事件

6.workbook_activate '打开工作簿事件

7.workbook_deactivate '离开工作簿事件（非关闭，含关闭）

8.workbook_sheetselectionchange() '与worksheet_selectionchange功能一样，差别
在于前者适用所有工作表 ，后者只适应单个工作表

9.