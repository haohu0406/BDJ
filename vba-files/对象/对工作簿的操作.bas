'workbook
'workbooks 泛指excel文件或工作簿
'workbooks（“A.xls”）名称为A的excel工作簿
windows("A.xls").visible = False '名称为A的excel工作簿设置为不可见


1.
Sub w1()
    If Len(Dir("d:/xls")) = 0 Then
        MsgBox ""
    Else
        MsgBox ""
    End If
End Sub
注意 ：
twbpath = thisworkbook.path & "\"
'Application.PathSeparator返回"\",效果同上
bpath = Dir(twbpath & "*.xls")
返回的时表格名字 ，并不带路径



1.判断A.xls文件按是否存在
If Len(Dir("d:/a.xls")) = 0 Then
    MsgBox "A文件不存在"
Else
    MsgBox "A文件存在"
End If
End Sub

2.判断A.xls文件是否打开
For x = 1 To wondows.count
    If windows(x).caption = "a.xls" Then
        MsgBox "a文件打开了"
        Exit Sub
    End If
Next
End Sub

3.excel文件新建和保存
Set wb = workbooks.add
wb.sheets("sheet1").range("a1") = "abcd"
wb.saveas "d:/b.xls"

4.excel文件打开和关闭
Dim wb As workbook
Set wb = workbooks.open("d:/b.xls")
MsgBox wb.sheets("sheet1").range("a1").value
wb.close True '关闭的时候保存
wb.close False '关闭的时候不保存
'注意:wb.sheet1不支持此定义


5.文件保存备份
wb.save
wb.savecopyas "d:abc.xls"

6.文件的复制和删除
FileCopy "d:abc.xls", "e:/abcd.xls"
Kill "d:/abc.xls"
'复制D盘abc文件到E盘命名为abcd
'删除D盘abc的文件

7.获取当前工作簿地址
Dim path As String
path = thisworkbook.path

