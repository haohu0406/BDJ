概念
sheets 所有工作表的集合
sheets("A") 名称为A的工作表
workbook(2), 按打开顺序 ，第二个打开的工作簿
sheets(2).range("a1") = 200 '从左到右第二个工作表
activesheet '当前操作的工作表


1.判断a工作表是否存在
Dim x As Integer
For x = 1 To sheets.count
    If sheets(x).Name = "a" Then
        MsgBox "a工作表存在“
        Exit Sub
    End If
Next
MsgBox "a工作表不存在"
End Sub


2.工作表的插入
Dim sh As worksheet
Set sh = sheets.add
sh.Name = "模板"
sh.range("a1") = 100
Sub 添加工作表()
    For i = 1 To 5
        Worksheets.Add.Name = i
    Next
End Sub



3.工作表的隐藏和取消隐藏
sheets(2).visible = True


4.工作表的移动
sheets("sheet2").move before:=sheets("sheet1") 'sheet2移动到sheet1前面
sheets("sheet1").move after:=sheets.(sheets.count) 'sheet1移动到最后面


5.工作表的复制
sheets("a").copy before:=sheets(1) '复制工作表a到第一个工作表后面
sheets("a").copy '复制工作表到新的工作簿
wb.saveas thiswoekbook.path & "/1日.xls" '保存到当前路径名为1日
wb.close True


6.保护工作表
sheets("sheet2").protect "123"


7.判断工作表是否添加了保护密码
sheets("sheet2").protectcontents = True


8.工作表删除
application.displayalerts = False
sheets("a").delete
application.displayalerts = True


9.工作表的选取
sheets("sheet2").Select

Sub fksqd()
    Dim h, i
    Dim arr
    h = Range("e3").End(xlDown).Row
    arr = Sheet1.Range(Cells(3, 1), Cells(h, 7))
    For i = 1 To UBound(arr)
        Sheet2.Copy before:=Sheet2
        Sheets("模版 (2)").Name = arr(i, 5)
        With Sheets(arr(i, 5)) '工作表的明细用数组表示
             .Cells(6, 21) = arr(i, 2)
             .Cells(7, 3) = arr(i, 3)
             .Cells(7, 21) = arr(i, 4)
             .Cells(11, 2) = arr(i, 7)
             .Cells(11, 30) = arr(i, 6)
        End With
    Next
End Sub
Sub wshzw()
    Dim i As Integer
    For i = 1 To 5
        Sheets("Sheet1").Copy After:=Sheets(1)        
        'Before/After 复制新表在 Sheets("Sheet1") 前/后
        ActiveSheet.Name = i & "月" '为复制的新表命名
    Next i
    Sheets("Sheet1").Name = "总表" '为 Sheets("Sheet1") 改名
End Sub


Sub wshzw()
    Dim i As Integer
    For i = 1 To 5
        Sheets("Sheet1").Copy After:=Sheets(1) '复制新表
        Sheets("Sheet1 (2)").Name = i & "月" '为复制的新表命名
    Next i
    Sheets("Sheet1").Name = "总表" '为 Sheets("Sheet1") 改名
End Sub

