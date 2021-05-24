Sub shishi()
    Dim tb As Table
    For Each tb In ActiveDocument.Tables
        MsgBox tb.Cell(1, 1).Range
    Next
End Sub

'去方框内容查询
Sub shishi()
    Dim tb As Table
    Dim m
    For Each tb In ActiveDocument.Tables
        m = Split(tb.Range.Cells(1).Range, Chr(13) & Chr(7))(0)
        MsgBox m
    Next
End Sub

'利用书签定位表格
Sub shuqian()
    Dim tb As table
    activedocument.bookmarks("表1").range.tables(1).Select
End Sub
