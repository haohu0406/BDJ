Sub 调整列宽()
    Dim i%    
    For i = 1 To Sheets.Count '遍历工作簿中所有的工作表
        Sheets(i).Columns("A:K").AutoFit '把每个工作表的[A:K]列调整为最佳列宽
    Next i    
End Sub