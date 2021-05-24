Sub 工作簿汇总()
    ActiveSheet.UsedRange.Offset(1, 0).Clear
    Dim sh As Worksheet
    For Each sh In Worksheets
        If sh.Name <> ActiveSheet.Name Then
            sh.UsedRange.Offset(1, 0).Resize(, 12).Copy
            '只复制现有区域下移一行，其中的12列数据
            [a65535].End(3).Offset(1, 0).PasteSpecial(xlPasteValues)
            '选择性粘贴为数值
        End If
    Next
End Sub