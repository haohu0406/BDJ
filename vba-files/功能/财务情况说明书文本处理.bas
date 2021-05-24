Sub word()
    Dim tb As Table
    Dim docpath, wbpath, wbname, myname
    Dim ex, wb, ws As Object
    Set ex = CreateObject("excel.application")
    docpath = ActiveDocument.Path & "\"
    wbname = Dir(docpath & "*情况说明书取数表.xls*")
    wbpath = docpath & wbname
    Set wb = ex.Workbooks.Open(wbpath)
    Set ws = wb.sheets("sheet1")
    Dim i, j, m
    With ActiveDocument.Bookmarks("表1").Range.Tables(1)
        m = Split(.Cell(1, 1).Range, Chr(13) & Chr(7))(0)
        If m = "经营情况" Then
            For i = 2 To 6
                 .Cell(i, 2).Range = ws.Cells(i, 7)
                 .Cell(i, 4).Range = ws.Cells(i, 9)
            Next
            For i = 8 To 14
                 .Cell(i, 2).Range = ws.Cells(i, 7)
                 .Cell(i, 4).Range = ws.Cells(i, 9)
            Next
        End If
    End With
    
    Dim MyRange
    Set MyRange = ActiveDocument.Content
    MyRange.Find.ClearFormatting
    MyRange.Find.Replacement.ClearFormatting
    With ws
        MyRange.Find.Execute FindText:="[收到经费拨款]",  _
                ReplaceWith:=.Range("i17").Value, Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[日常经费拨款]",  _
                ReplaceWith:=.Range("i18").Value, Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[专项经费拨款]",  _
                ReplaceWith:=.Range("i19"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[资本性经费]",  _
                ReplaceWith:=.Range("i20"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[累计收到经费拨款]",  _
                ReplaceWith:=.Range("i21"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[本期内部应收款]",  _
                ReplaceWith:=.Range("i22"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[本月内部往来余额]",  _
                ReplaceWith:=.Range("i23"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[上级拨入经费余额]",  _
                ReplaceWith:=.Range("i24"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[内部应收款余额]",  _
                ReplaceWith:=.Range("i25"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[内部应付款余额]",  _
                ReplaceWith:=.Range("i26"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[内部资金往来]",  _
                ReplaceWith:=.Range("i27"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[经费节余]",  _
                ReplaceWith:=.Range("i28"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[奖励基金]",  _
                ReplaceWith:=.Range("i29"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[设备购置基金]",  _
                ReplaceWith:=.Range("i30"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[储备基金]",  _
                ReplaceWith:=.Range("i31"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[调拨固定资产]",  _
                ReplaceWith:=.Range("i32"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[本月其它业务收支]",  _
                ReplaceWith:=.Range("i33"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[本月营业外收入]",  _
                ReplaceWith:=.Range("i34"), Replace:=wdReplaceAll
        MyRange.Find.Execute FindText:="[本月营业外支出]",  _
                ReplaceWith:=.Range("i35"), Replace:=wdReplaceAll
    End With
    myname = Split(ActiveDocument.Name, ".")(0) & Format(Now, "yyyy-mm-dd h：mm：ss") & ".doc"
    ActiveDocument.SaveAs FileName:=docpath & myname
End Sub