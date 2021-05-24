Sub chazhao()
    Dim MyRange
    Set MyRange = ActiveDocument.Content
    MyRange.Find.ClearFormatting
    MyRange.Find.Replacement.ClearFormatting
    MyRange.Find.Execute FindText:="[收到经费拨款]",  _
            ReplaceWith:="1", Replace:=wdReplaceAll
End Sub
'注意必须定义对象content