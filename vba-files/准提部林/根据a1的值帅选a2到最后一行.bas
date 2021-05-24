Sub worksheet_change(ByVal target As range) '根据a1的值帅选a2到最后一行
    If target.address = "$A$1" Then 'A1写成a1则不相等
        If target = "" Then Me.autofiltermode = False: Exit Sub
        range("A1").currentregion.autofilter field:=1, criteria1:=[a1].value
    End If
End Sub