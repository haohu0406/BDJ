Dim x, y
'定义全局变量
Private Sub worksheet_selectionchange(ByVal target As range)
    y = target.value '改变前的值
End Sub

Private Sub worksheet_change(ByVal target As range)
    Dim ro, cl    
    ro = target.row
    cl = target.column
    x = target.value '改变后的值
    If ro > 3 And cl > 1 And cl < 9 Then
        If y <> x Then
            range("i" & ro) = Format(Now, "yyyy-mm-dd h:mm:ss")
        End If
    End If
End Sub