单击第一个单元格双击第二个单元格互换内容
Public OldRow As Integer, R As Integer
Public OldColumn As Integer, C As Integer
Public OldValue As Variant, aav As Variant

Private Sub Worksheet_Activate()
    R = ActiveCell.Row
    C = ActiveCell.Column
    aav = ActiveCell.Value
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean) '双击
    'Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)'右键
    Cancel = True '完成动作后返回
    If Target.Count = 1 Then
        Cells(OldRow, OldColumn) = Target.Value
        Target.Value = OldValue
        R = Target.Row
        C = Target.Column
        aav = Target.Value
        
        MsgBox "成功调换！"
    End If
End Sub

Sub Worksheet_SelectionChange(ByVal Target As Range)
    'On Error Resume Next
    OldRow = R
    OldColumn = C
    OldValue = aav
    If Target.Count = 1 Then
        R = Target.Row
        C = Target.Column
        aav = Target.Value
    End If
End Sub
