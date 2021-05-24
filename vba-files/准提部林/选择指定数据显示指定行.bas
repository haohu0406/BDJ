Private Sub combobox1_change()
    Dim v&, r&
    application.screenupdating = False
    cells.entirerow.hidden = False
    v = Val(combobox1.value)
    If v = 0 Then [a4].Select: Exit Sub
    r = (v - 1) * 110 + 4
    range("a4:a1322").entirerow.hidden = True
    range("a" & r).resize(109).hidden = False
    application.screenupdating = True
    range("a" & r).Select
End Sub