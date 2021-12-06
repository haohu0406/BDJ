Function loadingbar(ByVal addnum As Integer)
Dim t&
With Forms!ProgressBar
    For t = 1 To addnum
        .LBLCount.Caption = .LBLCount.Caption + 1
        .LBLProgress.Width = .LBLProgress.Width + 50
    Next t
End With
End Function