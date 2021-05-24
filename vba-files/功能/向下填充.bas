Sub 向下填充()
    Dim arr, i, j, k
    If Selection.Cells.Count > 1 Then
        arr = Selection.Value '此时arr是一个二维数组，转置后成为1维数组
        'ubound(application.transpose(arr))输出的是一维数组的维度
        For j = 1 To UBound(arr, 2)
            For i = 1 To UBound(arr)
                If arr(i, j) <> "" Then
                    For k = i + 1 To UBound(arr)
                        If arr(k, j) = "" Then
                            arr(k, j) = arr(i, j)
                        Else
                            Exit For
                        End If
                    Next
                End If
            Next
        Next
    Else
        MsgBox "请选择填充区域"
    End If
    Selection = arr
    Erase arr
End Sub