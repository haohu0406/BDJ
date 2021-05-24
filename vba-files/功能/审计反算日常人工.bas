Sub cgsj()
    Dim i, j
    Dim arr, temp
    If Sheet1.Range("c2") = "12220201\内部往来\上级拨入经费\日常经费" Then
    Else
        MsgBox "科目不正确"
        Exit Sub
    End If
    
    arr = Sheet1.Range("c2").CurrentRegion
    For i = 2 To UBound(arr)
        For j = UBound(arr) To 3 Step -1
            If arr(j, 4) < arr(j - 1, 4) Then
                For k = 2 To 12
                    temp = arr(j - 1, k)
                    arr(j - 1, k) = arr(j, k)
                    arr(j, k) = temp
                Next
            End If
        Next
    Next '获得排序完毕的年份
    
    Dim d As Object
    Set d = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arr)
        d(arr(i, 4)) = d(arr(i, 4)) + 1
    Next '获得排序且不重复的年份
    
    With Sheet2
        k =  .Range("A4").End(4).Row - 4
        n = d.Count
        If k < n Then
             .Rows(5).Resize(n - k).Insert
        Else
            If k > n Then
                 .Rows(5).Resize(k - n).EntireRow.Delete
            End If
        End If
         .Range("a4:a" & n + 3) = Application.Transpose(d.keys) '得到列标
        
        Dim m
        For i = 2 To UBound(arr) Step 2
            m = i / 2 + 3
             .Cells(m, 4) = arr(i, 7)
             .Cells(m, 5) = arr(i + 1, 7)
             .Cells(m, 2) = arr(i, 8)
             .Cells(m, 3) = arr(i + 1, 8)
             .Cells(m, 6) =  .Cells(m, 2) -  .Cells(m, 4)
             .Cells(m, 7) =  .Cells(m, 3) -  .Cells(m, 5)
        Next
    End With
End Sub