Sub qoptex()
    Dim eapp As Object, i&, j&, sr$, m&, n&, k&, arr_x&, arr_y&, exportpath$
    Dim arr, crr(1 To 1, 1 To 20) '(1 to 2000,1 to 13)
    Dim brr(1 To 300, 1 To 13), d As Object
    Dim currentfilepath As String, sfullpath As String
    Dim pdarr(1 To 4, 1 To 1)
    Dim Process As Object
    
    
    pdarr(1, 1) = ncwx: pdarr(2, 1) = pxwx: pdarr(3, 1) = ycwx: pdarr(4, 1) = xywx
    
    m = MsgBox("提示：此程序运行将关闭已打开excel程序，继续吗？", vbOKCancel, "导出excel表")
    If Not m = 1 Then Exit Sub
    DoCmd.OpenForm "ProgressBar"
    
    For Each Process In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='Excel.EXE'")
        Process.Terminate (0)
    Next
    loadingbar (2)
    
    Set Process = Nothing
    currentfilepath = Application.CurrentProject.Path
    Set eapp = CreateObject("Excel.Application")
    Dim wb As Object
    Dim rst As New ADODB.Recordset
    Set d = CreateObject("scripting.dictionary")
    loadingbar (1)
    
    For m = 1 To 4
        sr = "select * from " & pdarr(m, 1)
        rst.Open sr, CurrentProject.Connection
        loadingbar (3)
        arr = rst.GetRows
        Trans arr
        arr_y = UBound(arr, 2)
        arr_x = UBound(arr)
        k = 1
        For j = 0 To rst.Fields.Count - 2
            brr(k, j + 1) = rst.Fields(j).Name
        Next j
        loadingbar (1)
        
        For i = 0 To arr_x
            d(arr(i, 0)) = i + 1
            For j = 2 To arr_y
                sr = arr(i, 0) & "|" & brr(1, j + 1)
                arr(i, j) = IIf(IsNull(arr(i, j)), 0, arr(i, j))
                If Not d.exists(sr) Then
                    d(sr) = arr(i, j)
                Else
                    d(sr) = d(sr) + arr(i, j)
                End If
                crr(1, j + 1) = crr(1, j + 1) + arr(i, j)
            Next j
        Next i
        loadingbar (2)
        
        For i = 0 To arr_x
            n = d(arr(i, 0)): k = k + 1
            If i + 1 = n Then
                For j = 1 To arr_y
                    brr(k, j) = arr(i, j - 1)
                Next j
                k = k + 1: brr(k, 1) = arr(i, 0) & "小计"
                For j = 3 To arr_y
                    brr(k, j) = d(arr(i, 0) & "|" & brr(1, j))
                Next j
            Else
                For j = 1 To arr_y
                    brr(k, j) = arr(i, j - 1)
                Next j
            End If
        Next i
        k = k + 1: brr(k, 1) = "总计"
        For j = 3 To arr_y
            brr(k, j) = crr(1, j)
            crr(1, j) = 0
        Next j
        loadingbar (4)

        Set arr = Nothing
        sfullpath = currentfilepath & "\模版\" & pdarr(m, 1) & ".xlsx"
        Set wb = eapp.workbooks.Open(sfullpath, False, False)
        eapp.Application.screenupdating = False
        eapp.Visible = False
        loadingbar (1)
        
        With wb.sheets(1)
            .Range(.Cells(1, 1), .Cells(1, arr_y)).merge
            .Range("a1").HorizontalAlignment = 3
            .Range("a2").Resize(k, arr_y) = brr
            .UsedRange.Offset(1).Resize(k).Borders.LineStyle = xlContinuous
            .Columns("A:N").EntireColumn.AutoFit
            .Columns("A:N").Interior.Pattern = 0
        End With
        loadingbar (4)
        'Set arr = Nothing
        exportpath = storgepath & "\" & pdarr(m, 1) & ".xlsx"
        wb.Saveas exportpath
        loadingbar (9)
        wb.Close
        rst.Close
        d.RemoveAll
        
    Next m
    
    
    eapp.Application.screenupdating = True
    Set eapp = Nothing
    Set rst = Nothing
    Set d = Nothing
    Erase brr, crr
    
    DoCmd.Close acForm, "ProgressBar"
    MsgBox "导出完成"
    
End Sub

Public Function Getrst(ByVal strquery As String) As ADODB.Recordset
    Dim objrst As New ADODB.Recordset
    On Error GoTo Error_Getrst
    objrst.Open strquery, CurrentProject.Connection
    Set Getrst = objrst
    Exit Function
Error_Getrst:
    MsgBox (Err.Description)
End Function

Function Trans(ByRef mrr)
    Dim i&, j&, nrr
    On Error Resume Next
    ReDim nrr(LBound(mrr, 2) To UBound(mrr, 2), LBound(mrr, 1) To UBound(mrr, 1))
    If Err.Number <> 0 Then Exit Function
    For i = LBound(mrr, 1) To UBound(mrr, 1)
        For j = LBound(mrr, 2) To UBound(mrr, 2)
            nrr(j, i) = mrr(i, j)
        Next
    Next
    mrr = nrr
End Function

Function loadingbar(ByVal addnum As Integer)
Dim t&
With Forms!ProgressBar
    For t = 1 To addnum
        .LBLCount.Caption = .LBLCount.Caption + 1
        .LBLProgress.Width = .LBLProgress.Width + 50
    Next t
End With
End Function

Sub wx_hz_export()
    Dim sr$
    sr = storgepath
    'sr = CurrentProject.Path
    sr = sr & "\" & "五险单位及个人汇总.xlsx"
    DoCmd.OutputTo acOutputQuery, "五险单位及个人汇总", "ExcelWorkbook(*.xlsx)", sr, False, "", , acExportQualityPrint
    MsgBox "已导出"
End Sub

Sub wx_mx_export()
    Dim sr$, i&, k&, sr1$
    sr = storgepath
    
    Dim arr$(1 To 15, 1 To 1)
    arr(15, 1) = "养老南昌": arr(14, 1) = "医保南昌": arr(13, 1) = "医保萍乡": arr(12, 1) = "医保新余"
    arr(11, 1) = "医保宜春": arr(10, 1) = "失业南昌": arr(9, 1) = "失业宜春": arr(8, 1) = "工伤南昌"
    arr(7, 1) = "工伤萍乡": arr(6, 1) = "工伤宜春": arr(5, 1) = "生育南昌": arr(4, 1) = "大病南昌"
    arr(3, 1) = "大病萍乡": arr(2, 1) = "大病新余": arr(1, 1) = "大病宜春"

    k = MsgBox("是否导出大病明细？", vbOKCancel, "导出五险明细表")
    sr = storgepath
    'vbOK 1 确定、vbCancel 2、vbAbort 3 终止、vbRetry 4、vbIgnore 5 忽略、vbYes 6 是、vbNo 7
    If k = 1 Then
        DoCmd.OpenForm "ProgressBar"
        For i = 1 To 15
            sr1 = sr & "\" & arr(i, 1) & ".xlsx"
            DoCmd.OutputTo acOutputTable, arr(i, 1), "ExcelWorkbook(*.xlsx)", sr1, False, "", , acExportQualityPrint
            loadingbar (6)
        Next i
        loadingbar (10)
        DoCmd.Close acForm, "ProgressBar"
        MsgBox "已导出"
    ElseIf k = 2 Then
        DoCmd.OpenForm "ProgressBar"
        For i = 5 To 15
            sr1 = sr & "\" & arr(i, 1)
            DoCmd.OutputTo acOutputTable, arr(i, 1), "ExcelWorkbook(*.xlsx)", sr1, False, "", , acExportQualityPrint
            loadingbar (10)
        Next i
        DoCmd.Close acForm, "ProgressBar"
        MsgBox "已导出"
    End If
End Sub