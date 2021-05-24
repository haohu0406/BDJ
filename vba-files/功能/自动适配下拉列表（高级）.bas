Private Sub ListBox1_Click()    
    arr = Sheet7.Range("A1").CurrentRegion    
    t = UBound(arr)    
    On Error Resume Next
    k = Application.WorksheetFunction.Match(Me.ListBox1.Value, Sheet7.Range("A1:A" & t), 0)    
    ActiveCell.Value = Me.ListBox1.Value    
    ActiveCell.Offset(0, 1).Value = Application.WorksheetFunction.Index(Sheet7.Range("c:c"), k)    
    Me.TextBox1.Visible = False    
    
End Sub
Private Sub ListBox1 _ Click()
            ActiveCell.Value = Me.ListBoxl.Value
    Me.ListBox1.Visible = False
    Me.TextBox1.Visible = False
End Sub

Private Sub TextBox1_Change() '检测TextBox 中是否有输入
    Dim arr, i%, j%, d
    Set d = CreateObject("scripting.dictionary") '创建字典用于保存搜索到的结果
    arr = Sheet7.Range("A1").CurrentRegion '获取页面内容
    For i = 2 To UBound(arr)
        If InStr(arr(i, 1), Me.TextBox1.Value) Then '遍历数据源，搜索符合条件的用户名
            d(arr(i, 1)) = "" '保存符合条件的数据
            
        End If
    Next
    Me.ListBox1.Clear    
    If d.Count >= 1 Then Me.ListBox1.List = d.keys '输出搜索结果
    
End Sub
Private Sub TextBox1_Change()
    Dim arr, i% ，j% ，d
    Set d = Create0bject(" scripting. dictionary' )
    arr = Sheet2.Range(" A1 ").CurrentRegion
    For i = 2 To UBound(arr)
        If InStr(arr(i, 1) ，Me.TextBox1.Value) Then
            d(arr(i, 1)) = “”
        End If
    Next
    Me.ListBox1.Clear
    If d.Count >= 1 Then Me.ListBox1.List = d.keys
End Sub
    
    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        If Target.Count > 1 Then Me.TextBox1.Visible = False: Me.ListBox1.Visible = False: Exit Sub        
        If Target.Column <> 5 Then Me.TextBox1.Visible = False: Me.ListBox1.Visible = False: Exit Sub
        
        If Target.Row < 2 Then Me.TextBox1.Visible = False: Me.ListBox1.Visible = False: Exit Sub        
        Dim arr, i%, j%, d
        Set d = CreateObject("scripting.dictionary") '获取页面内容
        arr = Sheet7.Range("A1").CurrentRegion '创建字典用于保存搜索到的结果
        For i = 2 To UBound(arr)
            d(arr(i, 1)) = "" '保存符合条件的数据
        Next
        With Me.TextBox1 '显示TextBox
             .Top = Target.Top
             .Left = Target.Left
             .Width = Target.Width
             .Height = Target.Height
             .Activate
             .Value = ""
             .Visible = True
        End With
        With Me.ListBox1 '显示ListBox
             .Clear
             .Top = Target.Offset(1, 1).Top
             .Left = Target.Offset(0, 1).Left
             .Height = Target.Offset(0, 1).Height * 8
             .Width = Target.Offset(0, 1).Width * 4
            
             .List = d.keys
             .Visible = True
        End With
    End Sub