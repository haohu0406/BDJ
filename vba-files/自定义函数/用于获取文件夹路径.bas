子函数1 ：用于获取文件夹路径
Public Function GetMyPath()
    Dim mypath$
    With Application.FileDialog(msoFileDialogFolderPicker)
         .ButtonName = "Pick Me !"
        If  .Show = True Then
            mypath =  .SelectedItems(1)
        Else: Exit Function
        End If
    End With
    GetMyPath = mypath
End Function