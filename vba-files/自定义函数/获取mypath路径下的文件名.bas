子函数2 ：获取mypath路径下的文件名
Public Function GetFileNameArr(mypath)
    Dim i&, Arr_File
    If mypath <> "" Then
        Set Fso = CreateObject("scripting.filesystemobject")
        Set ff = Fso.getfolder(mypath)
        For Each f In ff.Files
            i = i + 1
        Next f
        ReDim Arr_File(1 To i)
        i = 0
        For Each f In ff.Files
            i = i + 1
            Arr_File(i) = f.Name
        Next f
        GetFileNameArr = Arr_File
    Else
        GetFileNameArr = Array("")
    End If
End Function