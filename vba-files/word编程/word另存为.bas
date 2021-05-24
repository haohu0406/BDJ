Sub word2()
    Dim docpath, myname, wbname
    docpath = ActiveDocument.Path & "\"
    wbname = Dir(docpath & "*情况说明书取数表.xls*")
    myname = Split(ActiveDocument.Name, ".")(0) & Format(Now, "yyyy-mm-dd h：mm：ss") & ".doc"
    ActiveDocument.SaveAs FileName:=docpath & myname
End Sub
