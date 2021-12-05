Public  Sub nxcx()
    Dim arr,brr,i&,j&,d as object,k&,v&
    set d=CreateObject("scripting.dictionary")
    arr=sheet1.usedrange
    redim brr(0,1 to ubound(arr)-1)
    For j = 2 To ubound(arr)
        d(arr(1,j))=j-1
    Next j
    arr=sheet2.usedrange
    For i = 2 To ubound(arr)
        v=arr(i,17):k=d(arr(i,6))
        if v>0 and k>0 then brr(0,k)=brr(0,k)+1
    Next i
    sheet1.range("b2").resize(,ubound(brr,2))=brr
End Sub