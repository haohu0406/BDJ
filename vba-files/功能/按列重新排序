Public  Sub 按列重排序()
    dim arr,i&,j&,ecol&,d as Object,rng as Range,k&,sr$,rngr&
    set rng=Application.InputBox("请选择标题1、2所在行","标题1、2",,,,,,8).entirerow
    '0公式、1数字、2文本、4逻辑值、8range、16错误值、64数值数组
    
    set d=CreateObject("Scripting.Dictionary")
    With activesheet
        arr=.usedrange
        sr=.usedrange.cells(1,1).entirerow.address
        k=InStr(,sr,":",)
        k=Mid(sr,2,k-1)

        sr=rng.address
        rngr=InStr(,sr,":",)
        rngr=Mid(sr,2,rngr-1)
        
        if rngr>k then 
        i=ubound(arr):j=ubound(arr,2)
        redim preserve arr(1 to i,1 to j+1)
        ecol=j+1
        For j = 1 to ecol-1
            d(arr(rngr,j))=j
        Next j
        For j = 1 to ecol-1
            if arr(2,j)=0 then goto 101
            If d.exists(arr(2,j)) and d(arr(2,j))<>j Then
                For i = rngr+1 to ubound(arr)
                    arr(i,ecol)=arr(i,j)
                    arr(i,j)=arr(i,d(arr(2,j)))
                    arr(i,d(arr(2,j)))=arr(i,ecol)
                Next i
                j=j-1
            End If
        101:Next j
        
    End With
End Sub