
currentregion只的是 "连续单元格组成的矩形区域", 除了边界的单元格,
一般单元格有8个相邻单元格, (下图中红线区域)
usedrange是 "当前工作表已经使用的单元格组成的矩形区域", 设置格式也属于
已经使用(下图中的兰线区域) 这两个区域有时相同, 有时不同, 本图中, 二者结果
不同的原因在于黄色区域是空白的

currentregion和usedregion

[a1].currentregion
sheet1.usedrange



想不到UsedRange还可以这样用 ，又学到了 ！
有了这个就可以轻松取得当前Sheet的最末行和最末列号了 ：
Sub test()
    Dim myRange As String
    myRange = ActiveSheet.UsedRange.Address
    Debug.Print "LastRow=" & Cells.SpecialCells(xlCellTypeLastCell).Row
    Debug.Print "LastColumn=" & Cells.SpecialCells(xlCellTypeLastCell).Column
    myRange = ""
End Sub
