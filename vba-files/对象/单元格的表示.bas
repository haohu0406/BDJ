range("a1")
range("a" & 1)
cells(1, 1)
cells(1, "a")
cells(1)
[a1]
区域
range("a1:b5")
range("a1", "c5")
range(cells(1, 1), cells(5, 3))
range("a1:a10").offset(0, 1) 'offset(上下，左右）
range("a1").resize(5, 3).Select 'a1包括自身所在行往下5格，即不含自身往下4格，
列也是一样

range("a1,c1:f4,a7") '不相邻单元格a1区域c1:f4单元格a7
union(range("a1"), range("a1:f4"), range("a7")) '同上
union(rg, cells(x, 1) 'rg必须有初始值
表示行
rows(1)
rows("3:7")
range("c4:f5").entirerow
表示列
columns(1)
columns("a:b")
range("a:b,d:e")
range("c4:f5").entirecolumn 'c到f列

重置坐标系啊的单元格的表示方法
range("b2").range("a1") '原来坐标系初始顶点时A1，现在时b2为顶点的坐标系
'b2变成a1
selection.value = 100 '正在选路的单元格区域=100

Set Rng = Sheets(3).Range(Cells(1 + i, 1), Cells(3 + i, 4))
在你的bai语句du里 ，Range() 的母对象是Sheets(3) ，而Cells() 的母对象是当前活动工作表 ，
当Sheets(3) 不是活动表 ，Cells() 和Range() 的母对象交叉混乱了 ，在VBA里是非法的 。
即前后单元格区域的母对象必须一致 ，可以在后面的Cells() 前面都加一个Sheets(3).